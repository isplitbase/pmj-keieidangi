# -*- coding: utf-8 -*-
"""
pmj-keieidangi : 経営談義 AI財務分析 Cloud Run サービス (Flask)

エンドポイント:
  GET  /         : ヘルスチェック
  POST /analyze  : 財務データをAI分析して返す(非ストリーミング)
      body(JSON): { "report": {...報告書JSON...}, "tone": "expert|plain",
                    "providers": ["claude","gemini","openai"] }
      返り値    : { "status":"OK", "tone":..., "providers":[...],
                    "results": { provider: {"sections":{...}} | {"error":...} } }

APIキーは環境変数で渡す(Cloud Runの環境変数/Secret):
  ANTHROPIC_API_KEY, OPENAI_API_KEY, GOOGLE_API_KEY
  任意: ANTHROPIC_MODEL, OPENAI_MODEL, GEMINI_MODEL
"""
from __future__ import annotations
import os, re, threading
from flask import Flask, request, jsonify

app = Flask(__name__)

# ---- トーン --------------------------------------------------------------
TONE_INSTRUCTIONS = {
    "expert": (
        "【表現: 財務・会計の専門用語を積極的に使用】\n"
        "経理部門・金融機関担当者・税理士などが読む前提で、財務会計の正式な専門用語を正しく使ってください。\n"
        "■推奨用語例: 『売上総利益率(粗利率)』『営業利益率』『売上債権回転日数(DSO)』『棚卸資産回転日数』\n"
        "　『買入債務回転日数』『自己資本比率』『固定長期適合率』『流動比率』『当座比率』『EBITDA』\n"
        "　『キャッシュ・コンバージョン・サイクル(CCC)』『損益分岐点売上高(BEP)』『営業レバレッジ』\n"
        "　『運転資本』『インタレスト・カバレッジ・レシオ』『ROA』『ROE』『販管費率』 等を状況に応じて使用。\n"
        "■数値表現は専門家の書式で。例: 『粗利率14.31%(前期比▲2.73pt)』『DSO 20.8日(+3.1日)』。\n"
        "■省略せずに正確な勘定科目名・指標名を用いる。略語は初出時に日本語名を併記。\n"
        "■文体: 常体(である調)で簡潔に。一般人向けの補足説明は不要。"
    ),
    "plain": (
        "【表現: 財務・会計の専門用語を極力使用しない】\n"
        "財務知識のない中小企業の経営者・店長が読んで直感的に理解できる、平易な日本語で書いてください。\n"
        "■専門用語は原則として使わない。やむを得ず使う場合は、必ず括弧書きで平易な言い換えを添える。\n"
        "　例: 『売掛金（商品を売ったがまだ受け取っていないお金）』『自己資本比率（会社の体力を示す割合）』。\n"
        "■避ける語: 『DSO』『EBITDA』『CCC』『営業レバレッジ』等の略語・カタカナ財務用語は使用禁止。\n"
        "　代わりに『売掛金の回収までの日数』『本業の稼ぐ力』等の日常語に。\n"
        "■数値表現は日常感覚で。例: 『現金が約4日分しか手元にない』『前の年より利益が約1,300万円減った』。\n"
        "■文体: 敬体(です・ます調)でやさしく。難しい表現は避け、1文を短めに。比喩や身近な例えを活用。"
    ),
}

SECTION_NAMES = ["REPORT", "SALES_ISSUE", "SALES_PROPOSAL",
                 "INCOME_ISSUE", "INCOME_PROPOSAL", "CAPITAL_ISSUE", "CAPITAL_PROPOSAL"]
MARKER_RE = re.compile(r"===\s*(" + "|".join(SECTION_NAMES) + r")\s*===")

# ---- プロンプト ----------------------------------------------------------
def _format_financial(data):
    si = (data or {}).get("store_info", {}) or {}
    fin = (data or {}).get("financials", {}) or {}
    h = fin.get("headers", {}) or {}
    yc, yp, yp2 = h.get("year_current"), h.get("year_previous"), h.get("year_previous2")
    lines = [
        "店コード: %s" % si.get("store_code"),
        "店名: %s" % si.get("store_name"),
        "営業所: %s" % si.get("office"),
        "担当者: %s" % si.get("person_in_charge"),
        "",
        "対象年度: %s年 / 前年: %s年 / 前前年: %s年" % (yc, yp, yp2),
        "",
        "【財務項目】",
    ]
    def fmt(v, u):
        if v is None or v == "":
            return "-"
        unit = u or ""
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return "{:,.2f}{}".format(v, unit)
        return "{}{}".format(v, unit)
    for r in fin.get("rows", []):
        lines.append(
            "[{}] {} {}: {}={}, {}={}, 前期差={}, {}={}, 前前期差={}".format(
                r.get("category") or "", r.get("no") or "", r.get("item") or "",
                yc, fmt(r.get("v_curr"), r.get("u_curr")),
                yp, fmt(r.get("v_prev"), r.get("u_prev")),
                fmt(r.get("diff_prev"), r.get("u_diff")),
                yp2, fmt(r.get("v_prev2"), r.get("u_prev2")),
                fmt(r.get("diff_prev2"), r.get("u_diff2")),
            )
        )
    return "\n".join(lines)

def build_prompt(data, tone):
    aliases = {"mild": "plain", "normal": "expert", "strict": "expert"}
    tone = aliases.get(tone, tone)
    tone = tone if tone in TONE_INSTRUCTIONS else "expert"
    body = _format_financial(data)
    return (
        "あなたは中小企業の経営コンサルタントです。下記の財務データを分析し、\n"
        "『経営談義報告書』と、販売・収支・資金それぞれの『課題』と『提案』を日本語で作成してください。\n\n"
        "【表現ルール — 100%遵守すること】\n"
        + TONE_INSTRUCTIONS[tone] + "\n\n"
        "■上記の表現ルールは出力全体（REPORT/各課題/各提案すべて）に適用すること。\n\n"
        "【出力フォーマット (厳守)】\n"
        "必ず下記のマーカー形式で出力してください。余計な前置き・挨拶は不要です。\n"
        "マーカーは半角英大文字のみ。課題と提案は明確に分けて書くこと。\n\n"
        "===REPORT===\n(経営談義報告書として、全体状況・キーメッセージを **300文字以内** で要約)\n\n"
        "===SALES_ISSUE===\n(販売の課題。箇条書き3〜5点。販売高/月商/ペイライン/従業員生産性の観点)\n\n"
        "===SALES_PROPOSAL===\n(販売の提案。課題に対応する具体的な改善アクション。箇条書き3〜5点)\n\n"
        "===INCOME_ISSUE===\n(収支の課題。箇条書き3〜5点。粗利率・管理経費率・営業利益率・支払利息に言及)\n\n"
        "===INCOME_PROPOSAL===\n(収支の提案。課題に対応する改善アクション。箇条書き3〜5点)\n\n"
        "===CAPITAL_ISSUE===\n(資金の課題。箇条書き3〜5点。現預金日数・売掛/棚卸/借入日数・自己資本に言及)\n\n"
        "===CAPITAL_PROPOSAL===\n(資金の提案。課題に対応する改善アクション。箇条書き3〜5点)\n\n"
        "【財務データ】\n" + body + "\n"
    )

def split_sections(text):
    out = {}
    if not text:
        return out
    matches = list(MARKER_RE.finditer(text))
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        out[m.group(1)] = text[start:end].strip()
    if not matches:
        out["REPORT"] = text.strip()
    return out

# ---- 各プロバイダ (非ストリーミング) ------------------------------------
def call_claude(prompt):
    import anthropic
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        raise RuntimeError("ANTHROPIC_API_KEY 未設定")
    model = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-5")
    client = anthropic.Anthropic(api_key=key)
    msg = client.messages.create(
        model=model, max_tokens=3000,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(getattr(b, "text", "") for b in msg.content
                   if getattr(b, "type", None) == "text")

def call_openai(prompt):
    from openai import OpenAI
    key = os.environ.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("OPENAI_API_KEY 未設定")
    model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
    client = OpenAI(api_key=key)
    r = client.chat.completions.create(
        model=model, temperature=0.3,
        messages=[{"role": "user", "content": prompt}],
    )
    return r.choices[0].message.content or ""

def call_gemini(prompt):
    from google import genai
    key = os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
    if not key:
        raise RuntimeError("GOOGLE_API_KEY 未設定")
    model = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
    client = genai.Client(api_key=key)
    r = client.models.generate_content(model=model, contents=prompt)
    return getattr(r, "text", "") or ""

PROVIDERS = {"claude": call_claude, "openai": call_openai, "gemini": call_gemini}

def _run(prov, prompt, results):
    try:
        results[prov] = {"sections": split_sections(PROVIDERS[prov](prompt))}
    except Exception as e:
        results[prov] = {"error": str(e)[:300]}

def analyze(data, tone, providers):
    prompt = build_prompt(data, tone)
    results = {}
    ts = []
    for p in providers:
        if p in PROVIDERS:
            t = threading.Thread(target=_run, args=(p, prompt, results), daemon=True)
            t.start(); ts.append(t)
    for t in ts:
        t.join(timeout=180)
    return results

# ---- ルーティング --------------------------------------------------------
@app.get("/")
def health():
    return jsonify({"status": "ok", "service": "pmj-keieidangi"})

@app.post("/analyze")
def analyze_route():
    body = request.get_json(silent=True) or {}
    data = body.get("report") or {}
    tone = body.get("tone", "expert")
    if tone not in TONE_INSTRUCTIONS:
        tone = "expert"
    raw = body.get("providers")
    if isinstance(raw, str):
        raw = [p.strip() for p in raw.split(",")]
    providers = [p for p in (raw or ["claude", "openai", "gemini"]) if p in PROVIDERS]
    if not providers:
        providers = ["claude", "openai", "gemini"]
    if not data.get("financials"):
        return jsonify({"status": "NG", "error": "report.financials がありません"}), 400
    results = analyze(data, tone, providers)
    return jsonify({"status": "OK", "tone": tone, "providers": providers, "results": results})

# ---- AI自動要約 (複数AIの分析を比較・検証) -----------------------------
PROVIDER_LABEL = {"claude": "Claude", "gemini": "Gemini", "openai": "OpenAI"}
SEC_LABELS = {
    "REPORT": "経営談義報告書",
    "SALES_ISSUE": "販売-課題", "SALES_PROPOSAL": "販売-提案",
    "INCOME_ISSUE": "収支-課題", "INCOME_PROPOSAL": "収支-提案",
    "CAPITAL_ISSUE": "資金-課題", "CAPITAL_PROPOSAL": "資金-提案",
}
SUMMARY_ORDER = ["claude", "openai", "gemini"]  # 要約担当の優先順

def _has_key(p):
    if p == "claude":
        return bool(os.environ.get("ANTHROPIC_API_KEY"))
    if p == "openai":
        return bool(os.environ.get("OPENAI_API_KEY"))
    if p == "gemini":
        return bool(os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY"))
    return False

def _format_results(results):
    """各AIの分析テキストを読みやすいブロックに整形(存在するAIのみ)。"""
    blocks = []
    for p in SUMMARY_ORDER:
        r = (results or {}).get(p)
        if not r:
            continue
        if r.get("error"):
            blocks.append("=== %s の分析: エラー(%s) ===" % (PROVIDER_LABEL.get(p, p), r.get("error")))
            continue
        sec = r.get("sections") or {}
        lines = ["=== %s の分析 ===" % PROVIDER_LABEL.get(p, p)]
        for key in SECTION_NAMES:
            v = sec.get(key)
            if v:
                lines.append("【%s】\n%s" % (SEC_LABELS.get(key, key), v))
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks)

def build_summary_prompt(data, results, tone):
    aliases = {"mild": "plain", "normal": "expert", "strict": "expert"}
    tone = aliases.get(tone, tone)
    tone = tone if tone in TONE_INSTRUCTIONS else "expert"
    fin = _format_financial(data)
    analyses = _format_results(results)
    return (
        "あなたは複数の生成AIが作成した経営分析をレビューする上級経営コンサルタントです。\n"
        "同じ財務データに対して各AIが作成した『経営分析』を比較・検証し、日本語で下記2部構成でまとめてください。\n"
        "（分析が存在するAIのみを対象にすること。AIは1〜3個のいずれの場合もある）\n\n"
        "【表現ルール — 100%遵守すること】\n"
        + TONE_INSTRUCTIONS[tone] + "\n\n"
        "【出力フォーマット (厳守)】\n"
        "■総合要約\n"
        "(各AIの分析を踏まえた、この企業の財務状況・経営課題の総合的な要約。重要点を簡潔に。)\n\n"
        "■各AIの結論の問題点\n"
        "(分析があるAIごとに、結論や指摘の誤り・根拠の薄さ・数値の誤読・見落とし・AI間の矛盾などを箇条書きで指摘。\n"
        " 問題が無ければ『特になし』。例:\n"
        " ・Claude: ...\n ・Gemini: ...\n ・OpenAI: ...)\n\n"
        "【財務データ】\n" + fin + "\n\n"
        "【各AIの分析】\n" + analyses + "\n"
    )

@app.post("/summarize")
def summarize_route():
    body = request.get_json(silent=True) or {}
    data = body.get("report") or {}
    results = body.get("results") or {}
    tone = body.get("tone", "expert")
    if tone not in TONE_INSTRUCTIONS:
        tone = "expert"
    if not results:
        return jsonify({"status": "NG", "error": "results(各AIの分析)がありません"}), 400

    # 要約担当AIを選定: SUMMARY_PROVIDER 指定 → 無ければキーのある先頭
    pref = (os.environ.get("SUMMARY_PROVIDER", "") or "").strip().lower()
    order = ([pref] if pref in PROVIDERS else []) + [p for p in SUMMARY_ORDER if p != pref]
    chosen = None
    for p in order:
        if _has_key(p):
            chosen = p
            break
    if not chosen:
        return jsonify({"status": "NG", "error": "要約用のAPIキーが未設定です"}), 200

    prompt = build_summary_prompt(data, results, tone)
    try:
        text = PROVIDERS[chosen](prompt)
    except Exception as e:
        return jsonify({"status": "NG", "error": str(e)[:300], "provider": chosen}), 200
    return jsonify({"status": "OK", "provider": chosen, "summary": text})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8080")))
