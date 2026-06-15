# pmj-keieidangi

経営談義 AI財務分析 Cloud Run サービス (Python / Flask)。
報告書(財務数値)JSON を受け取り、Claude / Gemini / OpenAI で分析し、
7セクション(報告書＋販売/収支/資金の課題・提案)を JSON で返す(非ストリーミング)。

通常は `pmj-door` 経由で呼ばれる(直接公開しない想定)。

## エンドポイント
| メソッド | パス | 説明 |
|---|---|---|
| GET  | `/`        | ヘルスチェック |
| POST | `/analyze` | 分析実行 |

### POST /analyze (リクエスト)
```json
{
  "report":   { "store_info": {...}, "financials": { "headers": {...}, "rows": [...] } },
  "tone":     "expert",
  "providers": ["claude", "gemini", "openai"]
}
```
`report` は zaiTask 側 `keieidangi_get_report_data` が返す形式(画面の手入力反映後)。

### レスポンス
```json
{
  "status": "OK",
  "tone": "expert",
  "providers": ["claude","gemini","openai"],
  "results": {
    "claude": { "sections": { "REPORT": "...", "SALES_ISSUE": "...", ... } },
    "openai": { "error": "..." }
  }
}
```

## 必要な環境変数 (Cloud Run に宣言)
| 変数 | 必須 | 説明 |
|---|---|---|
| `ANTHROPIC_API_KEY` | Claude使用時 | Anthropic APIキー |
| `OPENAI_API_KEY`    | OpenAI使用時 | OpenAI APIキー |
| `GOOGLE_API_KEY`    | Gemini使用時 | Google(Gemini) APIキー |
| `ANTHROPIC_MODEL`   | 任意 | 既定 `claude-sonnet-4-5` |
| `OPENAI_MODEL`      | 任意 | 既定 `gpt-4o-mini` |
| `GEMINI_MODEL`      | 任意 | 既定 `gemini-2.5-flash` |
| `PORT`              | 自動 | Cloud Run が自動設定(コード側は対応済) |

※ キーは Secret Manager 連携(`--set-secrets`)推奨。

## ローカル実行
```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=... OPENAI_API_KEY=... GOOGLE_API_KEY=...
python main.py   # http://localhost:8080
```

## デプロイ
Cloud Run の「リポジトリから継続的にデプロイ」でこのリポジトリを連携。
buildpacks が `Procfile`(gunicorn) を検出して起動する。
認証は原則 private(未認証呼び出しを許可しない)。呼び出し元(`pmj-door` のサービスアカウント)に
`roles/run.invoker` を付与する。
