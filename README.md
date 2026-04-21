# Revenue Leak Audit PPTX Service

FastAPI service that generates audit decks using `slide_templates.py`.
Designed to run on Railway; called from a Supabase edge function.

## Files

- `main.py` — FastAPI app with `/generate` endpoint
- `slide_templates.py` — your existing template (drop in the current version)
- `requirements.txt` — Python dependencies
- `railway.json` — Railway deployment config

## Deployment to Railway

1. **Create a new Railway project:**
   - Go to [railway.app/new](https://railway.app/new)
   - "Deploy from GitHub repo" → select this repo (after you push it)
   - Or use the Railway CLI: `railway init` → `railway up`

2. **Set environment variables** in Railway dashboard → Variables:
   - `SERVICE_SECRET` — a long random string (generate with `openssl rand -hex 32`). The Supabase edge function will pass this as the `x-service-secret` header.

3. **Railway auto-deploys.** Once live, you'll get a URL like `https://your-app-name.up.railway.app`. Copy it.

4. **Test it:**
   ```bash
   curl https://your-app-name.up.railway.app/healthz
   # → {"status":"ok","service":"pptx-generator"}
   ```

5. **Test generation** with a sample payload:
   ```bash
   curl -X POST https://your-app-name.up.railway.app/generate \
     -H "Content-Type: application/json" \
     -H "x-service-secret: your-secret-here" \
     -d @sample_data_dict.json \
     --output test_deck.pptx
   ```

## API

### `POST /generate`

**Headers:**
- `x-service-secret: {SERVICE_SECRET}` (required)
- `Content-Type: application/json`

**Body:**
```json
{
  "data_dict": { ... full data dict per DATA_SCHEMA ... },
  "audit_id": "optional-string-for-logging"
}
```

**Response:**
- `200 OK` — raw `.pptx` bytes with MIME type `application/vnd.openxmlformats-officedocument.presentationml.presentation`
- `400` — bad data_dict (missing keys, failed arithmetic check)
- `401` — missing or wrong `x-service-secret`
- `500` — generation error

**Response headers include:**
- `X-Generation-Duration-Ms` — how long generation took
- `X-Pptx-Size-Bytes` — size of generated file

## Arithmetic safety

Before generation, the service verifies `data_dict.guarantee.surplus_amount > 0`.
If not, returns 400 with a clear error message — so bad math never reaches a client deck.

## Cost

Railway's Hobby plan is $5/month flat. This service idles at near-zero CPU and spins
up to handle each request in 2-5 seconds. No cold starts on the paid tier.
