# Citation Builder

Marine Corps citation and award writing aid built with Vite, React, Express, and Cloudflare Pages Functions.

## Local setup

```bash
npm install
npm run build
npm start
```

The app runs on `http://localhost:3000` by default. Set `PORT` to use a different port.

## Environment variables

Copy `.env.example` to `.env` for local development and add your real values:

```bash
ANTHROPIC_API_KEY=your_anthropic_api_key_here
ANTHROPIC_MODEL=claude-sonnet-4-6
ANTHROPIC_FALLBACK_MODEL=claude-haiku-4-5-20251001
```

Do not commit `.env`. It contains secrets and is ignored by Git.

## Render deployment

Create a new Render Web Service from this GitHub repository.

- Runtime: `Node`
- Build command: `npm install && npm run build`
- Start command: `npm start`
- Environment variable: `ANTHROPIC_API_KEY`
- Optional environment variable: `ANTHROPIC_MODEL`

This repository also includes `render.yaml`, so you can use Render's blueprint flow if preferred.

## Cloudflare Pages deployment

Cloudflare Pages is the preferred free hosting target because the static app and `/api/*` Pages Functions do not run on a sleeping Node web service.

Create a new Cloudflare Pages project from this GitHub repository:

- Framework preset: `Vite`
- Build command: `npm run build`
- Build output directory: `dist`
- Root directory: leave blank unless this app is inside a monorepo path in Cloudflare
- Production branch: `main`

Set these production environment variables in Cloudflare Pages:

```bash
ANTHROPIC_API_KEY=your_anthropic_api_key_here
ANTHROPIC_MODEL=claude-sonnet-4-6
ANTHROPIC_FALLBACK_MODEL=claude-haiku-4-5-20251001
```

The app includes Cloudflare Pages Functions for:

- `GET /api/health`
- `POST /api/improve`

For local Cloudflare-style testing, create an ignored `.dev.vars` file with the same variables, then run:

```bash
npm run pages:dev
```
