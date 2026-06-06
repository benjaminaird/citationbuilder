# Citation Builder

Marine Corps citation and award writing aid built with Vite, React, and Express.

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
ANTHROPIC_MODEL=claude-sonnet-4-20250514
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
