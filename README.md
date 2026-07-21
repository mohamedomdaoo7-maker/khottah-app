# Khottah — Deployment Guide

This project deploys to **either Vercel or Netlify** — both are supported
out of the box, no code changes needed. Pick whichever you prefer; the
instructions below cover both.

## What you need before starting

1. An **Anthropic API key** — create one at https://console.anthropic.com
   (Settings → API Keys). This is separate from your claude.ai login.
2. A **GitHub account** (free) — to hold the code.
3. A hosting account — either **Vercel** (https://vercel.com) or
   **Netlify** (https://netlify.app), both free tier.
4. A **domain name** — either one you already own, or bought through any
   registrar (Namecheap, GoDaddy, or the host's own domain purchase).

## Step 1 — Put the code on GitHub

1. Create a new empty repository on GitHub (e.g. `khottah`).
2. From this folder, run:
   ```
   git init
   git add .
   git commit -m "Khottah initial deployment"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/khottah.git
   git push -u origin main
   ```

## Step 2A — Deploy to Netlify

1. Go to https://app.netlify.com and sign in (GitHub login is easiest).
2. Click **Add new site → Import an existing project**, then pick your
   `khottah` repository.
3. Netlify reads `netlify.toml` automatically — build command and publish
   folder are already set. Leave the defaults.
4. Before deploying, go to **Site configuration → Environment variables**
   and add:
   - Key: `ANTHROPIC_API_KEY`
   - Value: your actual key from Anthropic Console
5. Click **Deploy site**. You'll get a live URL like
   `khottah-xyz123.netlify.app` in about a minute — test it there first.
6. **Custom domain**: Site configuration → Domain management → Add a
   domain. Netlify shows the DNS records to add at your registrar
   (or use Netlify DNS directly if you transfer the domain to them).
   HTTPS is issued automatically once DNS resolves.

## Step 2B — Deploy to Vercel (alternative)

1. Go to https://vercel.com/new and sign in (GitHub login is easiest).
2. Click **Import** next to your `khottah` repository.
3. Vercel auto-detects the Vite framework — leave the default build
   settings as they are.
4. **Before clicking Deploy**, open **Environment Variables** and add:
   - Name: `ANTHROPIC_API_KEY`
   - Value: your actual key from Anthropic Console
   - Environment: Production, Preview, and Development (select all)
5. Click **Deploy**. In about a minute you'll get a live URL like
   `khottah-xyz123.vercel.app` — test it there first.
6. **Custom domain**: Settings → Domains → add your domain, then add the
   DNS record it shows you at your registrar.

> Deploying to both platforms at once is harmless but unnecessary — pick
> one. If you later want to switch, no code changes are required either
> way; both `api/generate.js` (Vercel) and `netlify/functions/generate.js`
> (Netlify) already exist in this project.

## Local development (optional)

```
npm install
cp .env.example .env        # then paste your real API key into .env

# For Netlify:
netlify dev                 # requires: npm i -g netlify-cli

# For Vercel:
vercel dev                  # requires: npm i -g vercel
```

Either command runs the frontend and the matching serverless function
together, mirroring production. Plain `npm run dev` runs the frontend only
— API calls will fail without one of the two commands above.

## Notes on cost and access

- **"Anyone can log in"**: this deployment is a public website with **no
  login screen** — anyone with the URL can use it, the same as the current
  claude.ai artifact link. If you want to restrict it to your team:
  Netlify offers **Password Protection** and **Netlify Identity** on paid
  tiers; Vercel offers **Deployment Protection** similarly. Or ask me to
  add a simple login screen directly into the app — works on either host.
- **Cost**: Vercel's free tier comfortably covers a pilot. The only real
  running cost is Anthropic API usage — a few hundred SAR per month even
  at moderate use, billed directly by Anthropic on your API key.
- **Rotate the key** immediately if it is ever exposed (e.g. accidentally
  committed to Git) — regenerate it in Anthropic Console.
