# ArgusRecruit — Astro + Sanity + Cloudflare Pages

This is the new architecture for argusrecruit.com:
- **Astro** (static site generator) — preserves your existing design pixel-for-pixel
- **Sanity CMS** — web-based admin for writing blog posts (no code, just a browser)
- **Cloudflare Pages** — auto-deploys on every git push
- **Web3Forms** — contact form with file attachments (already configured)

## Folder layout

```
argusrecruit/                  ← Astro project (deploy this)
├── public/                    ← static assets (images, logo, world-map.svg, robots.txt)
├── src/
│   ├── components/            ← shared Astro components (SiteNav, SiteHead, SharedCSS)
│   ├── lib/sanity.js          ← Sanity client + fallback posts
│   └── pages/
│       ├── index.astro        ← English home
│       ├── ru/index.astro     ← Russian home
│       ├── hy/index.astro     ← Armenian home
│       ├── blog/              ← English blog (index + [slug])
│       ├── ru/blog/
│       └── hy/blog/
├── astro.config.mjs
└── package.json

sanity-studio/                 ← Sanity CMS Studio (deploy separately)
├── sanity.config.ts
└── schemas/
    ├── index.ts
    └── post.ts
```

## ⚡ Step-by-step deployment

### 1. Push to GitHub (5 min)

```bash
# in argusrecruit-astro/
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin git@github.com:YOUR_USERNAME/argusrecruit.git
git push -u origin main
```

### 2. Create Sanity project (10 min)

```bash
# in sanity-studio/
npm install
npx sanity@latest login
npx sanity@latest init --create-project "ArgusRecruit" --dataset production --output-path .
# Choose: schema template = Clean project (skip — we already have one)
```

This gives you a **PROJECT ID** like `abc12def`. Copy it.

Then:
```bash
echo "SANITY_STUDIO_PROJECT_ID=YOUR_PROJECT_ID" > .env
npm run dev        # → http://localhost:3333 to test
npm run deploy     # → https://argusrecruit.sanity.studio
```

### 3. Connect Astro to Sanity (2 min)

In `argusrecruit-astro/`:
```bash
cp .env.example .env
# edit .env and set PUBLIC_SANITY_PROJECT_ID=YOUR_PROJECT_ID
```

### 4. Deploy on Cloudflare Pages (15 min)

1. Go to https://dash.cloudflare.com → Workers & Pages → Create → Pages
2. Connect to Git → select your `argusrecruit` GitHub repo
3. Build settings:
   - **Framework preset**: Astro
   - **Build command**: `npm run build`
   - **Build output directory**: `dist`
   - **Environment variables**:
     - `PUBLIC_SANITY_PROJECT_ID` = your Sanity project ID
     - `PUBLIC_SANITY_DATASET` = `production`
4. Save and Deploy.
5. After build succeeds, go to **Custom domains** → add `argusrecruit.com` and `www.argusrecruit.com`.

### 5. Trigger rebuilds when content changes (5 min)

In Sanity Studio, configure a webhook:
1. https://www.sanity.io/manage → your project → API → Webhooks → Create
2. URL: get from Cloudflare Pages → Settings → Builds → "Deploy hooks"
3. Trigger on: Create / Update / Delete

Now every published post auto-deploys to your live site.

## ✏️ Writing blog posts

1. Go to `argusrecruit.sanity.studio` and log in
2. Click **+ New Document** → **Blog Post**
3. Fill in: title, slug (auto from title), language, tag, excerpt, cover image, body, publishedAt
4. Click **Publish**
5. ~30 seconds later, your post is live on argusrecruit.com/blog/

## 🛠 Local development

```bash
# In argusrecruit-astro/
npm install
npm run dev       # http://localhost:4321
```

If Sanity isn't configured, the blog will use the built-in fallback posts (3 sample articles).

## 📋 What's already set up

- ✅ All 3 languages (EN/RU/HY) with proper hreflang
- ✅ Sitemap auto-generated at /sitemap-index.xml
- ✅ robots.txt
- ✅ JSON-LD schema (Organization + Service)
- ✅ Contact form using Web3Forms (file attachments up to 10MB)
- ✅ All images optimized in /public/

## 📚 SEO submission to Google

1. Go to https://search.google.com/search-console
2. Add property: `argusrecruit.com`
3. Verify via DNS TXT record (Cloudflare DNS makes this easy)
4. Submit sitemap: `https://argusrecruit.com/sitemap-index.xml`
5. Done — Google starts indexing within 1–7 days.
