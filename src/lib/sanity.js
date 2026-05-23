import { createClient } from '@sanity/client';

const projectId = import.meta.env.PUBLIC_SANITY_PROJECT_ID || '';
const dataset = import.meta.env.PUBLIC_SANITY_DATASET || 'production';

export const sanityClient = projectId ? createClient({
  projectId, dataset, apiVersion: '2024-01-01', useCdn: true,
}) : null;

const BODY_LEAD = `<p>When the wrong leader walks through the door, the cost is rarely the salary line. It's the eighteen months of compounding damage that comes after — a team that quietly stops trusting the next leader, customers who don't renew because something feels off, and the silent veto your best people apply when they update their LinkedIn.</p>
<p>Across our placements, the most expensive mistake we see at the executive level isn't a missed quota or a botched product launch. It's the time it takes to admit the hire was wrong.</p>
<h2>What we actually measure</h2>
<p>When clients ask us to quantify the cost of a bad senior hire, we walk them through three buckets that almost always go unaccounted for:</p>
<h3>1. Decision Latency</h3>
<p>A leader who can't make calls slows down every team beneath them.</p>
<h3>2. Talent Drift</h3>
<p>Strong individual contributors don't quit on day one. They quit on day ninety.</p>
<h3>3. The Successor Tax</h3>
<p>The hire that comes after a bad hire is harder. The team is bruised.</p>
<p style="margin-top: 50px; padding-top: 30px; border-top: 1px solid rgba(212,175,55,0.15); color: rgba(255,255,255,0.6); font-style: italic;">— The ArgusRecruit Team</p>`;

const BODY_DISCREET = `<p>Confidential executive search is the quiet art of finding leaders without disturbing the market. This post is a placeholder — edit it in Sanity Studio.</p>`;
const BODY_TIMEZONES = `<p>Three time zones, one process. This post is a placeholder — edit it in Sanity Studio.</p>`;

const enPosts = [
  {
    _id: 'fb-1',
    title: 'The Hidden Cost of a Bad Leadership Hire',
    slug: 'hidden-cost-of-bad-leadership-hire',
    tag: 'Leadership Hiring',
    excerpt: "When the wrong C-level hire walks through the door, the bill goes far beyond their salary. We break down the real numbers.",
    cover: '/slide-1.jpg',
    publishedAt: '2026-05-22T00:00:00Z',
    body: BODY_LEAD
  },
  {
    _id: 'fb-2',
    title: 'Discreet Search: When Confidentiality Matters Most',
    slug: 'discreet-search-when-confidentiality-matters',
    tag: 'Confidential Search',
    excerpt: "How modern boutique firms run a discreet executive search without tipping off the market — and why it matters.",
    cover: '/slide-2.jpg',
    publishedAt: '2026-05-18T00:00:00Z',
    body: BODY_DISCREET
  },
  {
    _id: 'fb-3',
    title: "Working Across Three Time Zones — What We've Learned",
    slug: 'working-across-three-time-zones',
    tag: 'Global Search',
    excerpt: 'Lessons from running executive searches across the UK, UAE, and Canada — culture, candidate motivation, and timing.',
    cover: '/slide-3.jpg',
    publishedAt: '2026-05-14T00:00:00Z',
    body: BODY_TIMEZONES
  }
];

const FALLBACK_POSTS = {
  en: enPosts,
  ru: enPosts.map(p => ({...p})),
  hy: enPosts.map(p => ({...p}))
};

export async function getPosts(lang = 'en') {
  if (!sanityClient) return FALLBACK_POSTS[lang] || FALLBACK_POSTS.en;
  try {
    const query = `*[_type=="post" && defined(publishedAt) && (language=="${lang}" || !defined(language))] | order(publishedAt desc) {
      _id, title, slug, tag, excerpt, "cover": cover.asset->url, publishedAt, body
    }`;
    const posts = await sanityClient.fetch(query);
    if (!posts || posts.length === 0) return FALLBACK_POSTS[lang] || FALLBACK_POSTS.en;
    return posts.map(p => ({ ...p, slug: p.slug?.current || p.slug }));
  } catch (e) {
    console.error('Sanity fetch failed, using fallback:', e.message);
    return FALLBACK_POSTS[lang] || FALLBACK_POSTS.en;
  }
}

export async function getPost(slug, lang = 'en') {
  const posts = await getPosts(lang);
  return posts.find(p => p.slug === slug) || null;
}

// ===================== JOBS =====================

export async function getJobs(lang = 'en', { includeClosed = false } = {}) {
  if (!sanityClient) return [];
  try {
    const statusFilter = includeClosed ? `&& status != "hidden"` : `&& status == "active"`;
    const query = `*[_type=="job" ${statusFilter} && (language=="${lang}" || !defined(language)) && (!defined(expiresAt) || expiresAt > now())] | order(featured desc, publishedAt desc) {
      _id, title, slug, status, featured, department, employmentType, workplaceType,
      locationCity, locationCountry, salaryMin, salaryMax, salaryCurrency, salaryNegotiable,
      excerpt, description, responsibilities, requirements, niceToHave, tags,
      publishedAt, expiresAt
    }`;
    const jobs = await sanityClient.fetch(query);
    if (!jobs) return [];
    return jobs.map(j => ({ ...j, slug: j.slug?.current || j.slug }));
  } catch (e) {
    console.error('Sanity job fetch failed:', e.message);
    return [];
  }
}

export async function getJob(slug, lang = 'en') {
  const jobs = await getJobs(lang, { includeClosed: true });
  return jobs.find(j => j.slug === slug) || null;
}

