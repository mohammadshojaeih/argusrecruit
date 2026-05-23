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

const FALLBACK_JOBS_EN = [
  {
    _id: 'fb-job-1',
    title: 'Chief Marketing Officer',
    slug: 'chief-marketing-officer',
    status: 'active',
    featured: true,
    department: 'Marketing',
    employmentType: 'FULL_TIME',
    workplaceType: 'hybrid',
    locationCity: 'London',
    locationCountry: 'United Kingdom',
    salaryMin: 180000,
    salaryMax: 240000,
    salaryCurrency: 'GBP',
    excerpt: 'Lead the marketing function at a fast-scaling B2B SaaS company. Build the brand, build the team, drive growth.',
    description: [{ _type: 'block', children: [{ _type: 'span', text: 'Our client is a Series C B2B SaaS company looking for a senior marketing leader to take their brand to the next level. You will own go-to-market strategy, build a world-class team, and report directly to the CEO.' }] }],
    responsibilities: [
      'Own GTM strategy across product, content, performance, and lifecycle',
      'Build and lead a marketing org of 20+',
      'Partner with sales on pipeline and ABM',
      'Define brand positioning for enterprise expansion'
    ],
    requirements: [
      '10+ years of B2B marketing experience',
      'Prior CMO or VP Marketing experience at a $50M+ ARR company',
      'Proven track record building marketing teams from scratch',
      'Deep understanding of demand generation and ABM'
    ],
    niceToHave: ['SaaS experience', 'PE/VC-backed company experience'],
    tags: ['CMO', 'B2B SaaS', 'Leadership'],
    publishedAt: '2026-05-20T00:00:00Z'
  },
  {
    _id: 'fb-job-2',
    title: 'Senior Backend Engineer (Go / Kubernetes)',
    slug: 'senior-backend-engineer-go',
    status: 'active',
    featured: false,
    department: 'Engineering',
    employmentType: 'FULL_TIME',
    workplaceType: 'remote',
    locationCity: 'Remote',
    locationCountry: 'Europe / North America',
    salaryMin: 130000,
    salaryMax: 180000,
    salaryCurrency: 'USD',
    excerpt: 'Build distributed systems at scale. Modern Go stack, K8s, observability done right.',
    description: [{ _type: 'block', children: [{ _type: 'span', text: 'Join an engineering team obsessed with reliability. Our client runs mission-critical infrastructure for global financial customers and needs senior engineers comfortable with distributed systems.' }] }],
    responsibilities: [
      'Design and implement scalable microservices in Go',
      'Own services end-to-end including on-call rotation',
      'Improve observability and incident response',
      'Mentor mid-level engineers'
    ],
    requirements: [
      '6+ years of backend engineering',
      'Strong Go (or willingness to ramp from another typed language)',
      'Hands-on Kubernetes in production',
      'Comfortable with distributed systems trade-offs'
    ],
    niceToHave: ['gRPC', 'PostgreSQL internals', 'Open-source contributions'],
    tags: ['Go', 'Kubernetes', 'Backend', 'Remote'],
    publishedAt: '2026-05-15T00:00:00Z'
  },
  {
    _id: 'fb-job-3',
    title: 'VP of Finance',
    slug: 'vp-of-finance',
    status: 'active',
    featured: false,
    department: 'Finance',
    employmentType: 'FULL_TIME',
    workplaceType: 'onsite',
    locationCity: 'Dubai',
    locationCountry: 'United Arab Emirates',
    salaryMin: 220000,
    salaryMax: 300000,
    salaryCurrency: 'USD',
    excerpt: 'Lead finance at a fast-growing FinTech expanding across MENA. Pre-IPO trajectory.',
    description: [{ _type: 'block', children: [{ _type: 'span', text: 'Confidential search for a VP Finance to lead the function at a regulated FinTech. You will own FP&A, treasury, controllership, and prepare the company for an IPO within 24 months.' }] }],
    responsibilities: [
      'Lead FP&A, controllership, and treasury',
      'Own audit, tax, and regulatory reporting',
      'Drive IPO readiness',
      'Build out finance org from 8 to 25+'
    ],
    requirements: [
      'CPA / ACCA / CFA',
      '12+ years in finance, including senior roles at PE/VC-backed scale-ups',
      'Prior IPO experience (or strong adjacent)',
      'MENA regulatory familiarity strongly preferred'
    ],
    niceToHave: ['FinTech background', 'Big 4 experience'],
    tags: ['VP Finance', 'FinTech', 'IPO'],
    publishedAt: '2026-05-10T00:00:00Z'
  }
];

const FALLBACK_JOBS = {
  en: FALLBACK_JOBS_EN,
  ru: FALLBACK_JOBS_EN,
  hy: FALLBACK_JOBS_EN
};

export async function getJobs(lang = 'en', { includeClosed = false } = {}) {
  if (!sanityClient) {
    const all = FALLBACK_JOBS[lang] || FALLBACK_JOBS.en;
    return includeClosed ? all : all.filter(j => j.status === 'active');
  }
  try {
    const statusFilter = includeClosed ? `&& status != "hidden"` : `&& status == "active"`;
    const query = `*[_type=="job" ${statusFilter} && (language=="${lang}" || !defined(language)) && (!defined(expiresAt) || expiresAt > now())] | order(featured desc, publishedAt desc) {
      _id, title, slug, status, featured, department, employmentType, workplaceType,
      locationCity, locationCountry, salaryMin, salaryMax, salaryCurrency, salaryNegotiable,
      excerpt, description, responsibilities, requirements, niceToHave, tags,
      publishedAt, expiresAt
    }`;
    const jobs = await sanityClient.fetch(query);
    if (!jobs || jobs.length === 0) {
      const fb = FALLBACK_JOBS[lang] || FALLBACK_JOBS.en;
      return includeClosed ? fb : fb.filter(j => j.status === 'active');
    }
    return jobs.map(j => ({ ...j, slug: j.slug?.current || j.slug }));
  } catch (e) {
    console.error('Sanity job fetch failed, using fallback:', e.message);
    const fb = FALLBACK_JOBS[lang] || FALLBACK_JOBS.en;
    return includeClosed ? fb : fb.filter(j => j.status === 'active');
  }
}

export async function getJob(slug, lang = 'en') {
  const jobs = await getJobs(lang, { includeClosed: true });
  return jobs.find(j => j.slug === slug) || null;
}

