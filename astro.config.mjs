import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

export default defineConfig({
  site: 'https://argusrecruit.com',
  build: { format: 'directory' },
  integrations: [sitemap()]
});
