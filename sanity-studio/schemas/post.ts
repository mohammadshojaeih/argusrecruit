import { defineType, defineField } from 'sanity'

export default defineType({
  name: 'post',
  title: 'Blog Post',
  type: 'document',
  fields: [
    defineField({ name: 'title', type: 'string', validation: r => r.required() }),
    defineField({
      name: 'slug',
      type: 'slug',
      options: {
        source: 'title',
        // Scope uniqueness by language so EN/RU/HY translations of the same
        // post can share a slug. Default behaviour is to enforce uniqueness
        // across ALL docs of the type, which blocks translation publishing.
        isUnique: async (slug, context) => {
          const { document, getClient } = context;
          const client = getClient({ apiVersion: '2024-01-01' });
          const publishedId = document._id.replace(/^drafts\./, '');
          const draftId = 'drafts.' + publishedId;
          const lang = document.language || 'en';
          const count = await client.fetch(
            `count(*[_type=="post" && !(_id in [$draftId, $publishedId]) && slug.current==$slug && language==$lang])`,
            { draftId, publishedId, slug, lang }
          );
          return count === 0;
        }
      },
      validation: r => r.required()
    }),
    defineField({
      name: 'language', type: 'string',
      options: { list: [
        { title: 'English', value: 'en' },
        { title: 'Russian', value: 'ru' },
        { title: 'Armenian', value: 'hy' }
      ] },
      initialValue: 'en'
    }),
    defineField({ name: 'tag', type: 'string', description: 'e.g. Leadership Hiring' }),
    defineField({ name: 'excerpt', type: 'text', rows: 3 }),
    defineField({ name: 'cover', type: 'image', options: { hotspot: true } }),
    defineField({ name: 'body', type: 'array', of: [{ type: 'block' }, { type: 'image' }] }),
    defineField({ name: 'publishedAt', type: 'datetime', initialValue: () => new Date().toISOString() })
  ],
  preview: {
    select: { title: 'title', media: 'cover', subtitle: 'tag' }
  }
})
