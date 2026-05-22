import { defineType, defineField } from 'sanity'

export default defineType({
  name: 'post',
  title: 'Blog Post',
  type: 'document',
  fields: [
    defineField({ name: 'title', type: 'string', validation: r => r.required() }),
    defineField({ name: 'slug', type: 'slug', options: { source: 'title' }, validation: r => r.required() }),
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
