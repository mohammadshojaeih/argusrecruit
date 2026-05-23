import { defineType, defineField } from 'sanity'

export default defineType({
  name: 'job',
  title: 'Job Posting',
  type: 'document',
  fields: [
    defineField({ name: 'title', title: 'Job Title', type: 'string', validation: r => r.required() }),
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
    defineField({
      name: 'status', type: 'string',
      options: { list: [
        { title: 'Active', value: 'active' },
        { title: 'Closed', value: 'closed' },
        { title: 'Hidden', value: 'hidden' }
      ] },
      initialValue: 'active'
    }),
    defineField({ name: 'featured', title: 'Featured', type: 'boolean', initialValue: false }),
    defineField({
      name: 'department', title: 'Department', type: 'string',
      options: { list: [
        'Executive', 'Engineering', 'Marketing', 'Finance',
        'Healthcare', 'Operations', 'Sales', 'Product', 'HR', 'Other'
      ] }
    }),
    defineField({
      name: 'employmentType', title: 'Employment Type', type: 'string',
      options: { list: [
        { title: 'Full-time', value: 'FULL_TIME' },
        { title: 'Part-time', value: 'PART_TIME' },
        { title: 'Contract', value: 'CONTRACTOR' },
        { title: 'Temporary', value: 'TEMPORARY' },
        { title: 'Internship', value: 'INTERN' }
      ] },
      initialValue: 'FULL_TIME'
    }),
    defineField({
      name: 'workplaceType', title: 'Workplace Type', type: 'string',
      options: { list: [
        { title: 'On-site', value: 'onsite' },
        { title: 'Remote', value: 'remote' },
        { title: 'Hybrid', value: 'hybrid' }
      ] }
    }),
    defineField({ name: 'locationCity', title: 'City', type: 'string' }),
    defineField({ name: 'locationCountry', title: 'Country', type: 'string' }),
    defineField({ name: 'salaryMin', title: 'Salary Min (annual)', type: 'number' }),
    defineField({ name: 'salaryMax', title: 'Salary Max (annual)', type: 'number' }),
    defineField({
      name: 'salaryCurrency', title: 'Currency', type: 'string',
      options: { list: ['USD','EUR','GBP','AED','AMD','CAD'] },
      initialValue: 'USD'
    }),
    defineField({ name: 'excerpt', title: 'Short Summary', type: 'text', rows: 2 }),
    defineField({ name: 'description', title: 'Description', type: 'array', of: [{ type: 'block' }] }),
    defineField({ name: 'responsibilities', title: 'Responsibilities', type: 'array', of: [{ type: 'string' }] }),
    defineField({ name: 'requirements', title: 'Requirements', type: 'array', of: [{ type: 'string' }] }),
    defineField({ name: 'niceToHave', title: 'Nice to Have', type: 'array', of: [{ type: 'string' }] }),
    defineField({ name: 'tags', title: 'Tags', type: 'array', of: [{ type: 'string' }] }),
    defineField({ name: 'publishedAt', type: 'datetime', initialValue: () => new Date().toISOString() }),
    defineField({ name: 'expiresAt', title: 'Expires At', type: 'datetime' })
  ],
  preview: {
    select: { title: 'title', subtitle: 'department', status: 'status' },
    prepare({ title, subtitle, status }) {
      return { title, subtitle: `${subtitle || ''} ${status ? '· ' + status : ''}` }
    }
  }
})
