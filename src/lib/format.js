export const fmtEmploymentType = (e) => ({
  FULL_TIME: 'Full-time',
  PART_TIME: 'Part-time',
  CONTRACTOR: 'Contract',
  TEMPORARY: 'Temporary',
  INTERN: 'Internship'
}[e] || e);

export const fmtWorkplaceType = (w) => ({
  onsite: 'On-site',
  remote: 'Remote',
  hybrid: 'Hybrid'
}[w] || (w ? w[0].toUpperCase() + w.slice(1) : w));
