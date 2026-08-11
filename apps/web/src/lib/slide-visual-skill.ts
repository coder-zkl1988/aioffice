export const NATIVE_VISUAL_SKILL_ID = 'genoffice-native-v1'

export const NATIVE_VISUAL_LAYOUTS = [
  'cover_statement',
  'section_divider',
  'title_body',
  'split_visual',
  'comparison',
  'three_points',
  'big_number',
  'timeline',
  'process',
  'data_chart',
  'quote',
  'closing',
] as const

export const GENOFFICE_NATIVE_VISUAL_SKILL = `GenOffice Native Visual Skill v1
- Start from the communication job: decide the one claim this slide must make.
- Use a takeaway title, not a generic section label. Keep one claim per slide.
- Prefer one strong composition over a dashboard of small cards. Use cards only for real repeated items.
- Use a 64 px outer safe area and an 8 px spacing rhythm. Align objects to shared edges.
- Typography budgets: title 32-48 pt, key number 40-64 pt, body 18-26 pt, annotation at least 14 pt.
- Density budgets: at most 3 content groups, at most 6 short bullets, at most 40 native elements.
- Create contrast through scale, weight, whitespace, and one restrained accent color.
- Vary layouts across a deck. Choose one semantic layout: ${NATIVE_VISUAL_LAYOUTS.join(', ')}.
- Keep all important content editable: text stays text, diagrams use shapes and lines, data uses tables or charts.
- Never place decorative rounded rectangles behind every paragraph. Never use gradients, decorative blobs, or tiny UI-like controls.
- Preserve factual accuracy. Do not invent names, dates, numbers, sources, or image URLs.`
