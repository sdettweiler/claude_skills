# The taxonomy of AI writing tells

Source: Wikipedia, *Signs of AI writing*. No single tell is proof — clusters are. Weight
accordingly.

## 1. Word choice

**AI-favored vocabulary** (delete or replace with a plain word):
delve, tapestry, testament, boasts, bolstered, crucial, pivotal, key, landscape,
meticulous/meticulously, underscore, intricate/intricacies, interplay, enduring, garner,
vibrant, robust, profound, groundbreaking, renowned, nestled, showcase/showcasing,
foster/fostering, enhance, highlight/highlighting, align with, resonate with, valuable
insights, comprehensive, seamless.

**Copula avoidance** — replacing plain "is/are" with inflated verbs:
"serves as a," "stands as a," "marks the," "represents," "features," "offers," "boasts."
Fix: use "is."

**Vague attribution** — claims pinned to no one:
"industry reports," "observers have cited," "experts argue," "some critics argue,"
"several sources," "studies show." Fix: name the source or cut the claim.

## 2. Sentence & paragraph structure

- **Negative parallelism:** "Not just X, but Y," "Not only… but also," "It's not X, it's Y,"
  "X rather than Y." Fix: state Y directly.
- **Rule of three:** compulsive three-item lists and "adjective, adjective, adjective"
  triples. Fix: use one or two; vary the count.
- **Present-participle tails:** clauses ending in "-ing" that add fake analysis —
  "…, highlighting its importance," "…, reflecting broader trends," "…, ensuring success."
  Fix: delete the tail or make it a real, separate claim.
- **Metronomic rhythm:** runs of same-length sentences. Fix: mix short and long.

## 3. Punctuation & typography

- **Em-dash overuse** as connective glue (—). Fix: period, comma, or parenthesis.
- **Curly/smart quotes and apostrophes** where straight ones belong.
- **Emoji as section markers or bullets.**
- **Title Case In Headings** for every main word.
- **Thematic breaks (horizontal rules) before headings.**

## 4. Formatting

- **Bold overuse** — "key takeaways" bolding, boldfaced inline list headers followed by a
  colon and a description.
- **Inline-header vertical lists** everywhere instead of prose.
- **Rigid outline scaffolding:** bolt-on sections like "Challenges," "Future Outlook,"
  "In conclusion," regardless of whether the topic needs them.
- **Markdown leaking into plain-text / wiki contexts** (stray `**`, `#`).

## 5. Tone & rhetoric

- **Significance inflation:** "stands as a testament to," "plays a vital/pivotal/crucial
  role," "leaves an indelible mark," "in an ever-evolving landscape," "setting the stage
  for," "marks a turning point," "deeply rooted." Fix: state the fact, drop the halo.
- **Promotional / brochure voice:** "vibrant," "rich cultural heritage," "in the heart of,"
  "nestled," "diverse array," "commitment to." Reads like marketing, not information.
- **False balance:** formulaic "Despite its … it faces several challenges …" hedged
  even-handedness with no real content.
- **Overgeneralization:** one source's view presented as universal consensus.
- **Reader-addressing throat-clearing:** "It's important to note," "It's worth
  mentioning," "In today's world," "Let's dive in."

## 6. Content red flags

- **Fabricated significance:** "has generated debate" / "has been widely praised" with no
  evidence.
- **Generic filler that fits any topic:** statements so broad they carry no information.
- **Etymology/heritage padding** connecting mundane details to "broader significance."
- **Knowledge-cutoff disclaimers** or speculation about missing information.

## 7. Machine artifacts (dead giveaways when present)

Leftover generation markup: `contentReference`, `oaicite`, `oai_citation`, `turn0search0`,
`:::`, `+1`, `attached_file`, `grok_card`; broken/invalid DOIs and ISBNs; UTM tracking
params in citation URLs; references declared but never used. Any of these alone is
near-conclusive — remove immediately and check the surrounding text was pasted from an LLM.

## Detection reality check

Humans do no better than chance at spotting AI text by feel; even heavy LLM users top out
around 90% and only with clusters of signals. So: don't over-trust a single tell, and don't
claim certainty. Report what you found and how strong the cluster is.
