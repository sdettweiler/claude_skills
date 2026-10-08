---
name: remove-AI-slop
description: Detect AI writing tells in a piece of text, then rewrite it to sound human. Use when the user runs /remove-AI-slop, pastes text and asks to "de-AI" / "humanize" / "remove AI slop", or asks whether text sounds AI-written.
metadata:
  trigger: /remove-AI-slop, "humanize this", "does this sound AI", "remove the AI tells"
  source: Wikipedia — "Signs of AI writing" (en.wikipedia.org/wiki/Wikipedia:Signs_of_AI_writing)
---

# Remove AI Slop

Two jobs, in order: **diagnose** the text against a taxonomy of AI tells, then **rewrite** it so a human wrote it. Never rewrite silently — the user should see what was wrong.

Related: `stop-slop` is the *don't-write-slop-in-the-first-place* companion. This skill is the *fix-existing-text* tool. When in doubt, run both.

## Input

Whatever the user gives you: pasted text, a file path, a draft you just wrote, or "the last thing you sent." If no text is supplied, ask for it (once) — do not invent a sample.

## Step 1 — Diagnose

Read the full taxonomy in [references/tells.md](references/tells.md). Scan the text against every category and produce a short report **before** touching the prose:

```
## AI-slop scan

**Verdict:** <clean / light tells / heavily AI-flavored>

**Tells found:**
1. [Word choice] "delve into the rich tapestry" — line 2
2. [Structure] "Not just X, but Y" negative parallelism — line 5
3. [Punctuation] 4 em dashes used as connective glue
4. [Tone] Promotional closer "stands as a testament to…"
...
```

Rules for the scan:
- Cite the **actual offending words**, not a paraphrase. Quote them.
- Point to where (line, paragraph, or the phrase itself).
- Multiple weak signals matter more than one — a single em dash is not slop; em dashes + rule-of-three + "moreover" together are.
- If the text is genuinely clean, say so and stop. Don't manufacture problems.

## Step 2 — Rewrite

Rewrite the text to remove every tell you flagged while preserving meaning, facts, and the author's intent. Then show it.

Hard rules:
1. **Keep every fact.** Never drop, add, or soften a claim to make prose flow. If a claim was vague in the original, flag it — don't invent specifics.
2. **Match the author's register.** A casual Slack message stays casual; a legal memo stays formal. "Human" ≠ "chatty."
3. **Cut, don't swap.** The fix for AI vocabulary is usually deletion, not a thesaurus lookup. "Delve into" → "look at" is fine; often the whole clause goes.
4. **Vary rhythm.** Break runs of same-length sentences. Two beats three. Not every paragraph needs a punchy closer.
5. **No em dashes** as connective glue. Use a period, comma, or parenthesis.
6. **Active voice, real subjects.** A person or thing does the verb. No "the decision underscores…".
7. **Kill the significance-inflation.** Delete "stands as a testament," "plays a vital role," "in an ever-evolving landscape," "it's important to note." Just state the thing.
8. **No formulaic scaffolding.** Drop bolt-on "Challenges / Future Outlook" sections, forced rule-of-three, "In conclusion," and title-case headings unless the format demands them.

## Step 3 — Confirm

End with a one-line diff-in-words: what changed and why, e.g.
> Removed 3 em dashes, cut "delve/tapestry/testament," broke two rule-of-three lists, replaced the promotional closer with a plain statement. Facts unchanged.

If the user only wanted a diagnosis (they said "check" not "fix"), stop after Step 1.

## Scoring (optional, on request)

Rate the *original* 1–5 on each; anything ≤3 needs the rewrite:

| Dimension | Question |
|-----------|----------|
| Vocabulary | Free of AI-favored words (delve, tapestry, testament, robust, pivotal)? |
| Structure | Varied sentences, no negative parallelism / forced triples? |
| Punctuation | No em-dash glue, straight quotes? |
| Tone | Neutral, not promotional or significance-inflated? |
| Specificity | Concrete facts, not vague filler that fits any topic? |

## Anti-goals

- Do not make text worse to "prove" you edited it.
- Do not strip an author's genuine voice just because it's polished.
- Do not add hedges, disclaimers, or knowledge-cutoff caveats.
- Do not turn everything into terse fragments — that's just a different tell.
