# interview-review · Interview Review Skill

**[简体中文](README.md) | English**

Turn an interview recording transcript into a **question-by-question, accumulative** Word review document.

The core value is not record-keeping, but the **improved answers** and **transferable takeaways** — rewriting every poorly-answered question into a version you can use directly in your next interview.

## What Problem It Solves

After an interview, you have the recording transcribed — but then what?

- You don't know which answers were good and which bombed
- You keep making the same mistakes (self-deprecation, drowning the conclusion in details) with no systematic tracking
- Interview experience posts online only have questions, not a side-by-side of "my answer vs. the ideal answer"

This skill standardizes the review workflow: **verbatim extraction → per-question rating (✅/🟡/🔴) → perfect-version improved answers → transferable takeaways**, accumulating every interview into a single growth archive.

## Features

- **Per-question structuring**: each question gets "Interviewer's question (verbatim) / My answer (verbatim) / AI evaluation / Improved answer / Takeaway"
- **Three-tier rating**: ✅ Good / 🟡 Could improve / 🔴 Problem — not everything gets a ✅; the 🔴 and 🟡 are where reviews earn their value
- **Improved answers are *perfect* answers**: not a polish of your words — the skill first judges whether the content itself is wrong (direction / facts / knowledge), and if so, rewrites from scratch
- **Integrity red line**: improved answers never fabricate experience or skills you don't have — there are three ways to tell the truth (fabricating ❌ / pure self-deprecation ❌ / honesty + strategy ✅)
- **Q&A (reverse-ask) section handled separately**: your questions are evaluated (were they good questions?) and the interviewer's answers are analyzed (what signals did they reveal?)
- **Self-sabotage detection**: built-in recognition of 5 disguised forms of self-deprecation, plus the key distinction between "rational honesty" and "self-deprecation"
- **Multi-session accumulation**: each interview appends as a chapter to one Word master archive; split into standalone per-session docs on demand

## Installation

Place the whole folder into your AI agent's skills directory (e.g. `~/.zcode/skills/`, or the corresponding directory for Claude/Codex):

```
skills/
└── interview-review/               # rename after cloning (or git clone directly to this name)
    ├── SKILL.md                    # Workflow & disciplines
    ├── agents/openai.yaml          # Interface definition
    ├── references/
    │   ├── methodology.md          # Rating criteria / answer frameworks / improved-answer spec (core)
    │   ├── session-schema.md       # Data structure spec
    │   └── transcription-errors.md # Transcription error lookup table (grows per session)
    └── scripts/
        ├── gen_review.py           # Word document generator (python-docx)
        └── split_sessions.py       # Split master archive into per-session docs
```

Dependency: `pip install python-docx`

## Usage

After an interview, send the recording transcript to your agent and say "help me review this interview."

**First session** (create the master archive):
```bash
python scripts/gen_review.py --session session.json --out review_archive.docx
```

**Append** (every subsequent session):
```bash
python scripts/gen_review.py --session session.json --out review_archive.docx --append
```

**Split per-session docs** (one standalone doc per interview, quick to browse):
```bash
python scripts/split_sessions.py --in review_archive.docx --outdir sessions/
```

The session.json field spec is in `references/session-schema.md`.

## Output Format

One table per question:

| Field | Content |
|-------|---------|
| Interviewer's question | Verbatim from the recording (stutters cleaned, never summarized) |
| My answer (verbatim) | Verbatim, all substantive content preserved |
| AI evaluation | ✅/🟡/🔴 three tiers + specific commentary |
| Improved answer | Perfect version — content correctness judged first, then rewritten; ready to memorize |
| Takeaway | Transferable expression techniques, not vague "it's better now" |

The reverse-ask section is rendered separately: your questions get evaluated (good or not), and the interviewer's answers get analyzed (what signals they reveal).

## Review Methodology at a Glance

Full rating criteria in `references/methodology.md`. Core principles:

1. **Never abbreviate verbatim quotes** — evaluation and improvement need anchors
2. **Improved answers follow "Conclusion → Actions → Result"** — one-sentence summary → key actions → result
3. **Judge content correctness before rewriting** — polishing an answer whose direction is wrong is pointless
4. **Integrity first** — no fabricated content; expression strategy stays within the bounds of truth
5. **Distinguish rational honesty from self-deprecation** — saying "this couldn't be measured" based on real constraints + offering alternative evidence = a plus; gratuitously negating yourself = a minus

## License

MIT
