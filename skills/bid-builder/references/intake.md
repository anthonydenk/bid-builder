# Intake gate

## Required naming and grouping line

Before reading newly attached spreadsheets, require one natural-language line that establishes:

1. project name;
2. proposal number;
3. customer or general contractor;
4. city/location; and
5. which floor, suite, phase, or building exports belong in each proposal.

Use this exact prompt when the line is absent:

> Please provide the job naming and grouping line in this format: `Falcon A, proposal B212492, GCON, Mesa. Floors 1 and 2 go together. Suite 126 is separate.`

The example means two outputs:

- Proposal B212492 for Falcon A / GCON / Mesa, combining Floors 1 and 2.
- A separate proposal group for Suite 126. If it needs a distinct proposal number or customer-facing title, ask for it.

Never infer the grouping from filenames. Never silently reuse a proposal number for a separate group.

## Follow-up facts

After confirming grouping, collect only facts that remain missing:

- proposal date;
- plan date;
- prepared-by name and contact block;
- customer contact/address if required by the template;
- explicit partition scope description;
- unresolved manufacturer mappings;
- approved clause profile; and
- tax, bond, permit, freight, installation, or escalation treatment when company policy requires it.

Restate assumptions as a compact checklist before generation.
