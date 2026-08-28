# Proposal rules

## Grouping

Build one proposal model per grouping explicitly approved in the intake line. Multiple source files can feed one section or multiple sections within a proposal. Never merge suites, floors, phases, or buildings merely because their filenames look related.

## Accessories

- Aggregate identical public products across all files in the same customer-facing section.
- Show quantity, manufacturer, model/description, and customer-facing scope language.
- Keep internal pricing rows out of the public item schedule.
- Reconcile `visible public value + hidden allowance = source section total`.

## Partitions

- The spreadsheet price can establish the section total.
- The public description must come from explicit scope input, such as `(5) Stalls / Powder Coated / Overhead Braced`.
- Never derive stall count or construction type from row count, quantity, filename, or description fragments.

## Manufacturer catalog

Apply `item-catalog.json` rules in order. Prefer exact item matches, then description regular expressions. An unmatched public line receives an `UNRESOLVED` manufacturer and blocks client-ready output.

## Clauses

The included clauses are a normalized starting library, not approved legal language. Set `clauses_approved: true` only after the operating company has reviewed the selected profile. Until then, generated documents must display `DRAFT — NOT FOR ISSUE`.

## Money and auditability

- Use `Decimal`, never binary floating point, for commercial calculations.
- Round displayed currency to cents.
- Preserve source filenames in the reconciliation report and model.
- If any arithmetic check fails, show the exact workbook/sheet/row and block client-ready status.
