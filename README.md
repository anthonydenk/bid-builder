# Bid Builder

Bid Builder turns CAD-exported Excel workbooks into checked commercial proposals for restroom accessories, toilet partitions, and similar build-out packages. It keeps the CAD system as the pricing source of truth, asks for the job facts the export does not contain, and produces an audit trail before a proposal is called client-ready.

The short trigger is **Bid Builder**. In Codex, invoke it as `$bid-builder`. In ChatGPT, select or mention the installed Bid Builder skill/plugin; ChatGPT’s explicit skill mention uses `@`, while `/` opens the command menu rather than naming the skill.

## The first question is intentional

After you trigger Bid Builder and attach spreadsheets, it asks for one complete intake packet before touching workbook contents. The packet includes:

- job naming, proposal numbers, and grouping;
- proposal date and plans-dated date;
- prepared-by contact;
- partition manufacturer, count, material, mounting/brace style, and installation basis; and
- whether the stored company terms are approved or draft-only.

The packet shows the Falcon example as the naming pattern. Every separate proposal needs its own confirmed proposal number. Filenames are not trusted to make those decisions.

## What it does

- Accepts one or more `.xlsx` exports.
- Groups floors, suites, phases, or buildings only as instructed.
- Reconciles every row of `Net Price`, including internal freight/labor/overhead rows.
- Shows product scope without leaking internal estimating lines.
- Uses a maintained catalog for manufacturer/model mapping and special cases.
- Stops on missing project facts, unresolved products, bad spreadsheet arithmetic, or unapproved business clauses.
- Produces `proposal-model.json`, `reconciliation.md`, `proposal.docx`, and optionally `proposal.pdf`.

## Fastest setup for your friend

This is a private repository, so first give your friend read access under **GitHub → Settings → Collaborators**.

### Codex desktop, CLI, or IDE

```bash
git clone https://github.com/anthonydenk/bid-builder.git
cd bid-builder
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Open that folder in Codex. The repo-scoped skill is discovered automatically. Start a new task, attach the exports, and say:

```text
$bid-builder build a proposal from these exports
```

If you want Bid Builder available outside this repository, copy `skills/bid-builder` into `~/.agents/skills/bid-builder` and restart Codex.

### ChatGPT desktop / ChatGPT-only user

Standalone local skills are supported by the ChatGPT desktop app. The installable plugin in this repo is also structured for ChatGPT and Codex, but a private plugin still needs a local/private marketplace or a workspace admin publication path; it does not automatically appear in the public plugin directory.

For the simplest ChatGPT-only handoff, download the repository ZIP, unzip it, and add/copy the `skills/bid-builder` folder through the desktop app’s Skills workflow. Then start a new chat, attach the workbooks, and mention **@Bid Builder**. Web/mobile availability requires the skill to be installed as a plugin in the user’s ChatGPT workspace.

### Install as a private plugin

With GitHub access configured on the machine:

```bash
codex plugin marketplace add anthonydenk/bid-builder
codex plugin add bid-builder@bid-builder-private
```

Restart the ChatGPT desktop app and use a new task. The repo contains both `.codex-plugin/plugin.json` and a repo marketplace entry.

## Daily workflow

1. Finish the takeoff/specification in CAD.
2. Export each relevant area to `.xlsx`.
3. Trigger Bid Builder and attach the exports.
4. Answer the single mandatory intake packet—including dates, preparer, partition scope, and terms approval.
5. Confirm the short intake summary. Only spreadsheet-specific exceptions, such as an unknown accessory manufacturer, should surface afterward.
6. Review the reconciliation report and generated proposal.

The proposal is marked `DRAFT — NOT FOR ISSUE` until the company has approved its clause profile. That is deliberate: exclusions and contract language are business/legal policy, not facts the model should invent.

## One-time company setup

Before production use, review and customize:

- `skills/bid-builder/references/item-catalog.json` — manufacturer/model mappings and public-row overrides.
- `skills/bid-builder/references/clauses.json` — approved inclusions, exclusions, commercial terms, and special profiles.
- `skills/bid-builder/references/company-profile.json` — company identity transcribed from the example proposal; verify it before issue.
- `skills/bid-builder/assets/proposal-template.docx` — logo, legal entity, address, contact information, colors, footer, and signature block.

Set `clauses_approved` to `true` in a job only after the selected clause profile is actually approved.

## Deterministic command-line engine

```bash
python skills/bid-builder/scripts/bid_builder.py inspect path/to/export.xlsx

python skills/bid-builder/scripts/bid_builder.py build \
  --job path/to/job.json \
  --output-dir outputs/Q10001
```

The engine reads `.xlsx` with Python’s standard library and uses `python-docx` only for document generation. LibreOffice is optional for PDF conversion.
It exits with status `2` when it successfully creates a draft that still has blockers; this is intentional and makes CI/automation stop before issue.

## Safety and privacy

Never commit client workbooks, local job files, generated pricing, or proposals. The repository ignores `client-data/`, `private/`, `tmp/`, `outputs/`, and `*.local.json`. The provided fixtures are synthetic.

## Development

```bash
pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
python /path/to/skill-creator/scripts/quick_validate.py skills/bid-builder
python /path/to/plugin-creator/scripts/validate_plugin.py .
```

The real sample attachments were used only for local verification and are not included in this repository.
