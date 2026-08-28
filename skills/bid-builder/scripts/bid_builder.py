#!/usr/bin/env python3
"""Deterministic CAD-export inspection and proposal generation for Bid Builder."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


MONEY = Decimal("0.01")
TOLERANCE = Decimal("0.02")
REQUIRED_HEADERS = {
    "type",
    "item",
    "quantity",
    "description",
    "net cost",
    "unit profit amount",
    "unit price",
    "extended cost",
    "extended profit",
    "net price",
}
INTERNAL_TYPES = {"freight", "labor", "overhead", "travel", "textura", "resource"}
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


class BidBuilderError(RuntimeError):
    """Actionable input or environment error."""


def money(value: Decimal) -> Decimal:
    return value.quantize(MONEY, rounding=ROUND_HALF_UP)


def quantity_text(value: Decimal) -> str:
    if value == value.to_integral_value():
        return format(value, ".0f")
    return format(value.normalize(), "f")


def decimal(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        return Decimal(str(value).replace(",", "").replace("$", "").strip())
    except (InvalidOperation, AttributeError) as exc:
        raise BidBuilderError(f"Expected a number, found {value!r}") from exc


def column_number(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    if not letters:
        return 0
    result = 0
    for char in letters.group(0):
        result = result * 26 + ord(char) - 64
    return result - 1


def normalized_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values: list[str] = []
    for item in root.findall(f"{{{NS_MAIN}}}si"):
        values.append("".join(node.text or "" for node in item.iter(f"{{{NS_MAIN}}}t")))
    return values


def _sheet_targets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{NS_PKG_REL}}}Relationship")
    }
    result: list[tuple[str, str]] = []
    sheets = workbook.find(f"{{{NS_MAIN}}}sheets")
    if sheets is None:
        return []
    for sheet in sheets:
        rel_id = sheet.attrib.get(f"{{{NS_REL}}}id", "")
        target = targets.get(rel_id, "")
        if target.startswith("/"):
            path = target.lstrip("/")
        elif target.startswith("xl/"):
            path = target
        else:
            path = f"xl/{target.lstrip('/')}"
        result.append((sheet.attrib.get("name", "Sheet"), path))
    return result


def _cell_value(cell: ET.Element, shared: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{NS_MAIN}}}t"))
    value_node = cell.find(f"{{{NS_MAIN}}}v")
    if value_node is None or value_node.text is None:
        return ""
    value = value_node.text
    if cell_type == "s":
        try:
            return shared[int(value)]
        except (ValueError, IndexError):
            return value
    if cell_type in {"str", "b"}:
        return value
    try:
        return Decimal(value)
    except InvalidOperation:
        return value


def _sheet_rows(archive: zipfile.ZipFile, target: str, shared: list[str]) -> list[tuple[int, list[Any]]]:
    root = ET.fromstring(archive.read(target))
    rows: list[tuple[int, list[Any]]] = []
    for row in root.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
        values: list[Any] = []
        for cell in row.findall(f"{{{NS_MAIN}}}c"):
            idx = column_number(cell.attrib.get("r", "A1"))
            while len(values) <= idx:
                values.append("")
            values[idx] = _cell_value(cell, shared)
        rows.append((int(row.attrib.get("r", len(rows) + 1)), values))
    return rows


@dataclass
class ExportRow:
    source: str
    sheet: str
    row_number: int
    type: str
    item: str
    quantity: Decimal
    description: str
    net_cost: Decimal
    unit_profit: Decimal
    unit_price: Decimal
    extended_cost: Decimal
    extended_profit: Decimal
    net_price: Decimal
    category: str

    def to_dict(self) -> dict[str, Any]:
        result = self.__dict__.copy()
        for key, value in result.items():
            if isinstance(value, Decimal):
                result[key] = str(money(value))
        return result


def read_export(path: Path) -> tuple[list[ExportRow], list[str]]:
    if path.suffix.lower() == ".xls":
        raise BidBuilderError(f"{path.name}: legacy .xls is not supported; re-export it as .xlsx")
    if path.suffix.lower() != ".xlsx":
        raise BidBuilderError(f"{path.name}: expected an .xlsx workbook")
    if not path.exists():
        raise BidBuilderError(f"Workbook does not exist: {path}")

    parsed: list[ExportRow] = []
    warnings: list[str] = []
    with zipfile.ZipFile(path) as archive:
        shared = _read_shared_strings(archive)
        for sheet_name, target in _sheet_targets(archive):
            rows = _sheet_rows(archive, target, shared)
            header_index = None
            header_map: dict[str, int] = {}
            for index, (_, values) in enumerate(rows):
                candidate = {normalized_header(v): i for i, v in enumerate(values) if str(v).strip()}
                if REQUIRED_HEADERS.issubset(candidate):
                    header_index = index
                    header_map = candidate
                    break
            if header_index is None:
                warnings.append(f"{path.name}/{sheet_name}: no recognized export header; sheet skipped")
                continue

            def get(values: list[Any], name: str) -> Any:
                idx = header_map.get(name)
                return values[idx] if idx is not None and idx < len(values) else ""

            for row_number, values in rows[header_index + 1 :]:
                description = str(get(values, "description") or "").strip()
                item = str(get(values, "item") or "").strip()
                quantity = decimal(get(values, "quantity"))
                net_price = decimal(get(values, "net price"))
                if not description and not item and quantity == 0 and net_price == 0:
                    continue
                parsed.append(
                    ExportRow(
                        source=path.name,
                        sheet=sheet_name,
                        row_number=row_number,
                        type=str(get(values, "type") or "").strip(),
                        item=item,
                        quantity=quantity,
                        description=description,
                        net_cost=decimal(get(values, "net cost")),
                        unit_profit=decimal(get(values, "unit profit amount")),
                        unit_price=decimal(get(values, "unit price")),
                        extended_cost=decimal(get(values, "extended cost")),
                        extended_profit=decimal(get(values, "extended profit")),
                        net_price=net_price,
                        category=str(get(values, "category") or "").strip(),
                    )
                )
    if not parsed:
        raise BidBuilderError(f"{path.name}: no export rows were found")
    return parsed, warnings


def validate_row(row: ExportRow) -> list[str]:
    checks = [
        (row.quantity * row.net_cost, row.extended_cost, "quantity × net cost", "extended cost"),
        (row.quantity * row.unit_profit, row.extended_profit, "quantity × unit profit", "extended profit"),
        (row.net_cost + row.unit_profit, row.unit_price, "net cost + unit profit", "unit price"),
        (row.extended_cost + row.extended_profit, row.net_price, "extended cost + extended profit", "net price"),
    ]
    issues = []
    for calculated, stated, formula, field in checks:
        if abs(money(calculated) - money(stated)) > TOLERANCE:
            issues.append(
                f"{row.source}/{row.sheet} row {row.row_number}: {formula} = ${money(calculated):,.2f}, "
                f"but {field} is ${money(stated):,.2f}"
            )
    return issues


class Catalog:
    def __init__(self, path: Path):
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.rules = payload.get("rules", [])

    def resolve(self, row: ExportRow) -> dict[str, Any]:
        public = row.type.strip().lower() not in INTERNAL_TYPES
        result = {"public": public, "manufacturer": None, "model": None, "rule_id": None}
        exact_rules = []
        regex_rules = []
        for rule in self.rules:
            match = rule.get("match", {})
            if match.get("item"):
                exact_rules.append(rule)
            else:
                regex_rules.append(rule)
        for rule in [*exact_rules, *regex_rules]:
            match = rule.get("match", {})
            item_match = not match.get("item") or row.item.casefold() == str(match["item"]).casefold()
            regex_match = not match.get("description_regex") or re.search(
                str(match["description_regex"]), row.description
            )
            if item_match and regex_match:
                result.update(
                    public=bool(rule.get("public", public)),
                    manufacturer=rule.get("manufacturer"),
                    model=rule.get("model") or row.item,
                    rule_id=rule.get("id"),
                )
                return result
        return result


def resolve_path(raw: str, job_path: Path) -> Path:
    path = Path(raw).expanduser()
    return path if path.is_absolute() else (job_path.parent / path).resolve()


def required_job_blockers(job: dict[str, Any]) -> list[str]:
    blockers = []
    for field in ("naming_line", "project", "proposal_number", "customer", "location"):
        if not str(job.get(field, "")).strip():
            blockers.append(f"Missing required job field: {field}")
    if job.get("grouping_confirmed") is not True:
        blockers.append("Spreadsheet grouping has not been explicitly confirmed")
    for field in ("proposal_date", "plan_date", "prepared_by"):
        if not str(job.get(field) or "").strip():
            blockers.append(f"Missing proposal fact: {field}")
    if not job.get("clauses_approved", False):
        blockers.append("Selected proposal clauses have not been approved by the company")
    if not job.get("sections"):
        blockers.append("No proposal sections were supplied")
    return blockers


def build_model(job: dict[str, Any], job_path: Path, catalog_path: Path, clauses_path: Path) -> dict[str, Any]:
    catalog = Catalog(catalog_path)
    clauses_payload = json.loads(clauses_path.read_text(encoding="utf-8"))
    profile_name = job.get("clauses_profile", "general_draft")
    profile = clauses_payload.get("profiles", {}).get(profile_name)
    blockers = required_job_blockers(job)
    if profile is None:
        blockers.append(f"Unknown clause profile: {profile_name}")
        profile = []

    sections = []
    grand_total = Decimal("0")
    seen_source_paths: set[str] = set()
    for section_index, section in enumerate(job.get("sections", []), start=1):
        title = str(section.get("title") or f"Section {section_index}")
        kind = str(section.get("kind") or "other").lower()
        if kind == "partitions" and not str(section.get("scope_summary") or "").strip():
            blockers.append(f"{title}: partition scope summary is required")

        all_rows: list[ExportRow] = []
        section_warnings: list[str] = []
        source_files = []
        if not section.get("source_files"):
            blockers.append(f"{title}: no source files were supplied")
        for raw_path in section.get("source_files", []):
            path = resolve_path(str(raw_path), job_path)
            normalized_source = str(path)
            if normalized_source in seen_source_paths:
                blockers.append(f"Workbook was assigned more than once: {path.name}")
            seen_source_paths.add(normalized_source)
            rows, warnings = read_export(path)
            all_rows.extend(rows)
            section_warnings.extend(warnings)
            source_files.append(str(path))

        arithmetic_issues = [issue for row in all_rows for issue in validate_row(row)]
        blockers.extend(arithmetic_issues)
        source_total = money(sum((row.net_price for row in all_rows), Decimal("0")))
        grand_total += source_total

        products: OrderedDict[tuple[str, str, str], dict[str, Any]] = OrderedDict()
        visible_value = Decimal("0")
        unresolved: list[str] = []
        if kind != "partitions":
            for row in all_rows:
                resolved = catalog.resolve(row)
                if not resolved["public"]:
                    continue
                manufacturer = resolved.get("manufacturer") or "UNRESOLVED"
                model = resolved.get("model") or row.item
                if manufacturer == "UNRESOLVED":
                    unresolved.append(
                        f"{title}: manufacturer unresolved for {row.item or '(no item)'} — {row.description}"
                    )
                key = (manufacturer, str(model), row.description)
                if key not in products:
                    products[key] = {
                        "manufacturer": manufacturer,
                        "model": model,
                        "description": row.description,
                        "quantity": Decimal("0"),
                        "source_value": Decimal("0"),
                    }
                products[key]["quantity"] += row.quantity
                products[key]["source_value"] += row.net_price
                visible_value += row.net_price
        blockers.extend(dict.fromkeys(unresolved))
        public_products = []
        for product in products.values():
            public_products.append(
                {
                    **product,
                    "quantity": quantity_text(product["quantity"]),
                    "source_value": str(money(product["source_value"])),
                }
            )

        hidden_allowance = money(source_total - visible_value)
        sections.append(
            {
                "title": title,
                "kind": kind,
                "scope_summary": section.get("scope_summary"),
                "source_files": source_files,
                "source_row_count": len(all_rows),
                "source_total": str(source_total),
                "public_source_value": str(money(visible_value)),
                "hidden_allowance": str(hidden_allowance),
                "products": public_products,
                "warnings": section_warnings,
                "arithmetic_issues": arithmetic_issues,
            }
        )

    unique_blockers = list(dict.fromkeys(blockers))
    return {
        "schema_version": 1,
        "generated_on": date.today().isoformat(),
        "client_ready": not unique_blockers,
        "draft_reason": None if not unique_blockers else "Resolve all blockers before issue.",
        "job": {
            key: job.get(key)
            for key in (
                "naming_line",
                "project",
                "proposal_number",
                "customer",
                "location",
                "proposal_date",
                "plan_date",
                "prepared_by",
            )
        },
        "sections": sections,
        "grand_total": str(money(grand_total)),
        "clause_profile": profile_name,
        "clauses": profile,
        "blockers": unique_blockers,
    }


def write_reconciliation(model: dict[str, Any], path: Path) -> None:
    lines = [
        f"# Reconciliation — {model['job'].get('proposal_number') or 'unassigned'}",
        "",
        f"**Status:** {'CLIENT READY' if model['client_ready'] else 'DRAFT — NOT FOR ISSUE'}",
        "",
        f"**Naming line:** {model['job'].get('naming_line') or 'MISSING'}",
        "",
    ]
    for section in model["sections"]:
        lines.extend(
            [
                f"## {section['title']}",
                "",
                f"- Source files: {', '.join(Path(p).name for p in section['source_files'])}",
                f"- Parsed rows: {section['source_row_count']}",
                f"- Public product value: ${Decimal(section['public_source_value']):,.2f}",
                f"- Hidden allowance: ${Decimal(section['hidden_allowance']):,.2f}",
                f"- Source section total: ${Decimal(section['source_total']):,.2f}",
                "",
            ]
        )
        if section["arithmetic_issues"]:
            lines.append("Arithmetic issues:")
            lines.extend(f"- {issue}" for issue in section["arithmetic_issues"])
            lines.append("")
        if section["warnings"]:
            lines.append("Warnings:")
            lines.extend(f"- {warning}" for warning in section["warnings"])
            lines.append("")
    lines.extend(["## Grand total", "", f"**${Decimal(model['grand_total']):,.2f}**", ""])
    lines.append("## Blockers")
    lines.append("")
    lines.extend(f"- {item}" for item in model["blockers"]) if model["blockers"] else lines.append("- None")
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def _set_cell_shading(cell: Any, fill: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shading)


def write_docx(model: dict[str, Any], path: Path, template_path: Path | None = None) -> None:
    try:
        from docx import Document
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Inches, Pt, RGBColor
    except ImportError as exc:
        raise BidBuilderError("DOCX generation requires python-docx: pip install -r requirements.txt") from exc

    doc = Document(str(template_path)) if template_path and template_path.exists() else Document()
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.72)
    section.right_margin = Inches(0.72)

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(9.5)
    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.size = Pt(25)
    styles["Title"].font.color.rgb = RGBColor(23, 50, 77)
    styles["Heading 1"].font.name = "Aptos Display"
    styles["Heading 1"].font.size = Pt(15)
    styles["Heading 1"].font.color.rgb = RGBColor(23, 50, 77)

    header = section.header.paragraphs[0]
    header.text = "PARTITIONS & ACCESSORIES CO.  /  COMMERCIAL PROPOSAL"
    header.runs[0].font.name = "Aptos"
    header.runs[0].font.size = Pt(8)
    header.runs[0].font.bold = True
    header.runs[0].font.color.rgb = RGBColor(216, 138, 61)

    footer = section.footer.paragraphs[0]
    footer.text = ""
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("1220 South Pasadena, Mesa, AZ 85210  •  480-969-6606  •  ")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)

    if not model["client_ready"]:
        banner = doc.add_table(rows=1, cols=1)
        banner.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = banner.cell(0, 0)
        _set_cell_shading(cell, "F6D6D6")
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run("DRAFT — NOT FOR ISSUE")
        run.bold = True
        run.font.color.rgb = RGBColor(143, 33, 33)
        run.font.size = Pt(11)
        note = doc.add_paragraph("Resolve the issues listed in reconciliation.md before customer issue.")
        note.alignment = WD_ALIGN_PARAGRAPH.CENTER
        note.runs[0].italic = True
        note.runs[0].font.size = Pt(8)

    title = doc.add_paragraph(style="Title")
    title.add_run(str(model["job"].get("project") or "PROPOSAL"))
    subtitle = doc.add_paragraph()
    subtitle.add_run(f"Proposal {model['job'].get('proposal_number') or 'UNASSIGNED'}").bold = True
    subtitle.add_run(f"  •  {model['job'].get('location') or 'Location pending'}")

    info = doc.add_table(rows=4, cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info.style = "Table Grid"
    facts = [
        ("Prepared for", model["job"].get("customer") or "Pending"),
        ("Proposal date", model["job"].get("proposal_date") or "Pending"),
        ("Plan date", model["job"].get("plan_date") or "Pending"),
        ("Prepared by", model["job"].get("prepared_by") or "Pending"),
    ]
    for row, (label, value) in zip(info.rows, facts):
        row.cells[0].text = label
        row.cells[1].text = str(value)
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[0].width = Inches(1.6)

    doc.add_paragraph()
    for section_model in model["sections"]:
        doc.add_heading(section_model["title"], level=1)
        if section_model.get("scope_summary"):
            scope = doc.add_paragraph()
            scope.add_run(str(section_model["scope_summary"])).bold = True
        if section_model["products"]:
            table = doc.add_table(rows=1, cols=4)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = "Table Grid"
            headers = ["Qty", "Manufacturer", "Model", "Description"]
            for cell, label in zip(table.rows[0].cells, headers):
                cell.text = label
                _set_cell_shading(cell, "17324D")
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for run in cell.paragraphs[0].runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
            for product in section_model["products"]:
                cells = table.add_row().cells
                values = [
                    product["quantity"],
                    product["manufacturer"],
                    product["model"],
                    product["description"],
                ]
                for cell, value in zip(cells, values):
                    cell.text = str(value or "")
        amount = doc.add_paragraph()
        amount.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        amount.add_run(f"SECTION TOTAL  ${Decimal(section_model['source_total']):,.2f}").bold = True

    total_table = doc.add_table(rows=1, cols=2)
    total_table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    total_table.style = "Table Grid"
    total_table.cell(0, 0).text = "PROPOSAL TOTAL"
    total_table.cell(0, 1).text = f"${Decimal(model['grand_total']):,.2f}"
    for cell in total_table.rows[0].cells:
        _set_cell_shading(cell, "D88A3D")
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    doc.add_heading("Commercial terms", level=1)
    for clause in model.get("clauses", []):
        doc.add_paragraph(str(clause), style="List Bullet")

    doc.save(path)


def convert_pdf(docx_path: Path, output_dir: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    completed = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(docx_path)],
        text=True,
        capture_output=True,
        check=False,
    )
    pdf_path = output_dir / f"{docx_path.stem}.pdf"
    if completed.returncode != 0 or not pdf_path.exists():
        raise BidBuilderError(f"LibreOffice PDF conversion failed: {completed.stderr or completed.stdout}")
    return pdf_path


def command_inspect(paths: Iterable[str]) -> int:
    for raw in paths:
        path = Path(raw).expanduser().resolve()
        rows, warnings = read_export(path)
        issues = [issue for row in rows for issue in validate_row(row)]
        total = money(sum((row.net_price for row in rows), Decimal("0")))
        print(f"{path.name}: {len(rows)} rows, ${total:,.2f}")
        for message in [*warnings, *issues]:
            print(f"  - {message}")
    return 0


def command_build(job_arg: str, output_arg: str, no_pdf: bool) -> int:
    job_path = Path(job_arg).expanduser().resolve()
    output_dir = Path(output_arg).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    job = json.loads(job_path.read_text(encoding="utf-8"))
    skill_root = Path(__file__).resolve().parent.parent
    model = build_model(
        job,
        job_path,
        skill_root / "references/item-catalog.json",
        skill_root / "references/clauses.json",
    )
    model_path = output_dir / "proposal-model.json"
    model_path.write_text(json.dumps(model, indent=2) + "\n", encoding="utf-8")
    write_reconciliation(model, output_dir / "reconciliation.md")
    docx_path = output_dir / "proposal.docx"
    write_docx(model, docx_path, skill_root / "assets/proposal-template.docx")
    if not no_pdf:
        convert_pdf(docx_path, output_dir)
    print(f"Built {output_dir}")
    print(f"Status: {'CLIENT READY' if model['client_ready'] else 'DRAFT — NOT FOR ISSUE'}")
    print(f"Total: ${Decimal(model['grand_total']):,.2f}")
    if model["blockers"]:
        print(f"Blockers: {len(model['blockers'])}")
    return 0 if model["client_ready"] else 2


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(prog="bid-builder")
    subparsers = root.add_subparsers(dest="command", required=True)
    inspect_parser = subparsers.add_parser("inspect", help="Inspect and total CAD-exported XLSX files")
    inspect_parser.add_argument("workbooks", nargs="+")
    build_parser = subparsers.add_parser("build", help="Build a proposal package from a job JSON file")
    build_parser.add_argument("--job", required=True)
    build_parser.add_argument("--output-dir", required=True)
    build_parser.add_argument("--no-pdf", action="store_true")
    return root


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.command == "inspect":
            return command_inspect(args.workbooks)
        return command_build(args.job, args.output_dir, args.no_pdf)
    except (BidBuilderError, json.JSONDecodeError, OSError, zipfile.BadZipFile) as exc:
        print(f"Bid Builder error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
