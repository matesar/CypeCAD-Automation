#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import re
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

import yaml
from docx import Document
from docx.document import Document as DocxDocumentType
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


@dataclass
class SectionRule:
    id: str
    target_placeholder: str
    source_trigger_regex: Optional[str] = None
    table_header_regex: Optional[str] = None
    table_offset_after_trigger: int = 1


def iter_block_items(document: DocxDocumentType) -> Iterable[Paragraph | Table]:
    """Yield paragraphs and tables in document order."""
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def load_rules(mapping_path: Path) -> List[SectionRule]:
    with mapping_path.open("r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}

    sections = raw.get("sections", [])
    rules: List[SectionRule] = []
    for i, section in enumerate(sections, start=1):
        source_trigger_regex = section.get("source_trigger_regex")
        table_header_regex = section.get("table_header_regex")
        if not source_trigger_regex and not table_header_regex:
            raise ValueError(
                f"La sección {section.get('id', f'section_{i}')!r} debe tener al menos "
                "'source_trigger_regex' o 'table_header_regex'."
            )

        rules.append(
            SectionRule(
                id=section.get("id", f"section_{i}"),
                source_trigger_regex=source_trigger_regex,
                table_header_regex=table_header_regex,
                table_offset_after_trigger=int(section.get("table_offset_after_trigger", 1)),
                target_placeholder=section["target_placeholder"],
            )
        )
    return rules


def get_table_header_text(table: Table) -> str:
    """Return a normalized string from the first row of the table."""
    if not table.rows:
        return ""
    first_row = table.rows[0]
    values = [cell.text.strip() for cell in first_row.cells if cell.text.strip()]
    return " | ".join(values)


def matches_table_header(table: Table, header_regex: Optional[str]) -> bool:
    if not header_regex:
        return True
    header_text = get_table_header_text(table)
    return bool(header_text and re.search(header_regex, header_text))


def find_table_in_document(source_doc: DocxDocumentType, rule: SectionRule) -> Optional[Table]:
    """
    Selection rules:
    - If source_trigger_regex exists: start matching from that paragraph onward.
    - If table_header_regex exists: only tables whose first-row header matches are considered.
    - table_offset_after_trigger picks the Nth considered table.
    """
    tables_seen = 0

    if rule.source_trigger_regex:
        trigger_pattern = re.compile(rule.source_trigger_regex)
        seen_trigger = False

        for block in iter_block_items(source_doc):
            if isinstance(block, Paragraph):
                text = block.text.strip()
                if text and trigger_pattern.search(text):
                    seen_trigger = True
                    tables_seen = 0
                    continue

            if seen_trigger and isinstance(block, Table) and matches_table_header(block, rule.table_header_regex):
                tables_seen += 1
                if tables_seen == rule.table_offset_after_trigger:
                    return block
        return None

    for block in iter_block_items(source_doc):
        if isinstance(block, Table) and matches_table_header(block, rule.table_header_regex):
            tables_seen += 1
            if tables_seen == rule.table_offset_after_trigger:
                return block
    return None


def find_table_in_sources(source_docs: Sequence[tuple[Path, DocxDocumentType]], rule: SectionRule) -> tuple[Optional[Table], Optional[Path]]:
    for source_path, source_doc in source_docs:
        table = find_table_in_document(source_doc, rule)
        if table is not None:
            return table, source_path
    return None, None


def find_placeholder_paragraph(document: DocxDocumentType, placeholder: str) -> Optional[Paragraph]:
    for paragraph in document.paragraphs:
        if placeholder in paragraph.text:
            return paragraph
    return None


def replace_placeholder_with_table(template_doc: DocxDocumentType, placeholder: str, table_to_insert: Table) -> bool:
    paragraph = find_placeholder_paragraph(template_doc, placeholder)
    if paragraph is None:
        return False

    parent = paragraph._p.getparent()
    paragraph_index = parent.index(paragraph._p)

    new_table_xml = copy.deepcopy(table_to_insert._tbl)
    parent.insert(paragraph_index + 1, new_table_xml)

    paragraph.text = paragraph.text.replace(placeholder, "").strip()
    if not paragraph.text:
        parent.remove(paragraph._p)

    return True


def convert_to_pdf(docx_path: Path) -> Path:
    soffice = shutil.which("soffice")
    if not soffice:
        raise RuntimeError("No se encontró 'soffice' en PATH. Instalá LibreOffice para habilitar PDF.")

    out_dir = docx_path.parent
    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf",
        str(docx_path),
        "--outdir",
        str(out_dir),
    ]
    subprocess.run(cmd, check=True)

    pdf_path = out_dir / f"{docx_path.stem}.pdf"
    if not pdf_path.exists():
        raise RuntimeError("La conversión a PDF no generó el archivo esperado.")
    return pdf_path


def list_source_docx_files(source_docx: Optional[List[Path]], source_dir: Optional[Path]) -> List[Path]:
    source_files: List[Path] = []

    if source_docx:
        source_files.extend(source_docx)

    if source_dir:
        if not source_dir.exists() or not source_dir.is_dir():
            raise ValueError(f"La carpeta de fuentes no existe o no es válida: {source_dir}")
        source_files.extend(sorted(source_dir.glob("*.docx")))

    # remove duplicates while preserving order
    unique: List[Path] = []
    seen = set()
    for path in source_files:
        resolved = str(path.resolve())
        if resolved not in seen:
            seen.add(resolved)
            unique.append(path)

    if not unique:
        raise ValueError("No se encontraron DOCX de entrada. Usá --source-docx y/o --source-dir.")

    missing = [str(p) for p in unique if not p.exists()]
    if missing:
        raise ValueError(f"No existen estos archivos fuente: {', '.join(missing)}")

    return unique


def run(
    source_docx: Optional[List[Path]],
    source_dir: Optional[Path],
    template_docx: Path,
    mapping_yaml: Path,
    output_docx: Path,
    output_pdf: bool,
) -> None:
    rules = load_rules(mapping_yaml)
    if not rules:
        raise ValueError("No se encontraron secciones en el archivo de mapeo YAML.")

    source_files = list_source_docx_files(source_docx, source_dir)
    source_docs = [(source_file, Document(str(source_file))) for source_file in source_files]
    template_doc = Document(str(template_docx))

    print(f"Fuentes detectadas: {len(source_docs)}")
    for source_file, _ in source_docs:
        print(f"  - {source_file}")

    print(f"Reglas cargadas: {len(rules)}")
    for rule in rules:
        table, matched_source = find_table_in_sources(source_docs, rule)
        if table is None:
            print(
                f"[WARN] {rule.id}: no se encontró tabla en ninguna fuente "
                f"(trigger={rule.source_trigger_regex!r}, header={rule.table_header_regex!r}, "
                f"offset={rule.table_offset_after_trigger})"
            )
            continue

        inserted = replace_placeholder_with_table(template_doc, rule.target_placeholder, table)
        if not inserted:
            print(f"[WARN] {rule.id}: no se encontró marcador {rule.target_placeholder!r} en plantilla")
            continue

        print(f"[OK] {rule.id}: tabla insertada en {rule.target_placeholder} (fuente: {matched_source})")

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    template_doc.save(str(output_docx))
    print(f"Archivo generado: {output_docx}")

    if output_pdf:
        pdf_path = convert_to_pdf(output_docx)
        print(f"PDF generado: {pdf_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automatiza el pegado de tablas desde memorias DOCX de CYPECAD a plantilla de memoria propia."
    )
    parser.add_argument(
        "--source-docx",
        nargs="+",
        type=Path,
        help="Uno o más DOCX exportados desde CYPECAD (ej: --source-docx a.docx b.docx)",
    )
    parser.add_argument(
        "--source-dir",
        type=Path,
        help="Carpeta con múltiples DOCX fuente; se procesan todos los *.docx",
    )
    parser.add_argument("--template-docx", required=True, type=Path, help="Plantilla DOCX de memoria de cálculo")
    parser.add_argument("--mapping-yaml", required=True, type=Path, help="YAML con reglas de extracción/inserción")
    parser.add_argument("--output-docx", required=True, type=Path, help="Ruta del DOCX final")
    parser.add_argument(
        "--output-pdf",
        action="store_true",
        help="Convierte también el DOCX final a PDF con LibreOffice (soffice)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    run(
        source_docx=args.source_docx,
        source_dir=args.source_dir,
        template_docx=args.template_docx,
        mapping_yaml=args.mapping_yaml,
        output_docx=args.output_docx,
        output_pdf=args.output_pdf,
    )


if __name__ == "__main__":
    main()

if __name__ == "__main__":
    main()
