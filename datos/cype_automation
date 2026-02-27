#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import re
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional

import yaml
from docx import Document
from docx.document import Document as DocxDocumentType
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph


@dataclass
class SectionRule:
    id: str
    source_trigger_regex: str
    table_offset_after_trigger: int
    target_placeholder: str


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
        rules.append(
            SectionRule(
                id=section.get("id", f"section_{i}"),
                source_trigger_regex=section["source_trigger_regex"],
                table_offset_after_trigger=int(section.get("table_offset_after_trigger", 1)),
                target_placeholder=section["target_placeholder"],
            )
        )
    return rules


def find_table_after_trigger(source_doc: DocxDocumentType, regex: str, offset: int) -> Optional[Table]:
    trigger_pattern = re.compile(regex)
    seen_trigger = False
    tables_after_trigger = 0

    for block in iter_block_items(source_doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text and trigger_pattern.search(text):
                seen_trigger = True
                tables_after_trigger = 0
                continue

        if seen_trigger and isinstance(block, Table):
            tables_after_trigger += 1
            if tables_after_trigger == offset:
                return block

    return None


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


def run(source_docx: Path, template_docx: Path, mapping_yaml: Path, output_docx: Path, output_pdf: bool) -> None:
    rules = load_rules(mapping_yaml)
    if not rules:
        raise ValueError("No se encontraron secciones en el archivo de mapeo YAML.")

    source_doc = Document(str(source_docx))
    template_doc = Document(str(template_docx))

    print(f"Reglas cargadas: {len(rules)}")
    for rule in rules:
        table = find_table_after_trigger(source_doc, rule.source_trigger_regex, rule.table_offset_after_trigger)
        if table is None:
            print(f"[WARN] {rule.id}: no se encontró tabla para regex={rule.source_trigger_regex!r}")
            continue

        inserted = replace_placeholder_with_table(template_doc, rule.target_placeholder, table)
        if not inserted:
            print(f"[WARN] {rule.id}: no se encontró marcador {rule.target_placeholder!r} en plantilla")
            continue

        print(f"[OK] {rule.id}: tabla insertada en {rule.target_placeholder}")

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    template_doc.save(str(output_docx))
    print(f"Archivo generado: {output_docx}")

    if output_pdf:
        pdf_path = convert_to_pdf(output_docx)
        print(f"PDF generado: {pdf_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automatiza el pegado de tablas desde memoria CYPECAD a plantilla de memoria propia."
    )
    parser.add_argument("--source-docx", required=True, type=Path, help="DOCX exportado desde CYPECAD")
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
        template_docx=args.template_docx,
        mapping_yaml=args.mapping_yaml,
        output_docx=args.output_docx,
        output_pdf=args.output_pdf,
    )


if __name__ == "__main__":
    main()
