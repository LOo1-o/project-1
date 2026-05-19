"""Centralized DOCX table access utilities.

All table traversal should go through TableManager instead of ``doc.tables`` so
nested tables are discovered consistently via XPath.
"""

from dataclasses import dataclass
from hashlib import md5
from typing import List

from docx.table import Table


@dataclass(frozen=True)
class TableSignature:
    """Stable table signature used as a fallback when indexes shift."""

    rows: int
    columns: int
    preview_text: str
    hash: str


class TableManager:
    """Single access layer for DOCX tables, including nested tables."""

    def __init__(self, document):
        self.document = document
        self._tables: List[Table] = None
        self._signatures: List[TableSignature] = None

    def iter_tables(self) -> List[Table]:
        """Return all tables in document order, including nested tables."""
        if self._tables is None:
            self._tables = [Table(tbl, self.document) for tbl in self.document.element.body.xpath('.//w:tbl')]
        return self._tables

    def get_signature(self, table: Table) -> TableSignature:
        row_count = len(table.rows)
        column_count = max((len(row.cells) for row in table.rows), default=0)
        preview_parts = []
        for row in table.rows[:3]:
            for cell in row.cells[:3]:
                text = " ".join(cell.text.split())
                if text:
                    preview_parts.append(text)
        preview_text = " | ".join(preview_parts)[:300]
        digest_source = f"{row_count}x{column_count}|{preview_text}"
        return TableSignature(row_count, column_count, preview_text, md5(digest_source.encode('utf-8')).hexdigest())

    def print_table_map(self):
        """Debug helper: print shape, preview and hash for each discovered table."""
        for idx, table in enumerate(self.iter_tables(), start=1):
            signature = self.get_signature(table)
            print(
                f"#{idx}: {signature.rows}x{signature.columns} "
                f"hash={signature.hash} preview={signature.preview_text[:120]}"
            )
