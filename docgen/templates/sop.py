"""
Standard Operating Procedure (SOP) template.

Generates structured SOP documents with numbered sections, procedures,
responsibility assignments, and revision history.
"""

from typing import List, Optional
from ..styles import StyleConfig
from .base import BaseTemplate


class SOPTemplate(BaseTemplate):
    """Standard Operating Procedure document template."""

    def __init__(self, style: Optional[StyleConfig] = None,
                 theme: str = "327th"):
        super().__init__(style=style, theme=theme)
        self.title = "Standard Operating Procedure"
        self.subtitle = "327th Star Corps"
        self.sections: list = []
        self.purpose = ""
        self.scope = ""
        self.references: list = []
        self.revision_history: list = []

    def set_purpose(self, purpose: str):
        self.purpose = purpose

    def set_scope(self, scope: str):
        self.scope = scope

    def add_reference(self, reference: str):
        self.references.append(reference)

    def add_revision(self, date: str, version: str, description: str,
                     author: str = ""):
        self.revision_history.append({
            "date": date,
            "version": version,
            "description": description,
            "author": author,
        })

    def add_section(self, title: str, content: list = None):
        """Add an SOP section.

        content is a list of dicts, each with:
          - type: "text", "steps", "bullet_list", "note", "warning"
          - text/items: the content
        """
        self.sections.append({"title": title, "content": content or []})

    def build(self):
        e = self.engine

        # Title page
        e.add_title_page(
            title=self.title,
            subtitle=self.subtitle,
            author=self.author,
            formatted_by=self.formatted_by,
            version_date=self.version_date,
            extra_lines=[self.unit, self.company],
        )

        # Table of contents
        toc_entries = []
        if self.purpose:
            toc_entries.append({"title": "Purpose"})
        if self.scope:
            toc_entries.append({"title": "Scope"})
        if self.references:
            toc_entries.append({"title": "References"})
        for sec in self.sections:
            toc_entries.append({"title": sec["title"]})
        if self.revision_history:
            toc_entries.append({"title": "Revision History"})
        e.add_table_of_contents(toc_entries)

        # Purpose
        if self.purpose:
            e.add_heading("Purpose")
            e.add_paragraph(self.purpose)
            e.add_spacer()

        # Scope
        if self.scope:
            e.add_heading("Scope")
            e.add_paragraph(self.scope)
            e.add_spacer()

        # References
        if self.references:
            e.add_heading("References")
            e.add_bullet_list(self.references)
            e.add_spacer()

        # Main sections
        for sec in self.sections:
            e.add_heading(sec["title"])

            for block in sec.get("content", []):
                btype = block.get("type", "text")
                if btype == "text":
                    e.add_paragraph(block.get("text", ""))
                elif btype == "steps":
                    e.add_numbered_list(block.get("items", []))
                elif btype == "sub_steps":
                    e.add_lettered_sub_list(block.get("items", []))
                elif btype == "bullet_list":
                    e.add_bullet_list(block.get("items", []))
                elif btype == "note":
                    e.add_callout_box(block.get("text", ""), "note")
                elif btype == "warning":
                    e.add_callout_box(block.get("text", ""), "warning")
                elif btype == "important":
                    e.add_important_info(block.get("text", ""))
                elif btype == "divider":
                    e.add_divider()

            e.add_spacer()

        # Revision history
        if self.revision_history:
            e.add_page_break()
            e.add_heading("Revision History")
            headers = ["Date", "Version", "Description", "Author"]
            rows = [
                [r["date"], r["version"], r["description"], r["author"]]
                for r in self.revision_history
            ]
            e.add_table(headers, rows, col_widths=[1.2, 0.8, 3.0, 1.5])
