"""
Handbook / Guide template.

Generates informational handbooks with themed sections, chain of command,
code definitions, rules, and reference links — matching the format used
by JDU Information docs and battalion handbooks.
"""

from typing import List, Optional
from ..styles import StyleConfig
from .base import BaseTemplate


class HandbookTemplate(BaseTemplate):
    """Handbook / guide document template."""

    def __init__(self, style: Optional[StyleConfig] = None,
                 theme: str = "327th"):
        super().__init__(style=style, theme=theme)
        self.title = "Handbook"
        self.subtitle = ""
        self.important_links: list = []
        self.sections: list = []
        self.chain_of_command: list = []

    def add_link(self, label: str, description: str = ""):
        self.important_links.append({
            "label": label, "description": description
        })

    def add_section(self, title: str, content: list = None):
        """Add a handbook section.

        content is a list of dicts:
          - type: "text", "bullet_list", "numbered_list", "code_block",
                  "table", "sub_heading", "note", "warning", "divider"
          - text/items/headers/rows: the content
        """
        self.sections.append({"title": title, "content": content or []})

    def set_chain_of_command(self, chain: list):
        """Set the chain of command list (top to bottom)."""
        self.chain_of_command = chain

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

        # Build TOC
        toc_entries = []
        if self.important_links:
            toc_entries.append({"title": "Important Links"})
        for sec in self.sections:
            toc_entries.append({"title": sec["title"]})
        if self.chain_of_command:
            toc_entries.append({"title": "Chain of Command"})
        e.add_table_of_contents(toc_entries)

        # Important links
        if self.important_links:
            e.add_heading("Important Links")
            for link in self.important_links:
                text = link["label"]
                if link["description"]:
                    text += f" - {link['description']}"
                e.add_paragraph(
                    f"{self.style.section_symbol}  {text}"
                    if self.style.use_section_symbols
                    else text,
                    color_key=self.style.accent_color,
                    bold=True,
                )
            e.add_spacer()

        # Main sections
        for sec in self.sections:
            e.add_heading(sec["title"])
            e.add_metadata_line(
                author=self.author,
                formatted_by=self.formatted_by,
            )

            for block in sec.get("content", []):
                btype = block.get("type", "text")

                if btype == "text":
                    e.add_paragraph(
                        block.get("text", ""),
                        color_key=block.get("color", None),
                        bold=block.get("bold", False),
                        italic=block.get("italic", False),
                        indent=block.get("indent", 0),
                    )
                elif btype == "sub_heading":
                    e.add_heading(
                        block.get("text", ""),
                        level=2,
                    )
                elif btype == "code_block":
                    # Code definitions (e.g., Temple Codes, Defcons)
                    code_name = block.get("name", "")
                    code_color = block.get("color", None)
                    code_desc = block.get("description", "")
                    details = block.get("details", "")

                    p = e.add_paragraph(code_name, color_key=code_color,
                                        bold=True, alignment="center")
                    if code_desc:
                        e.add_paragraph(code_desc, italic=True,
                                        alignment="center")
                    if details:
                        e.add_paragraph(details, indent=0.25)
                    e.add_spacer()

                elif btype == "bullet_list":
                    e.add_bullet_list(
                        block.get("items", []),
                        indent=block.get("indent", 0.25),
                        color_key=block.get("color", None),
                    )
                elif btype == "numbered_list":
                    e.add_numbered_list(
                        block.get("items", []),
                        indent=block.get("indent", 0.25),
                    )
                elif btype == "table":
                    e.add_table(
                        headers=block.get("headers", []),
                        rows=block.get("rows", []),
                        col_widths=block.get("col_widths", None),
                    )
                elif btype == "note":
                    e.add_callout_box(block.get("text", ""), "note")
                elif btype == "warning":
                    e.add_callout_box(block.get("text", ""), "warning")
                elif btype == "divider":
                    e.add_divider()

            e.add_spacer()

        # Chain of command
        if self.chain_of_command:
            e.add_page_break()
            e.add_heading("Chain of Command")
            e.add_chain_of_command(self.chain_of_command)
