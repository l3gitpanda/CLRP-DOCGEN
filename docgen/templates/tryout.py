"""
Tryout Document template.

Generates structured tryout documents with phased sections, color-coded
instructions, Q&A blocks, and setup checklists — matching the format
used by 327th Star Corps and K Company tryout docs.
"""

from typing import List, Optional
from ..styles import StyleConfig
from .base import BaseTemplate


class TryoutTemplate(BaseTemplate):
    """Tryout document template with phases and color-coded instructions."""

    def __init__(self, style: Optional[StyleConfig] = None,
                 theme: str = "327th"):
        super().__init__(style=style, theme=theme)
        self.title = "Tryout Document"
        self.subtitle = ""
        self.introduction = ""
        self.setup_steps: list = []
        self.phases: list = []
        self.conclusion_steps: list = []
        self.show_color_legend = True
        self.strike_system = ""
        self.cooldown_info = ""

    def set_introduction(self, text: str):
        self.introduction = text

    def set_strike_system(self, text: str):
        self.strike_system = text

    def set_cooldown_info(self, text: str):
        self.cooldown_info = text

    def add_setup_step(self, text: str, sub_steps: list = None,
                       step_type: str = "normal"):
        """Add a setup/preparation step.

        step_type: 'normal', 'host_info', 'important', 'advert'
        """
        self.setup_steps.append({
            "text": text, "sub_steps": sub_steps or [], "type": step_type
        })

    def add_phase(self, title: str, content: list = None):
        """Add a tryout phase.

        content is a list of dicts, each with:
          - type: "text", "read_aloud", "host_info", "important",
                  "steps", "bullet_list", "qa", "note", "warning"
          - text/items/question/answer: the content
        """
        self.phases.append({"title": title, "content": content or []})

    def add_conclusion_step(self, text: str, sub_steps: list = None):
        self.conclusion_steps.append({
            "text": text, "sub_steps": sub_steps or []
        })

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
        if self.introduction:
            toc_entries.append({"title": "Introduction"})
        if self.show_color_legend:
            toc_entries.append({"title": "Color Codes"})
        if self.setup_steps:
            toc_entries.append({"title": "Setup / Preparation"})
        for phase in self.phases:
            toc_entries.append({"title": phase["title"]})
        if self.conclusion_steps:
            toc_entries.append({"title": "Conclusion"})
        e.add_table_of_contents(toc_entries)

        # Introduction
        if self.introduction:
            e.add_heading("Introduction")
            e.add_paragraph(self.introduction)

            if self.strike_system:
                e.add_spacer()
                e.add_callout_box(self.strike_system, "note")

            if self.cooldown_info:
                e.add_paragraph(self.cooldown_info, italic=True)

            e.add_spacer()

        # Color code legend
        if self.show_color_legend:
            e.add_color_code_legend()
            e.add_spacer()

        # Setup
        if self.setup_steps:
            e.add_heading("Setup / Preparation")
            for step in self.setup_steps:
                if step["type"] == "host_info":
                    e.add_host_info(step["text"])
                elif step["type"] == "important":
                    e.add_important_info(step["text"])
                elif step["type"] == "advert":
                    e.add_important_info(step["text"])
                else:
                    e.add_paragraph(step["text"])

                if step["sub_steps"]:
                    e.add_bullet_list(step["sub_steps"], indent=0.5)

            e.add_spacer()

        # Phases
        for i, phase in enumerate(self.phases):
            e.add_heading(phase["title"])
            e.add_metadata_line(
                author=self.author,
                formatted_by=self.formatted_by,
            )

            for block in phase.get("content", []):
                btype = block.get("type", "text")

                if btype == "text":
                    e.add_paragraph(
                        block.get("text", ""),
                        indent=block.get("indent", 0),
                    )
                elif btype == "read_aloud":
                    e.add_read_aloud(block.get("text", ""))
                elif btype == "host_info":
                    e.add_host_info(block.get("text", ""))
                elif btype == "important":
                    e.add_important_info(block.get("text", ""))
                elif btype == "steps":
                    e.add_numbered_list(block.get("items", []))
                elif btype == "sub_steps":
                    e.add_lettered_sub_list(block.get("items", []))
                elif btype == "bullet_list":
                    e.add_bullet_list(block.get("items", []))
                elif btype == "qa":
                    e.add_qa_block(
                        question=block.get("question", ""),
                        answer=block.get("answer", ""),
                        q_label=block.get("q_label", "Q"),
                        a_label=block.get("a_label", "A"),
                    )
                elif btype == "note":
                    e.add_callout_box(block.get("text", ""), "note")
                elif btype == "warning":
                    e.add_callout_box(block.get("text", ""), "warning")
                elif btype == "divider":
                    e.add_divider()
                elif btype == "table":
                    e.add_table(
                        headers=block.get("headers", []),
                        rows=block.get("rows", []),
                    )

            e.add_spacer()

        # Conclusion
        if self.conclusion_steps:
            e.add_page_break()
            e.add_heading("Conclusion")
            e.add_paragraph(
                "If the candidate has made it this far, they have passed. "
                "Congratulate them and go over the following:",
            )
            for step in self.conclusion_steps:
                e.add_paragraph(step["text"], bold=True)
                if step["sub_steps"]:
                    e.add_bullet_list(step["sub_steps"], indent=0.5)
