"""
YAML configuration loader for document generation.

Reads a YAML file describing the document structure, style overrides,
and content, then drives the DocumentEngine to produce the output.
"""

import os
from datetime import datetime
from typing import Any, Dict, Optional

import yaml

from .styles import StyleConfig, THEMES
from .engine import DocumentEngine


def load_yaml(filepath: str) -> dict:
    """Load and return a YAML config file."""
    with open(filepath, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _apply_style_overrides(style: StyleConfig, overrides: dict) -> StyleConfig:
    """Apply a dict of overrides onto a StyleConfig."""
    for key, value in overrides.items():
        if hasattr(style, key):
            setattr(style, key, value)
    return style


def build_style_from_config(config: dict) -> StyleConfig:
    """Build a StyleConfig from the 'style' section of a config."""
    style_cfg = config.get("style", {})

    # Start from a theme preset or default
    theme_name = style_cfg.get("theme", "327th")
    if theme_name in THEMES:
        style = StyleConfig(**THEMES[theme_name].__dict__.copy())
    else:
        style = StyleConfig()

    # Apply any explicit overrides
    overrides = {k: v for k, v in style_cfg.items() if k != "theme"}
    return _apply_style_overrides(style, overrides)


def _process_content_block(engine: DocumentEngine, block: dict):
    """Process a single content block from the YAML config."""
    block_type = block.get("type", "paragraph")

    if block_type == "title_page":
        engine.add_title_page(
            title=block.get("title", "Untitled Document"),
            subtitle=block.get("subtitle", ""),
            author=block.get("author", ""),
            formatted_by=block.get("formatted_by", ""),
            version_date=block.get("version_date",
                                   datetime.now().strftime("%m/%d/%Y")),
            extra_lines=block.get("extra_lines", []),
        )

    elif block_type == "table_of_contents":
        entries = block.get("entries", None)
        engine.add_table_of_contents(entries)

    elif block_type == "heading":
        engine.add_heading(
            text=block.get("text", ""),
            level=block.get("level", 1),
        )

    elif block_type == "paragraph":
        engine.add_paragraph(
            text=block.get("text", ""),
            color_key=block.get("color", None),
            bold=block.get("bold", False),
            italic=block.get("italic", False),
            alignment=block.get("alignment", "left"),
            indent=block.get("indent", 0),
        )

    elif block_type == "read_aloud":
        engine.add_read_aloud(block.get("text", ""))

    elif block_type == "host_info":
        engine.add_host_info(block.get("text", ""))

    elif block_type == "important_info":
        engine.add_important_info(block.get("text", ""))

    elif block_type == "bullet_list":
        engine.add_bullet_list(
            items=block.get("items", []),
            indent=block.get("indent", 0.25),
            color_key=block.get("color", None),
        )

    elif block_type == "numbered_list":
        engine.add_numbered_list(
            items=block.get("items", []),
            indent=block.get("indent", 0.25),
            color_key=block.get("color", None),
            start_num=block.get("start_num", 1),
        )

    elif block_type == "lettered_list":
        engine.add_lettered_sub_list(
            items=block.get("items", []),
            indent=block.get("indent", 0.5),
            color_key=block.get("color", None),
        )

    elif block_type == "qa_block":
        engine.add_qa_block(
            question=block.get("question", ""),
            answer=block.get("answer", ""),
            q_label=block.get("q_label", "Q"),
            a_label=block.get("a_label", "A"),
        )

    elif block_type == "table":
        engine.add_table(
            headers=block.get("headers", []),
            rows=block.get("rows", []),
            col_widths=block.get("col_widths", None),
        )

    elif block_type == "chain_of_command":
        engine.add_chain_of_command(block.get("chain", []))

    elif block_type == "color_code_legend":
        engine.add_color_code_legend()

    elif block_type == "callout":
        engine.add_callout_box(
            text=block.get("text", ""),
            style_type=block.get("callout_style", "info"),
        )

    elif block_type == "metadata_line":
        engine.add_metadata_line(
            author=block.get("author", ""),
            formatted_by=block.get("formatted_by", ""),
            created=block.get("created", ""),
            updated=block.get("updated", ""),
            alignment=block.get("alignment", "right"),
        )

    elif block_type == "divider":
        engine.add_divider()

    elif block_type == "spacer":
        engine.add_spacer(lines=block.get("lines", 1))

    elif block_type == "page_break":
        engine.add_page_break()


def generate_from_config(config_path: str, output_path: str = None,
                         fmt: str = None) -> str:
    """Generate a document from a YAML configuration file.

    Args:
        config_path: Path to the YAML config file.
        output_path: Output file path. If None, derived from config.
        fmt: Output format ('docx' or 'pdf'). If None, derived from output
             path extension or config.

    Returns:
        Path to the generated file.
    """
    config = load_yaml(config_path)

    # Determine output path
    if output_path is None:
        output_cfg = config.get("output", {})
        output_path = output_cfg.get(
            "path",
            os.path.join("output", config.get("document", {}).get(
                "title", "document").replace(" ", "_") + ".docx")
        )

    # Determine format
    if fmt is None:
        if output_path.lower().endswith(".pdf"):
            fmt = "pdf"
        else:
            fmt = config.get("output", {}).get("format", "docx")

    # Ensure correct extension
    base = output_path.rsplit(".", 1)[0] if "." in output_path else output_path
    output_path = f"{base}.{fmt}"

    # Build style
    style = build_style_from_config(config)

    # Create engine and build document
    engine = DocumentEngine(style)

    # Process content blocks
    content = config.get("content", [])
    for block in content:
        _process_content_block(engine, block)

    # Save
    return engine.save(output_path, fmt=fmt)
