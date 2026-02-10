"""
Command-line interface for CLRP-DOCGEN.

Provides two modes:
  1. Config mode:  Generate from a YAML config file
  2. Interactive mode:  Walk through prompts to build a document
"""

import argparse
import os
import sys
from datetime import datetime

from . import __version__
from .styles import StyleConfig, THEMES
from .config_loader import generate_from_config
from .templates.sop import SOPTemplate
from .templates.tryout import TryoutTemplate
from .templates.handbook import HandbookTemplate


def _prompt(msg: str, default: str = "") -> str:
    """Prompt user for input with optional default."""
    suffix = f" [{default}]" if default else ""
    val = input(f"{msg}{suffix}: ").strip()
    return val if val else default


def _prompt_choice(msg: str, choices: list, default: str = "") -> str:
    """Prompt user to choose from a list."""
    print(f"\n{msg}")
    for i, c in enumerate(choices, 1):
        marker = " (default)" if c == default else ""
        print(f"  {i}. {c}{marker}")
    val = input("Choice [number or name]: ").strip()

    # Try as number
    try:
        idx = int(val) - 1
        if 0 <= idx < len(choices):
            return choices[idx]
    except ValueError:
        pass

    # Try as name
    if val in choices:
        return val

    return default if default else choices[0]


def _interactive_sop(args):
    """Interactive SOP document builder."""
    print("\n=== SOP Document Generator ===\n")

    theme = _prompt_choice(
        "Select theme:", list(THEMES.keys()), default="327th"
    )
    tmpl = SOPTemplate(theme=theme)

    tmpl.set_metadata(
        title=_prompt("Document title", "Standard Operating Procedure"),
        subtitle=_prompt("Subtitle", "327th Star Corps"),
        author=_prompt("Author name"),
        formatted_by=_prompt("Formatted by", ""),
        version_date=_prompt("Version date",
                             datetime.now().strftime("%m/%d/%Y")),
        unit=_prompt("Unit", "327th Star Corps"),
        company=_prompt("Company", "K Company"),
    )

    tmpl.set_purpose(_prompt("Purpose of this SOP"))
    tmpl.set_scope(_prompt("Scope of this SOP"))

    # References
    print("\nAdd references (blank line to stop):")
    while True:
        ref = input("  Reference: ").strip()
        if not ref:
            break
        tmpl.add_reference(ref)

    # Sections
    print("\nAdd sections (blank title to stop):")
    while True:
        sec_title = input("\nSection title: ").strip()
        if not sec_title:
            break

        content = []
        print(f"  Add content to '{sec_title}' (type 'done' to finish):")
        while True:
            btype = _prompt_choice(
                "  Block type:",
                ["text", "steps", "bullet_list", "note", "warning", "done"],
                default="text",
            )
            if btype == "done":
                break

            if btype == "text":
                text = input("    Text: ").strip()
                content.append({"type": "text", "text": text})
            elif btype in ("steps", "bullet_list"):
                items = []
                print(f"    Enter items (blank to stop):")
                while True:
                    item = input("      - ").strip()
                    if not item:
                        break
                    items.append(item)
                content.append({"type": btype, "items": items})
            elif btype in ("note", "warning"):
                text = input(f"    {btype.title()} text: ").strip()
                content.append({"type": btype, "text": text})

        tmpl.add_section(sec_title, content)

    # Revision history
    add_rev = _prompt("Add revision history? (y/n)", "n")
    if add_rev.lower() == "y":
        while True:
            date = input("  Revision date (blank to stop): ").strip()
            if not date:
                break
            ver = input("  Version: ").strip()
            desc = input("  Description: ").strip()
            auth = input("  Author: ").strip()
            tmpl.add_revision(date, ver, desc, auth)

    return tmpl


def _interactive_tryout(args):
    """Interactive tryout document builder."""
    print("\n=== Tryout Document Generator ===\n")

    theme = _prompt_choice(
        "Select theme:", list(THEMES.keys()), default="k_company"
    )
    tmpl = TryoutTemplate(theme=theme)

    tmpl.set_metadata(
        title=_prompt("Document title", "K Company Tryout Document"),
        subtitle=_prompt("Subtitle", ""),
        author=_prompt("Author name"),
        formatted_by=_prompt("Formatted by", ""),
        version_date=_prompt("Version date",
                             datetime.now().strftime("%m/%d/%Y")),
        unit=_prompt("Unit", "327th Star Corps"),
        company=_prompt("Company", "K Company"),
    )

    tmpl.set_introduction(
        _prompt("Introduction text",
                "This is the tryout document. Follow this document throughout "
                "the entirety of the tryout.")
    )
    tmpl.set_strike_system(
        _prompt("Strike system description",
                "This tryout runs on a 2 strike system. "
                "If you receive 2 strikes you will fail and be dismissed.")
    )
    tmpl.set_cooldown_info(
        _prompt("Cooldown info",
                "Should you fail, there is a 3 hour cooldown before retrying.")
    )

    # Setup steps
    print("\nSetup steps (blank to stop):")
    while True:
        step = input("  Step: ").strip()
        if not step:
            break
        stype = _prompt_choice(
            "  Step type:", ["normal", "host_info", "important", "advert"],
            default="normal",
        )
        tmpl.add_setup_step(step, step_type=stype)

    # Phases
    print("\nAdd tryout phases (blank title to stop):")
    while True:
        phase_title = input("\nPhase title: ").strip()
        if not phase_title:
            break

        content = []
        print(f"  Add content to '{phase_title}' (type 'done' to finish):")
        while True:
            btype = _prompt_choice(
                "  Block type:",
                ["text", "read_aloud", "host_info", "important",
                 "steps", "bullet_list", "qa", "note", "warning", "done"],
                default="text",
            )
            if btype == "done":
                break

            if btype in ("text", "read_aloud", "host_info", "important"):
                text = input("    Text: ").strip()
                content.append({"type": btype, "text": text})
            elif btype in ("steps", "bullet_list"):
                items = []
                print("    Enter items (blank to stop):")
                while True:
                    item = input("      - ").strip()
                    if not item:
                        break
                    items.append(item)
                content.append({"type": btype, "items": items})
            elif btype == "qa":
                q = input("    Question: ").strip()
                a = input("    Answer: ").strip()
                content.append({
                    "type": "qa", "question": q, "answer": a,
                })
            elif btype in ("note", "warning"):
                text = input(f"    {btype.title()} text: ").strip()
                content.append({"type": btype, "text": text})

        tmpl.add_phase(phase_title, content)

    # Conclusion
    print("\nConclusion steps (blank to stop):")
    while True:
        step = input("  Step: ").strip()
        if not step:
            break
        subs = []
        print("    Sub-steps (blank to stop):")
        while True:
            sub = input("      - ").strip()
            if not sub:
                break
            subs.append(sub)
        tmpl.add_conclusion_step(step, subs)

    return tmpl


def _interactive_handbook(args):
    """Interactive handbook/guide builder."""
    print("\n=== Handbook / Guide Generator ===\n")

    theme = _prompt_choice(
        "Select theme:", list(THEMES.keys()), default="jdu"
    )
    tmpl = HandbookTemplate(theme=theme)

    tmpl.set_metadata(
        title=_prompt("Document title", "K Company JDU Information"),
        subtitle=_prompt("Subtitle", ""),
        author=_prompt("Author name"),
        formatted_by=_prompt("Formatted by", ""),
        version_date=_prompt("Version date",
                             datetime.now().strftime("%m/%d/%Y")),
        unit=_prompt("Unit", "327th Star Corps"),
        company=_prompt("Company", "K Company"),
    )

    # Important links
    print("\nImportant links (blank to stop):")
    while True:
        label = input("  Link label: ").strip()
        if not label:
            break
        desc = input("  Description: ").strip()
        tmpl.add_link(label, desc)

    # Sections
    print("\nAdd sections (blank title to stop):")
    while True:
        sec_title = input("\nSection title: ").strip()
        if not sec_title:
            break

        content = []
        print(f"  Add content to '{sec_title}' (type 'done' to finish):")
        while True:
            btype = _prompt_choice(
                "  Block type:",
                ["text", "sub_heading", "code_block", "bullet_list",
                 "numbered_list", "table", "note", "warning", "done"],
                default="text",
            )
            if btype == "done":
                break

            if btype == "text":
                text = input("    Text: ").strip()
                content.append({"type": "text", "text": text})
            elif btype == "sub_heading":
                text = input("    Sub-heading: ").strip()
                content.append({"type": "sub_heading", "text": text})
            elif btype == "code_block":
                name = input("    Code/status name: ").strip()
                desc = input("    Description: ").strip()
                details = input("    Details: ").strip()
                color = _prompt("    Color key", "")
                content.append({
                    "type": "code_block", "name": name,
                    "description": desc, "details": details,
                    "color": color if color else None,
                })
            elif btype in ("bullet_list", "numbered_list"):
                items = []
                print("    Enter items (blank to stop):")
                while True:
                    item = input("      - ").strip()
                    if not item:
                        break
                    items.append(item)
                content.append({"type": btype, "items": items})
            elif btype == "table":
                hdrs = input("    Headers (comma-separated): ").strip()
                headers = [h.strip() for h in hdrs.split(",")]
                rows = []
                print("    Rows (blank to stop):")
                while True:
                    row = input("      Row (comma-separated): ").strip()
                    if not row:
                        break
                    rows.append([c.strip() for c in row.split(",")])
                content.append({
                    "type": "table", "headers": headers, "rows": rows,
                })
            elif btype in ("note", "warning"):
                text = input(f"    {btype.title()} text: ").strip()
                content.append({"type": btype, "text": text})

        tmpl.add_section(sec_title, content)

    # Chain of command
    print("\nChain of command (blank to stop, enter ranks top to bottom):")
    chain = []
    while True:
        rank = input("  Rank: ").strip()
        if not rank:
            break
        chain.append(rank)
    if chain:
        tmpl.set_chain_of_command(chain)

    return tmpl


def main():
    parser = argparse.ArgumentParser(
        prog="clrp-docgen",
        description=(
            "CLRP-DOCGEN: Training Document Generator for "
            "Clone Wars RP — 327th Star Corps / K Company"
        ),
    )
    parser.add_argument(
        "--version", action="version",
        version=f"%(prog)s {__version__}",
    )

    subparsers = parser.add_subparsers(dest="command", help="Commands")

    # --- generate from config ---
    gen_parser = subparsers.add_parser(
        "generate", aliases=["gen"],
        help="Generate a document from a YAML config file",
    )
    gen_parser.add_argument(
        "config", help="Path to YAML config file",
    )
    gen_parser.add_argument(
        "-o", "--output", default=None,
        help="Output file path (default: derived from config)",
    )
    gen_parser.add_argument(
        "-f", "--format", choices=["docx", "pdf"], default=None,
        help="Output format (default: docx)",
    )

    # --- interactive mode ---
    int_parser = subparsers.add_parser(
        "interactive", aliases=["int"],
        help="Build a document interactively with prompts",
    )
    int_parser.add_argument(
        "doc_type", choices=["sop", "tryout", "handbook"],
        help="Type of document to create",
    )
    int_parser.add_argument(
        "-o", "--output", default=None,
        help="Output file path",
    )
    int_parser.add_argument(
        "-f", "--format", choices=["docx", "pdf"], default="docx",
        help="Output format (default: docx)",
    )

    # --- list themes ---
    subparsers.add_parser(
        "themes", help="List available themes and their colors",
    )

    # --- list templates ---
    subparsers.add_parser(
        "templates", help="List available document templates",
    )

    args = parser.parse_args()

    if args.command in ("generate", "gen"):
        if not os.path.exists(args.config):
            print(f"Error: Config file not found: {args.config}")
            sys.exit(1)
        output = generate_from_config(
            args.config,
            output_path=args.output,
            fmt=args.format,
        )
        print(f"Document generated: {output}")

    elif args.command in ("interactive", "int"):
        builders = {
            "sop": _interactive_sop,
            "tryout": _interactive_tryout,
            "handbook": _interactive_handbook,
        }
        tmpl = builders[args.doc_type](args)

        output = args.output
        if not output:
            safe_title = tmpl.title.replace(" ", "_")
            output = os.path.join("output", f"{safe_title}.{args.format}")

        os.makedirs(os.path.dirname(output) or ".", exist_ok=True)
        result = tmpl.save(output, fmt=args.format)
        print(f"\nDocument generated: {result}")

    elif args.command == "themes":
        print("\nAvailable Themes:")
        print("-" * 40)
        for name, theme in THEMES.items():
            print(f"\n  {name}:")
            print(f"    Title color:    {theme.title_color}")
            print(f"    Heading color:  {theme.heading_color}")
            print(f"    Accent color:   {theme.accent_color}")
            if theme.use_section_symbols:
                print(f"    Section symbol: {theme.section_symbol}")

    elif args.command == "templates":
        print("\nAvailable Document Templates:")
        print("-" * 40)
        templates = {
            "sop": "Standard Operating Procedure — structured procedures "
                   "with numbered steps, responsibilities, and revision "
                   "history.",
            "tryout": "Tryout Document — phased tryout guide with "
                      "color-coded instructions, Q&A, and setup "
                      "checklists.",
            "handbook": "Handbook / Guide — informational handbook with "
                        "codes, rules, chain of command, and reference "
                        "links.",
        }
        for name, desc in templates.items():
            print(f"\n  {name}:")
            print(f"    {desc}")

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
