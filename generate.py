#!/usr/bin/env python3
"""
Convenience entry point for CLRP-DOCGEN.

Usage:
    python generate.py generate configs/kc_tryout.yaml
    python generate.py generate configs/kc_tryout.yaml -f pdf
    python generate.py generate configs/kc_tryout.yaml -o my_doc.docx
    python generate.py interactive tryout
    python generate.py themes
    python generate.py templates
"""

from docgen.cli import main

if __name__ == "__main__":
    main()
