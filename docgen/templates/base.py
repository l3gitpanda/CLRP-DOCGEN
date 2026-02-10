"""
Base template class for document generators.

All document templates inherit from BaseTemplate, which provides
common setup, metadata, and save functionality.
"""

from datetime import datetime
from typing import Optional

from ..styles import StyleConfig, THEMES
from ..engine import DocumentEngine


class BaseTemplate:
    """Base class for all document templates."""

    def __init__(self, style: Optional[StyleConfig] = None,
                 theme: str = "327th"):
        if style:
            self.style = style
        elif theme in THEMES:
            self.style = StyleConfig(**THEMES[theme].__dict__.copy())
        else:
            self.style = StyleConfig()

        self.engine = DocumentEngine(self.style)

        # Default metadata
        self.title = "Untitled Document"
        self.subtitle = ""
        self.author = ""
        self.formatted_by = ""
        self.version_date = datetime.now().strftime("%m/%d/%Y")
        self.unit = "327th Star Corps"
        self.company = "K Company"

    def set_metadata(self, title: str = None, subtitle: str = None,
                     author: str = None, formatted_by: str = None,
                     version_date: str = None, unit: str = None,
                     company: str = None):
        """Set document metadata."""
        if title is not None:
            self.title = title
        if subtitle is not None:
            self.subtitle = subtitle
        if author is not None:
            self.author = author
        if formatted_by is not None:
            self.formatted_by = formatted_by
        if version_date is not None:
            self.version_date = version_date
        if unit is not None:
            self.unit = unit
        if company is not None:
            self.company = company

    def build(self):
        """Build the document. Override in subclasses."""
        raise NotImplementedError("Subclasses must implement build()")

    def save(self, filepath: str, fmt: str = "docx") -> str:
        """Build and save the document."""
        self.build()
        return self.engine.save(filepath, fmt=fmt)
