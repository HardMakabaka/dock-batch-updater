"""Format preservation utilities for DOCX processing.

This module provides utilities to maintain formatting when replacing text
in DOCX documents, ensuring that fonts, colors, sizes, and paragraph styles
are preserved during batch updates.
"""

from typing import Dict, Any, Optional
from docx.shared import RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import _Cell, Table
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


class FormatPreserver:
    """Utility class to preserve document formatting during text replacement.

    This class provides methods to capture and restore formatting properties
    from text runs and table cells, ensuring that formatting is preserved
    during batch replacement operations.
    """

    @staticmethod
    def capture_run_format(run) -> Dict[str, Any]:
        """Capture formatting properties from a text run.

        Args:
            run: A python-docx Run object

        Returns:
            Dictionary containing formatting properties
        """
        format_data = {}

        # Font properties
        if run.font.name:
            format_data['font_name'] = run.font.name

        if run.font.size:
            format_data['font_size'] = run.font.size

        if run.font.bold is not None:
            format_data['bold'] = run.font.bold

        if run.font.italic is not None:
            format_data['italic'] = run.font.italic

        if run.font.underline:
            format_data['underline'] = run.font.underline

        if run.font.color and run.font.color.rgb:
            format_data['color_rgb'] = run.font.color.rgb

        if run.font.highlight_color:
            format_data['highlight_color'] = run.font.highlight_color

        if run.font.strike is not None:
            format_data['strike'] = run.font.strike

        if run.font.subscript is not None:
            format_data['subscript'] = run.font.subscript

        if run.font.superscript is not None:
            format_data['superscript'] = run.font.superscript

        return format_data

    @staticmethod
    def apply_run_format(run, format_data: Dict[str, Any]) -> None:
        """Apply formatting properties to a text run.

        Args:
            run: A python-docx Run object
            format_data: Dictionary of formatting properties
        """
        # Font properties
        if 'font_name' in format_data:
            run.font.name = format_data['font_name']

        if 'font_size' in format_data:
            run.font.size = format_data['font_size']

        if 'bold' in format_data:
            run.font.bold = format_data['bold']

        if 'italic' in format_data:
            run.font.italic = format_data['italic']

        if 'underline' in format_data:
            run.font.underline = format_data['underline']

        if 'color_rgb' in format_data:
            run.font.color.rgb = format_data['color_rgb']

        if 'highlight_color' in format_data:
            run.font.highlight_color = format_data['highlight_color']

        if 'strike' in format_data:
            run.font.strike = format_data['strike']

        if 'subscript' in format_data:
            run.font.subscript = format_data['subscript']

        if 'superscript' in format_data:
            run.font.superscript = format_data['superscript']

    @staticmethod
    def capture_paragraph_format(paragraph: Paragraph) -> Dict[str, Any]:
        """Capture paragraph formatting properties.

        Args:
            paragraph: A python-docx Paragraph object

        Returns:
            Dictionary containing paragraph formatting properties
        """
        format_data = {}

        if paragraph.alignment:
            format_data['alignment'] = paragraph.alignment

        if paragraph.paragraph_format.left_indent:
            format_data['left_indent'] = paragraph.paragraph_format.left_indent

        if paragraph.paragraph_format.right_indent:
            format_data['right_indent'] = paragraph.paragraph_format.right_indent

        if paragraph.paragraph_format.first_line_indent:
            format_data['first_line_indent'] = paragraph.paragraph_format.first_line_indent

        if paragraph.paragraph_format.space_before:
            format_data['space_before'] = paragraph.paragraph_format.space_before

        if paragraph.paragraph_format.space_after:
            format_data['space_after'] = paragraph.paragraph_format.space_after

        if paragraph.paragraph_format.line_spacing:
            format_data['line_spacing'] = paragraph.paragraph_format.line_spacing

        if paragraph.style:
            format_data['style'] = paragraph.style

        return format_data

    @staticmethod
    def apply_paragraph_format(paragraph: Paragraph, format_data: Dict[str, Any]) -> None:
        """Apply paragraph formatting properties.

        Args:
            paragraph: A python-docx Paragraph object
            format_data: Dictionary of paragraph formatting properties
        """
        if 'alignment' in format_data:
            paragraph.alignment = format_data['alignment']

        if 'left_indent' in format_data:
            paragraph.paragraph_format.left_indent = format_data['left_indent']

        if 'right_indent' in format_data:
            paragraph.paragraph_format.right_indent = format_data['right_indent']

        if 'first_line_indent' in format_data:
            paragraph.paragraph_format.first_line_indent = format_data['first_line_indent']

        if 'space_before' in format_data:
            paragraph.paragraph_format.space_before = format_data['space_before']

        if 'space_after' in format_data:
            paragraph.paragraph_format.space_after = format_data['space_after']

        if 'line_spacing' in format_data:
            paragraph.paragraph_format.line_spacing = format_data['line_spacing']

        if 'style' in format_data:
            paragraph.style = format_data['style']

    @staticmethod
    def capture_cell_format(cell: _Cell) -> Dict[str, Any]:
        """Capture table cell formatting properties.

        Args:
            cell: A python-docx _Cell object

        Returns:
            Dictionary containing cell formatting properties
        """
        format_data = {}

        if cell.width:
            format_data['width'] = cell.width

        if cell.vertical_alignment:
            format_data['vertical_alignment'] = cell.vertical_alignment

        if hasattr(cell, 'shading') and cell.shading and cell.shading.background_color:
            format_data['background_color'] = cell.shading.background_color

        return format_data

    @staticmethod
    def apply_cell_format(cell: _Cell, format_data: Dict[str, Any]) -> None:
        """Apply table cell formatting properties.

        Args:
            cell: A python-docx _Cell object
            format_data: Dictionary of cell formatting properties
        """
        if 'width' in format_data:
            cell.width = format_data['width']

        if 'vertical_alignment' in format_data:
            cell.vertical_alignment = format_data['vertical_alignment']

        if 'background_color' in format_data:
            cell.shading.background_color = format_data['background_color']

    @staticmethod
    def find_text_in_paragraph(paragraph: Paragraph, search_text: str) -> Optional[tuple]:
        """Find text within a paragraph's runs and return run index and position.

        Args:
            paragraph: A python-docx Paragraph object
            search_text: Text to search for

        Returns:
            Tuple of (run_index, start_pos, end_pos) or None if not found
        """
        paragraph_text = paragraph.text

        if search_text not in paragraph_text:
            return None

        # Find position in paragraph text
        start_pos = paragraph_text.find(search_text)
        end_pos = start_pos + len(search_text)

        # Map paragraph position to run and character position within run
        current_pos = 0
        for run_idx, run in enumerate(paragraph.runs):
            run_text = run.text
            run_end = current_pos + len(run_text)

            # Check if search text overlaps with this run
            if start_pos >= current_pos and end_pos <= run_end:
                # Search text is entirely within this run
                return (run_idx, start_pos - current_pos, end_pos - current_pos)
            elif start_pos < run_end and end_pos > run_end:
                # Search text spans multiple runs - more complex case
                # For simplicity, return the starting run and position
                return (run_idx, start_pos - current_pos, len(run_text) - (start_pos - current_pos))

            current_pos = run_end

        return None

    @staticmethod
    def split_run_text(run: Run, split_pos: int) -> tuple:
        """Split a run into two parts at the specified position.

        Args:
            run: A python-docx Run object
            split_pos: Position to split at (0-based)

        Returns:
            Tuple of (before_text, after_text)
        """
        text = run.text
        before_text = text[:split_pos]
        after_text = text[split_pos:]
        return before_text, after_text
