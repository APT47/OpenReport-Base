"""
OpenReport Attribute Nodes Enumeration Module

This module defines all attribute node constants used throughout the OpenReport system.
These constants represent the various attributes that can be applied to document elements
such as formatting, positioning, styling, and configuration options.

Classes:
    ATTRIBUTE_NODES: Enumeration of all attribute node constants used in OpenReport
"""

import enum


class ATTRIBUTE_NODES(enum.StrEnum):
    """
    Enumeration of all attribute node constants used in OpenReport.
    """

    # DOCUMENT ATTRIBUTES
    NAME = "name"
    LANDSCAPE = "landscape"
    MARGINS = "margins"
    TOP_MARGIN = "top_margin"
    BOTTOM_MARGIN = "bottom_margin"
    LEFT_MARGIN = "left_margin"
    RIGHT_MARGIN = "right_margin"
    PAGE_NUMBERING = "page_numbering"
    PAGE_NUMBER_ALIGNMENT = "page_number_alignment"
    SKIP_COVER_PAGE = "skip_cover_page"

    # TEXT/HEADING ATTRIBUTES
    BODY = "body"
    COLOUR = "colour"
    LEVEL = "level"
    SIZE = "size"
    BOLD = "bold"
    ITALIC = "italic"
    UNDERLINE = "underline"
    HIGHLIGHT_COLOUR = "highlight_color"
    FONT = "font"
    ALIGNMENT = "alignment"
    LEFT_INDENT = "left_indent"
    RIGHT_INDENT = "right_indent"
    FIRST_LINE_INDENT = "first_line_indent"
    SPACE_BEFORE = "space_after"
    SPACE_AFTER = "space_before"
    LINE_SPACING = "line_spacing"
    LINE_SPACING_RULE = "line_spacing_rule"
    PAGE_BREAK_BEFORE = "page_break_before"
    KEEP_WITH_NEXT = "keep_with_next"
    KEEP_TOGETHER = "keep_together"

    # PAGE BREAK ATTRIBUTES
    NUMBER_OF_PAGES = "number_of_pages"

    # BULLET LIST ATTRIBUTES
    BULLET_LIST_STYLE = "bullet_list_style"
