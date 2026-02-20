"""
OpenReport Text Modes Enumeration Module

This module defines text formatting mode constants used for text styling and formatting
in OpenReport documents. These modes support various text effects including mathematical
notation, formatting styles, and special text effects.

Classes:
    TEXT_MODES: Enumeration of text formatting mode constants
"""

import enum


class TEXT_MODES(enum.StrEnum):
    """
    Enumeration of text formatting mode constants used in OpenReport.
    
    This class defines constants for various text formatting modes including
    mathematical notation, bold/italic/underline styles, and special effects.
    Some modes are specifically designed for LaTeX output format compatibility.
    
    The use of StrEnum allows these constants to be used as both strings and
    enumeration values for consistent text formatting across different output formats.
    """

    MATH_MODE = "math_mode"
    NORMAL = "normal"
    TEXTBF = "textbf"  # Latex
    TEXTIT = "textit"  # Latex
    TEXTUN = "textun"  # Latex
    TEXTBFIT = "textbfit"
    TEXTBFUN = "textbfun"
    TEXTITUN = "textbfitun"
    TEXTBFITUN = "textbfitun"
    TEXTSUPERSCRIPT = "textsuperscript"  # Latex
    TEXTSUBSCRIPT = "textsubscript"  # Latex
    TEXTSHADOW = "textshadow"
    TEXTOUTLINE = "textoutline"
    TEXTNOPROOF = "textnoproof"
