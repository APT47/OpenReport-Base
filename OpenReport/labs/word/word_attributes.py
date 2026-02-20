"""
OpenReport Word Attributes Module

This module contains attribute classes that handle formatting and styling for Word document
elements. Each attribute class provides a static apply method that applies specific formatting
to Word document objects like runs, fonts, paragraphs, and sections.
"""

from __future__ import annotations
from abc import ABC
from OpenReport.labs.word.word_utilities import *

from docx.enum.section import WD_ORIENT
from docx.shared import Pt, Cm


from OpenReport.enumerations.attribute_nodes import ATTRIBUTE_NODES
from OpenReport.enumerations.text_modes import TEXT_MODES


class SkipAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        pass


# TEXT ATTRIBUTES
class BoldAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'object_runs'):
                raise Exception("OpenReport Error: Target object missing text runs. Cannot apply bold formatting.")
            if not isinstance(value, bool):
                raise Exception(f"OpenReport Error: Bold formatting value must be true or false, got {type(value).__name__}: {value}")
            if not target_object.object_runs:
                return
            for run in target_object.object_runs:
                run.bold = value
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply bold formatting. {str(e)}") from e


class ItalicAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'object_runs'):
                raise Exception("OpenReport Error: Target object missing text runs. Cannot apply italic formatting.")
            if not isinstance(value, bool):
                raise Exception(f"OpenReport Error: Italic formatting value must be true or false, got {type(value).__name__}: {value}")
            if not target_object.object_runs:
                return
            for run in target_object.object_runs:
                run.italic = value
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply italic formatting. {str(e)}") from e


class SizeAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'object_fonts'):
                raise Exception("OpenReport Error: Target object missing text fonts. Cannot apply font size formatting.")
            if not isinstance(value, (int, float)):
                raise Exception(f"OpenReport Error: Font size must be a number, got {type(value).__name__}: {value}")
            if value <= 0:
                raise Exception(f"OpenReport Error: Font size must be positive, got {value}")
            if value > 400:
                raise Exception(f"OpenReport Error: Font size {value} is too large. Maximum recommended size is 400 points.")
            if not target_object.object_fonts:
                return
            for run_font in target_object.object_fonts:
                run_font.size = Pt(value)
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply font size formatting. {str(e)}") from e


class ColourAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'object_fonts'):
                raise Exception("OpenReport Error: Target object missing text fonts. Cannot apply color formatting.")
            if not isinstance(value, str):
                raise Exception(f"OpenReport Error: Color value must be a string (color name or hex code), got {type(value).__name__}: {value}")
            if not value.strip():
                raise Exception("OpenReport Error: Color value cannot be empty. Please provide a color name or hex code.")
            if not target_object.object_fonts:
                return
            try:
                color_rgb = recognise_colour(value)
            except Exception as e:
                raise Exception(f"OpenReport Error: Invalid color '{value}'. {str(e)}") from e
            for run_font in target_object.object_fonts:
                run_font.color.rgb = color_rgb
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply color formatting. {str(e)}") from e


class FontAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'object_fonts'):
                raise Exception("OpenReport Error: Target object missing text fonts. Cannot apply font family formatting.")
            if not isinstance(value, str):
                raise Exception(f"OpenReport Error: Font family name must be a string, got {type(value).__name__}: {value}")
            if not value.strip():
                raise Exception("OpenReport Error: Font family name cannot be empty.")
            if not target_object.object_fonts:
                return
            for run_font in target_object.object_fonts:
                run_font.name = value
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply font family formatting. {str(e)}") from e


class AlignmentAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'this_object'):
                raise Exception("OpenReport Error: Target object missing paragraph reference. Cannot apply alignment formatting.")
            if not isinstance(value, str):
                raise Exception(f"OpenReport Error: Alignment value must be a string, got {type(value).__name__}: {value}")
            if not value.strip():
                raise Exception("OpenReport Error: Alignment value cannot be empty.")
            alignment_upper = value.upper()
            if alignment_upper not in wd_align_paragraph_mapping:
                valid_alignments = list(wd_align_paragraph_mapping.keys())
                raise Exception(f"OpenReport Error: Invalid alignment '{value}'. Valid alignments are: {valid_alignments}")
            target_object.this_object.alignment = wd_align_paragraph_mapping[alignment_upper]
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply alignment formatting. {str(e)}") from e


class UnderlineAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'object_runs'):
                raise Exception("OpenReport Error: Target object missing text runs. Cannot apply underline formatting.")
            if not target_object.object_runs:
                return
            if isinstance(value, bool):
                for run in target_object.object_runs:
                    run.underline = value
            elif isinstance(value, str):
                if not value.strip():
                    raise Exception("OpenReport Error: Underline type cannot be empty.")
                underline_upper = value.upper()
                if underline_upper not in wd_underline_mapping:
                    valid_types = list(wd_underline_mapping.keys())
                    raise Exception(f"OpenReport Error: Invalid underline type '{value}'. Valid types are: {valid_types}")
                for run in target_object.object_runs:
                    run.underline = wd_underline_mapping[underline_upper]
            else:
                raise Exception(f"OpenReport Error: Underline value must be true/false or an underline type string, got {type(value).__name__}: {value}")
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply underline formatting. {str(e)}") from e


class HighlightColourAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'object_fonts'):
                raise Exception("OpenReport Error: Target object missing text fonts. Cannot apply highlight color formatting.")
            if not isinstance(value, str):
                raise Exception(f"OpenReport Error: Highlight color must be a string, got {type(value).__name__}: {value}")
            if not value.strip():
                raise Exception("OpenReport Error: Highlight color cannot be empty.")
            if not target_object.object_fonts:
                return
            highlight_upper = value.upper()
            if highlight_upper not in wd_color_index_mapping:
                valid_colors = list(wd_color_index_mapping.keys())
                raise Exception(f"OpenReport Error: Invalid highlight color '{value}'. Valid highlight colors are: {valid_colors}")
            for run_font in target_object.object_fonts:
                run_font.highlight_color = wd_color_index_mapping[highlight_upper]
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply highlight color formatting. {str(e)}") from e


class PageBreakBeforeAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'paragraph_format'):
                raise Exception("OpenReport Error: Target object missing paragraph format.")
            if not isinstance(value, bool):
                raise Exception(f"OpenReport Error: Page break before value must be true or false, got {type(value).__name__}: {value}")
            target_object.paragraph_format.page_break_before = value
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply page break before formatting. {str(e)}") from e


class KeepWithNextAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'paragraph_format'):
                raise Exception("OpenReport Error: Target object missing paragraph format.")
            if not isinstance(value, bool):
                raise Exception(f"OpenReport Error: Keep with next value must be true or false, got {type(value).__name__}: {value}")
            target_object.paragraph_format.keep_with_next = value
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply keep with next formatting. {str(e)}") from e


class KeepTogetherAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'paragraph_format'):
                raise Exception("OpenReport Error: Target object missing paragraph format.")
            if not isinstance(value, bool):
                raise Exception(f"OpenReport Error: Keep together value must be true or false, got {type(value).__name__}: {value}")
            target_object.paragraph_format.keep_together = value
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply keep together formatting. {str(e)}") from e


class FirstLineIndentAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'paragraph_format'):
                raise Exception("OpenReport Error: Target object missing paragraph format.")
            if not isinstance(value, (int, float)):
                raise Exception(f"OpenReport Error: First line indent must be a number, got {type(value).__name__}: {value}")
            target_object.paragraph_format.first_line_indent = Cm(value)
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply first line indent formatting. {str(e)}") from e


class SpaceBeforeAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'paragraph_format'):
                raise Exception("OpenReport Error: Target object missing paragraph format.")
            if not isinstance(value, (int, float)):
                raise Exception(f"OpenReport Error: Space before must be a number, got {type(value).__name__}: {value}")
            target_object.paragraph_format.space_before = Cm(value)
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply space before formatting. {str(e)}") from e


class SpaceAfterAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        try:
            if not hasattr(target_object, 'paragraph_format'):
                raise Exception("OpenReport Error: Target object missing paragraph format.")
            if not isinstance(value, (int, float)):
                raise Exception(f"OpenReport Error: Space after must be a number, got {type(value).__name__}: {value}")
            target_object.paragraph_format.space_after = Cm(value)
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to apply space after formatting. {str(e)}") from e


class LineSpacingAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            target_object.paragraph_format.line_spacing = Pt(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


class LineSpacingRuleAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, str):
            target_object.paragraph_format.line_spacing_rule = wd_line_spacing_mapping[
                value.upper()
            ]
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be string."
            )


class LeftIndentAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            target_object.paragraph_format.left_indent = Cm(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


class RightIndentAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            target_object.paragraph_format.right_indent = Cm(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


# TEXT RUN ATTRIBUTES
class NormalRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        pass


class BoldRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.bold = True


class ItalicRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.italic = True


class UnderlineRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.underline = True


class NoProofRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.font.no_proof = True


class OutlineRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.font.outline = True


class ShadowRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.font.shadow = True


class SubscriptRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.font.subscript = True


class SuperscriptRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.font.superscript = True


class BoldItalicRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.bold = True
        target_object.italic = True


class BoldUnderlineRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.underline = True
        target_object.bold = True


class ItalicUnderlineRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.underline = True
        target_object.italic = True


class BoldItalicUnderlineRunAttribute(ABC):
    @staticmethod
    def apply(target_object):
        target_object.underline = True
        target_object.bold = True
        target_object.italic = True


# DOCUMENT STYLE ATTRIBUTES
class StyleAlignmentAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, str):
            target_object.font.alignment = wd_align_paragraph_mapping[value.upper()]
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be string."
            )


class StyleBoldAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, bool):
            target_object.font.bold = value
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be bool."
            )


class StyleColourAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, str):
            target_object.font.color.rgb = recognise_colour(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be string."
            )


class StyleFontAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, str):
            target_object.font.name = value
            rFonts = target_object.element.rPr.rFonts
            rFonts.set(qn("w:asciiTheme"), value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be string."
            )


class StyleHighlightColourAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, str):
            target_object.font.highlight_color = wd_color_index_mapping[value.upper()]
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be string."
            )


class StyleItalicAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, bool):
            target_object.font.italic = value
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be bool."
            )


class StyleSizeAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            target_object.font.size = Pt(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


class StyleUnderlineAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, bool):
            target_object.font.underline = value
        elif isinstance(value, str):
            target_object.font.underline = wd_underline_mapping[value.upper()]
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be bool or string."
            )


# DOCUMENT PARAMETERS ATTRIBUTES
class RightMarginAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            for section in target_object.sections:
                section.right_margin = Cm(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


class LeftMarginAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            for section in target_object.sections:
                section.left_margin = Cm(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


class BottomMarginAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            for section in target_object.sections:
                section.bottom_margin = Cm(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


class TopMarginAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, int) or isinstance(value, float):
            for section in target_object.sections:
                section.top_margin = Cm(value)
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be int or float."
            )


class LandscapeAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, bool):
            if value:
                for section in target_object.sections:
                    section.orientation = WD_ORIENT.LANDSCAPE
                    new_width, new_height = section.page_height, section.page_width
                    section.page_width = new_width
                    section.page_height = new_height
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be bool."
            )


class PageNumberAlignmentAttribute(ABC):
    @staticmethod
    def apply(target_object, value):
        if isinstance(value, str):
            target_object.sections[0].footer.paragraphs[0].alignment = (
                wd_align_page_mapping[value.upper()]
            )
        else:
            raise TypeError(
                f"Bad value type {type(value)} for value {value}, should be string."
            )


attribute_to_class_map = {
    ATTRIBUTE_NODES.ALIGNMENT: AlignmentAttribute,
    ATTRIBUTE_NODES.BODY: SkipAttribute,
    ATTRIBUTE_NODES.BOLD: BoldAttribute,
    ATTRIBUTE_NODES.BOTTOM_MARGIN: BottomMarginAttribute,
    ATTRIBUTE_NODES.COLOUR: ColourAttribute,
    ATTRIBUTE_NODES.FIRST_LINE_INDENT: FirstLineIndentAttribute,
    ATTRIBUTE_NODES.FONT: FontAttribute,
    ATTRIBUTE_NODES.HIGHLIGHT_COLOUR: HighlightColourAttribute,
    ATTRIBUTE_NODES.ITALIC: ItalicAttribute,
    ATTRIBUTE_NODES.KEEP_TOGETHER: KeepTogetherAttribute,
    ATTRIBUTE_NODES.KEEP_WITH_NEXT: KeepWithNextAttribute,
    ATTRIBUTE_NODES.LANDSCAPE: LandscapeAttribute,
    ATTRIBUTE_NODES.LEFT_INDENT: LeftIndentAttribute,
    ATTRIBUTE_NODES.LEFT_MARGIN: LeftMarginAttribute,
    ATTRIBUTE_NODES.LEVEL: SkipAttribute,
    ATTRIBUTE_NODES.LINE_SPACING_RULE: LineSpacingRuleAttribute,
    ATTRIBUTE_NODES.LINE_SPACING: LineSpacingAttribute,
    ATTRIBUTE_NODES.PAGE_BREAK_BEFORE: PageBreakBeforeAttribute,
    ATTRIBUTE_NODES.PAGE_NUMBERING: SkipAttribute,
    ATTRIBUTE_NODES.PAGE_NUMBER_ALIGNMENT: PageNumberAlignmentAttribute,
    ATTRIBUTE_NODES.RIGHT_INDENT: RightIndentAttribute,
    ATTRIBUTE_NODES.RIGHT_MARGIN: RightMarginAttribute,
    ATTRIBUTE_NODES.SIZE: SizeAttribute,
    ATTRIBUTE_NODES.SKIP_COVER_PAGE: SkipAttribute,
    ATTRIBUTE_NODES.SPACE_AFTER: SpaceAfterAttribute,
    ATTRIBUTE_NODES.SPACE_BEFORE: SpaceBeforeAttribute,
    ATTRIBUTE_NODES.TOP_MARGIN: TopMarginAttribute,
    ATTRIBUTE_NODES.UNDERLINE: UnderlineAttribute,
}

run_attribute_to_class_map = {
    TEXT_MODES.NORMAL: NormalRunAttribute,
    TEXT_MODES.TEXTBF: BoldRunAttribute,
    TEXT_MODES.TEXTIT: ItalicRunAttribute,
    TEXT_MODES.TEXTUN: UnderlineRunAttribute,
    TEXT_MODES.TEXTBFIT: BoldItalicRunAttribute,
    TEXT_MODES.TEXTBFUN: BoldUnderlineRunAttribute,
    TEXT_MODES.TEXTITUN: ItalicUnderlineRunAttribute,
    TEXT_MODES.TEXTBFITUN: BoldItalicUnderlineRunAttribute,
    TEXT_MODES.TEXTSUPERSCRIPT: SuperscriptRunAttribute,
    TEXT_MODES.TEXTSUBSCRIPT: SubscriptRunAttribute,
    TEXT_MODES.TEXTSHADOW: ShadowRunAttribute,
    TEXT_MODES.TEXTOUTLINE: OutlineRunAttribute,
    TEXT_MODES.TEXTNOPROOF: NoProofRunAttribute,
}

math_expression_attribute_to_class_map = {
    ATTRIBUTE_NODES.ALIGNMENT: StyleAlignmentAttribute,
    ATTRIBUTE_NODES.BODY: SkipAttribute,
    ATTRIBUTE_NODES.BOLD: StyleBoldAttribute,
    ATTRIBUTE_NODES.COLOUR: StyleColourAttribute,
    ATTRIBUTE_NODES.FIRST_LINE_INDENT: FirstLineIndentAttribute,
    ATTRIBUTE_NODES.FONT: StyleFontAttribute,
    ATTRIBUTE_NODES.HIGHLIGHT_COLOUR: StyleHighlightColourAttribute,
    ATTRIBUTE_NODES.ITALIC: StyleItalicAttribute,
    ATTRIBUTE_NODES.KEEP_TOGETHER: KeepTogetherAttribute,
    ATTRIBUTE_NODES.KEEP_WITH_NEXT: KeepWithNextAttribute,
    ATTRIBUTE_NODES.LEFT_INDENT: LeftIndentAttribute,
    ATTRIBUTE_NODES.LEFT_MARGIN: LeftMarginAttribute,
    ATTRIBUTE_NODES.LINE_SPACING_RULE: LineSpacingRuleAttribute,
    ATTRIBUTE_NODES.LINE_SPACING: LineSpacingAttribute,
    ATTRIBUTE_NODES.PAGE_BREAK_BEFORE: PageBreakBeforeAttribute,
    ATTRIBUTE_NODES.RIGHT_INDENT: RightIndentAttribute,
    ATTRIBUTE_NODES.SIZE: StyleSizeAttribute,
    ATTRIBUTE_NODES.SPACE_AFTER: SpaceAfterAttribute,
    ATTRIBUTE_NODES.SPACE_BEFORE: SpaceBeforeAttribute,
    ATTRIBUTE_NODES.UNDERLINE: StyleUnderlineAttribute,
}


document_text_font_attribute_to_class_map = {
    ATTRIBUTE_NODES.ALIGNMENT: StyleAlignmentAttribute,
    ATTRIBUTE_NODES.BOLD: StyleBoldAttribute,
    ATTRIBUTE_NODES.COLOUR: StyleColourAttribute,
    ATTRIBUTE_NODES.FONT: StyleFontAttribute,
    ATTRIBUTE_NODES.HIGHLIGHT_COLOUR: StyleHighlightColourAttribute,
    ATTRIBUTE_NODES.ITALIC: StyleItalicAttribute,
    ATTRIBUTE_NODES.LEVEL: SkipAttribute,
    ATTRIBUTE_NODES.SIZE: StyleSizeAttribute,
    ATTRIBUTE_NODES.UNDERLINE: StyleUnderlineAttribute,
}

document_text_paragraph_format_attribute_to_class_map = {
    ATTRIBUTE_NODES.FIRST_LINE_INDENT: FirstLineIndentAttribute,
    ATTRIBUTE_NODES.KEEP_TOGETHER: KeepTogetherAttribute,
    ATTRIBUTE_NODES.KEEP_WITH_NEXT: KeepWithNextAttribute,
    ATTRIBUTE_NODES.LEFT_INDENT: LeftIndentAttribute,
    ATTRIBUTE_NODES.LEFT_MARGIN: LeftMarginAttribute,
    ATTRIBUTE_NODES.LINE_SPACING_RULE: LineSpacingRuleAttribute,
    ATTRIBUTE_NODES.LINE_SPACING: LineSpacingAttribute,
    ATTRIBUTE_NODES.PAGE_BREAK_BEFORE: PageBreakBeforeAttribute,
    ATTRIBUTE_NODES.RIGHT_INDENT: RightIndentAttribute,
    ATTRIBUTE_NODES.SPACE_AFTER: SpaceAfterAttribute,
    ATTRIBUTE_NODES.SPACE_BEFORE: SpaceBeforeAttribute,
}
