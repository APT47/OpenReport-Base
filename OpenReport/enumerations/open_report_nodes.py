import enum


class NODES(enum.StrEnum):
    """Specifies OpenReport non-attribute nodes."""

    # REQUIRED
    DOCUMENT = "document"
    DOCUMENT_PARAMS = "document_params"
    STRUCTURE = "structure"

    # OPTIONAL
    NAME = "name"
    HEADING = "heading"
    TEXT = "text"
    PAGE_BREAK = "page_break"
    MATH_EXPRESSION = "math_expression"
    TABLE_OF_CONTENTS = "table_of_contents"
    BULLET_LIST = "bullet_list"
    ITEMS = "items"
    DOCUMENT_STYLE = "document_style"
    DEFAULT_TEXT_STYLE = "default_text_style"
    DEFAULT_HEADING_STYLE = "default_heading_style"
