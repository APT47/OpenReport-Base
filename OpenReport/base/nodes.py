"""
OpenReport Node Module

This module contains the abstract syntax tree node classes that represent
different document components in the OpenReport system. All nodes inherit
from the base OpenReportNode class and implement the visitor pattern.
"""

from __future__ import annotations
from abc import ABC


class OpenReportNode(ABC):
    """
    Base abstract class for all OpenReport document nodes.
    """

    def __init__(self) -> None:
        self.parent = None
        self.name = None
        self.children: list = []
        self.attributes: dict = {}
        self.attributes_sorted = None
        self.signature_id: str = ""

    def add_node(self, node) -> None:
        try:
            if node is None:
                raise Exception("OpenReport Error: Cannot add None as a child node.")

            if not isinstance(node, OpenReportNode):
                raise Exception(f"OpenReport Error: Can only add OpenReportNode instances as children, got {type(node).__name__}.")

            self.children.append(node)
            node.parent = self

        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to add child node. {str(e)}") from e

    def insert_node(self, node_position, node) -> None:
        try:
            if node is None:
                raise Exception("OpenReport Error: Cannot insert None as a child node.")

            if not isinstance(node, OpenReportNode):
                raise Exception(f"OpenReport Error: Can only insert OpenReportNode instances as children, got {type(node).__name__}.")

            if not isinstance(node_position, int):
                raise Exception(f"OpenReport Error: Node position must be an integer, got {type(node_position).__name__}.")

            if node_position < 0 or node_position > len(self.children):
                raise Exception(f"OpenReport Error: Invalid node position {node_position}. Must be between 0 and {len(self.children)}.")

            self.children.insert(node_position, node)

        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to insert child node at position {node_position}. {str(e)}") from e

    def remove_node(self, node) -> None:
        try:
            if node is None:
                raise Exception("OpenReport Error: Cannot remove None node.")

            if node not in self.children:
                raise Exception("OpenReport Error: Node is not a child of this node and cannot be removed.")

            self.children.remove(node)
            node.parent = None

        except ValueError as e:
            raise Exception("OpenReport Error: Node is not found in children list and cannot be removed.") from e
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to remove child node. {str(e)}") from e

    def add_attribute(self, attribute: dict) -> None:
        try:
            if attribute is None:
                return

            if not isinstance(attribute, dict):
                raise Exception(f"OpenReport Error: Attributes must be a dictionary, got {type(attribute).__name__}.")

            self.attributes.update(attribute)

        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to add attributes to node. {str(e)}") from e

    def add_attributes_from_attribute_nodes(self):
        for child_node in self.children:
            self.add_attribute(
                child_node.attributes if isinstance(child_node, AttributeNode) else {}
            )


class DocumentNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_document(self, context)


class DocumentParamsNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_document_params(self, context)


class DocumentStyleNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_document_style(self, context)


class ParagraphStyleNode(OpenReportNode):
    def __init__(self, default_style_type: str) -> None:
        super().__init__()
        self.default_style_type = default_style_type

    def accept(self, visitor, context):
        visitor.visit_and_process_text_style(self, context)


class StructureNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_structure(self, context)


class TextNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_text(self, context)


class HeadingNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_heading(self, context)


class AttributeNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_attribute(self, context)


class MathExpressionNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_math_expression(self, context)


class BulletListNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_bullet_list(self, context)


class ItemsNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_items(self, context)


class TableOfContentsNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_table_of_contents(self, context)


class PageBreakNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_page_break(self, context)


class NameNode(OpenReportNode):
    def accept(self, visitor, context):
        visitor.visit_and_process_name(self, context)
