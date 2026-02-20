"""
OpenReport Visitor Module

This module contains the base abstract Visitor class that implements the visitor pattern
for processing OpenReport document nodes.

Classes:
    Visitor: Base abstract class for all visitor implementations
"""

from __future__ import annotations
from abc import abstractmethod, ABC

from OpenReport.base.nodes import (
    AttributeNode,
    ItemsNode,
    NameNode,
)
from OpenReport.enumerations.attribute_nodes import ATTRIBUTE_NODES
from OpenReport.enumerations.open_report_nodes import NODES
from OpenReport.visitor_context.visitor_context import VisitorContext


class Visitor(ABC):
    """
    Base abstract Visitor class implementing the visitor pattern for OpenReport document processing.
    """
    def resume_traverse(self, node, context):
        for child_node in node.children:
            child_node.accept(visitor=self, context=context)

    def resume_traverse_with_indexing(self, node, context):
        for child_id, child_node in enumerate(node.children):
            context.current_index = child_id
            child_node.accept(visitor=self, context=context)

    def resume_traverse_non_attribute_nodes(self, node, context):
        for child_node in node.children:
            if not isinstance(child_node, AttributeNode):
                child_node.accept(visitor=self, context=context)

    def resume_traverse_items_node(self, node, context):
        for child_node in list(node.children):
            if isinstance(child_node, ItemsNode):
                child_node.bullet_list_attributes = node.attributes
                child_node.accept(visitor=self, context=context)

    @staticmethod
    def process_parent_node(node):
        if isinstance(node.parent, ItemsNode):
            node.bullet_list_attributes = node.parent.parent.attributes

    def process_children_nodes(self, node, context):
        for child_node in list(node.children):
            if isinstance(child_node, AttributeNode):
                node.add_attribute(attribute=child_node.attributes)

    @staticmethod
    def prepare_target_style(node):
        node.target_style_name = None
        if node.default_style_type == NODES.DEFAULT_TEXT_STYLE:
            node.target_style_name = "Normal"
        elif node.default_style_type == NODES.DEFAULT_HEADING_STYLE:
            node.target_style_name = f"Heading {node.attributes[ATTRIBUTE_NODES.LEVEL]}"

    @abstractmethod
    def visit_and_process_document(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_document_params(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_document_style(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_text_style(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_structure(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_heading(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_text(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_attribute(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_math_expression(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_items(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_bullet_list(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_table_of_contents(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_page_break(self, node, context):
        pass

    @abstractmethod
    def visit_and_process_name(self, node, context):
        pass
