"""
OpenReport Word Visitor Module

This module contains the WordVisitor class that implements document generation
for Microsoft Word format.

Classes:
    WordVisitor: Concrete visitor implementation for Word document generation
"""

from __future__ import annotations
import os
from pathlib import Path

from OpenReport.labs.word.word_objects import *
from OpenReport.visitor_context.visitor_context import WordVisitorContext
from OpenReport.visitors.visitor import Visitor


class WordVisitor(Visitor):
    """
    Concrete visitor implementation for generating Microsoft Word documents.
    """
    def __init__(self):
        self.word_object = None
        self.document = Document()

    def generate_content(self, yaml_tree):
        try:
            context = WordVisitorContext()
            yaml_tree.accept(visitor=self, context=context)
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to generate Word document content. {str(e)}") from e

    def save_content_docx(self, save_location: str, yaml_tree: object):
        try:
            if not save_location:
                raise Exception("OpenReport Error: Save location not specified.")

            if not hasattr(yaml_tree, 'name') or not yaml_tree.name:
                raise Exception("OpenReport Error: Document name not specified in YAML structure.")

            save_path = Path(save_location)
            try:
                save_path.mkdir(parents=True, exist_ok=True)
            except PermissionError:
                raise Exception(f"OpenReport Error: Permission denied when trying to create save directory '{save_location}'. Please check permissions.")
            except OSError as e:
                raise Exception(f"OpenReport Error: Cannot create save directory '{save_location}'. {str(e)}")

            full_name = save_path / yaml_tree.name

            if not str(full_name).lower().endswith('.docx'):
                full_name = full_name.with_suffix('.docx')

            try:
                self.document.save(str(full_name))
                print(f"File saved at: {full_name}")
            except PermissionError:
                raise Exception(f"OpenReport Error: Permission denied when trying to save document to '{full_name}'. Please check file permissions or ensure the file is not open.")
            except OSError as e:
                raise Exception(f"OpenReport Error: Cannot save document to '{full_name}'. {str(e)}")

            try:
                if os.name == 'nt' and not os.getenv('OPENREPORT_HEADLESS'):
                    os.startfile(str(full_name))
            except Exception:
                pass

        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to save Word document. {str(e)}") from e

    def save_content_pdf(self, file_name: str):
        try:
            if not file_name:
                raise Exception("OpenReport Error: No file name provided for PDF conversion.")

            docx_path = Path(file_name)

            if not docx_path.exists():
                raise Exception(f"OpenReport Error: DOCX file not found at '{file_name}' for PDF conversion.")

            if not str(docx_path).lower().endswith('.docx'):
                raise Exception(f"OpenReport Error: Expected a DOCX file for PDF conversion, but got '{docx_path.suffix}'.")

            pdf_file = str(docx_path.with_suffix('.pdf'))

            try:
                from docx2pdf import convert
                convert(str(docx_path), pdf_file)
                print(f"PDF saved at: {pdf_file}")
            except ImportError as e:
                raise Exception("OpenReport Error: PDF conversion requires additional libraries. Please install docx2pdf: pip install docx2pdf") from e
            except Exception as e:
                raise Exception(f"OpenReport Error: Failed to convert DOCX to PDF. This may be due to missing Microsoft Word installation or other system requirements. {str(e)}") from e

        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to convert document to PDF. {str(e)}") from e

    def visit_and_process_document(self, node, context):
        self.resume_traverse(node=node, context=context)

    def visit_and_process_document_params(self, node, context):
        node.add_attributes_from_attribute_nodes()
        self.word_object = DocumentParams(attributes=node.attributes)
        self.word_object.add_to_document(word_document=self)
        self.resume_traverse(node=node, context=context)

    def visit_and_process_document_style(self, node, context):
        self.resume_traverse_non_attribute_nodes(node=node, context=context)

    def visit_and_process_text_style(self, node, context):
        node.add_attributes_from_attribute_nodes()
        self.prepare_target_style(node=node)
        self.word_object = ParagraphTextStyle(
            attributes=node.attributes, target_style_name=node.target_style_name
        )
        self.word_object.add_to_document(word_document=self)
        self.resume_traverse(node=node, context=context)

    def visit_and_process_structure(self, node, context):
        self.resume_traverse(node=node, context=context)

    def visit_and_process_heading(self, node, context):
        node.cell_attributes = {}
        node.bullet_list_attributes = {}

        self.process_children_nodes(node=node, context=context)
        self.word_object = Heading(
            attributes=node.attributes, cell_attributes=node.cell_attributes
        )
        self.word_object.add_to_document(word_document=self)
        self.resume_traverse(node=node, context=context)

    def visit_and_process_text(self, node, context):
        node.cell_attributes = {}
        node.bullet_list_attributes = {}
        self.process_parent_node(node=node)
        self.process_children_nodes(node=node, context=context)
        self.word_object = Text(
            attributes=node.attributes,
            cell_attributes=node.cell_attributes,
            bullet_list_attributes=node.bullet_list_attributes,
        )
        self.word_object.add_to_document(word_document=self)
        self.resume_traverse(node=node, context=context)

    def visit_and_process_math_expression(self, node, context):
        node.cell_attributes = {}
        node.bullet_list_attributes = {}
        self.process_parent_node(node=node)
        self.process_children_nodes(node=node, context=context)
        self.word_object = MathExpression(
            attributes=node.attributes,
            cell_attributes=node.cell_attributes,
            bullet_list_attributes=node.bullet_list_attributes,
        )
        self.word_object.add_to_document(word_document=self)
        self.resume_traverse(node=node, context=context)

    def visit_and_process_bullet_list(self, node, context):
        node.cell_attributes = {}
        node.add_attributes_from_attribute_nodes()
        self.resume_traverse_items_node(node=node, context=context)

    def visit_and_process_items(self, node, context):
        node.bullet_list_attributes = {}
        self.process_parent_node(node=node)
        self.resume_traverse_non_attribute_nodes(node=node, context=context)

    def visit_and_process_table_of_contents(self, node, context):
        self.word_object = TableOfContents()
        self.word_object.add_to_document(word_document=self)
        self.resume_traverse(node=node, context=context)

    def visit_and_process_page_break(self, node, context):
        self.process_children_nodes(node=node, context=context)
        self.word_object = PageBreak(attributes=node.attributes)
        self.word_object.add_to_document(word_document=self)
        self.resume_traverse(node=node, context=context)

    def visit_and_process_name(self, node, context):
        node.parent.name = node.attributes[ATTRIBUTE_NODES.NAME]
        self.resume_traverse(node=node, context=context)

    def visit_and_process_attribute(self, node, context):
        pass
