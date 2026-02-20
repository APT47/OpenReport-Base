"""
OpenReport Document Generator Module

This module contains the main orchestrator class for processing YAML specifications
and generating documents in various formats.

Classes:
    OpenReportDocumentGenerator: Main class for processing YAML input and generating documents
"""

from __future__ import annotations
import yaml
import os
from pathlib import Path
from OpenReport.base.nodes import *
from OpenReport.visitors.visitor_word import WordVisitor
from OpenReport.enumerations.open_report_nodes import NODES
from OpenReport.enumerations.attribute_nodes import ATTRIBUTE_NODES


class OpenReportDocumentGenerator:
    """
    A class to process .yaml into desired output format.
    :param yaml_input: .yaml file with report specifications.
    :param output_format: desired output format, can be word, latex, html, excel
    :param save_location: the path + file name for where to save the output.
    """

    def __init__(self, *, yaml_input: object, output_format: str, save_location: str):
        self.yaml_input = yaml_input
        self.output_format = output_format
        self.save_location = save_location
        self.yaml_tree = None
        self.yaml_trees = []

    def process(self):
        """
        Main processing method that orchestrates the entire document generation pipeline.
        """
        try:
            self._read_yaml()
            self._identify_document_type()
            self._create_object_tree(data=self.yaml_dict[self.main_node], target=self.yaml_tree)
            self.yaml_trees = [self.yaml_tree]
            for yaml_tree in self.yaml_trees:
                self.process_yaml_tree(yaml_tree=yaml_tree)
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Document generation failed. {str(e)}") from e

    def process_yaml_tree(self, yaml_tree):
        """
        Process a single YAML tree through the visitor pipeline.
        """
        try:
            self._initiate_visitors()
            self.main_visitor.generate_content(yaml_tree=yaml_tree)
            self.main_visitor.save_content_docx(save_location=self.save_location, yaml_tree=yaml_tree)
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to process document tree. {str(e)}") from e

    def _read_yaml(self):
        """
        Read and parse the YAML input file.
        """
        try:
            if not self.yaml_input:
                raise Exception("OpenReport Error: No YAML input file specified. Please provide a valid YAML file path.")

            yaml_path = Path(self.yaml_input)

            if not yaml_path.exists():
                raise Exception(f"OpenReport Error: YAML file not found at '{self.yaml_input}'. Please check the file path and ensure the file exists.")

            if not yaml_path.is_file():
                raise Exception(f"OpenReport Error: '{self.yaml_input}' is not a valid file. Please provide a path to a YAML file.")

            if yaml_path.suffix.lower() not in ['.yaml', '.yml']:
                raise Exception(f"OpenReport Error: '{self.yaml_input}' does not appear to be a YAML file. Expected .yaml or .yml extension.")

            with open(yaml_path, "r", encoding='utf-8') as file:
                self.yaml_dict = yaml.safe_load(file)

            if self.yaml_dict is None:
                raise Exception(f"OpenReport Error: YAML file '{self.yaml_input}' is empty or contains no valid data.")

        except yaml.YAMLError as e:
            raise Exception(f"OpenReport Error: Invalid YAML syntax in '{self.yaml_input}'. {str(e)}") from e
        except UnicodeDecodeError as e:
            raise Exception(f"OpenReport Error: Cannot read YAML file '{self.yaml_input}' due to encoding issues. Please ensure the file is saved in UTF-8 format.") from e
        except PermissionError as e:
            raise Exception(f"OpenReport Error: Permission denied when trying to read '{self.yaml_input}'. Please check file permissions.") from e
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to read YAML file '{self.yaml_input}'. {str(e)}") from e

    def _initiate_visitors(self):
        self.main_visitor = WordVisitor()

    def _identify_document_type(self):
        try:
            if not self.yaml_dict:
                raise Exception("OpenReport Error: YAML file contains no data. Please ensure your YAML file has valid content.")

            if not isinstance(self.yaml_dict, dict):
                raise Exception("OpenReport Error: YAML file must contain a dictionary at the root level. Please check your YAML structure.")

            try:
                self.main_node = next(iter(self.yaml_dict))
            except StopIteration:
                raise Exception("OpenReport Error: YAML file is empty or contains no valid keys.")

            if self.main_node == NODES.DOCUMENT:
                self.yaml_tree = DocumentNode()
            else:
                raise Exception(f"OpenReport Error: YAML file must start with '{NODES.DOCUMENT}'. Found '{self.main_node}' instead. Document loop is not supported in the base version.")

        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to identify document type. {str(e)}") from e

    def _create_object_tree(self, data, target):
        try:
            if isinstance(data, dict):
                for key, value in data.items():
                    try:
                        new_node = self._identify_node(key, value)
                        target.add_node(new_node)
                        self._create_object_tree(value, new_node)
                    except Exception as e:
                        if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                            raise
                        raise Exception(f"OpenReport Error: Failed to process YAML node '{key}'. {str(e)}") from e
            elif isinstance(data, list):
                for i, item in enumerate(data):
                    try:
                        self._create_object_tree(item, target)
                    except Exception as e:
                        if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                            raise
                        raise Exception(f"OpenReport Error: Failed to process YAML list item at index {i}. {str(e)}") from e
        except Exception as e:
            if hasattr(e, 'args') and e.args and "OpenReport Error:" in str(e.args[0]):
                raise
            raise Exception(f"OpenReport Error: Failed to create document tree. {str(e)}") from e

    @staticmethod
    def _identify_node(input_item, input_specs):
        if input_item == NODES.DOCUMENT:
            node = DocumentNode()
            node.parent = DocumentNode()
        elif input_item == NODES.STRUCTURE:
            node = StructureNode()
        elif input_item == NODES.DOCUMENT_PARAMS:
            node = DocumentParamsNode()
        elif input_item == NODES.DOCUMENT_STYLE:
            node = DocumentStyleNode()
        elif input_item == NODES.DEFAULT_TEXT_STYLE:
            node = ParagraphStyleNode(default_style_type=NODES.DEFAULT_TEXT_STYLE)
        elif input_item == NODES.DEFAULT_HEADING_STYLE:
            node = ParagraphStyleNode(default_style_type=NODES.DEFAULT_HEADING_STYLE)
        elif input_item == NODES.HEADING:
            node = HeadingNode()
        elif input_item == NODES.TEXT:
            node = TextNode()
        elif input_item == NODES.BULLET_LIST:
            node = BulletListNode()
        elif input_item == NODES.ITEMS:
            node = ItemsNode()
        elif input_item == NODES.MATH_EXPRESSION:
            node = MathExpressionNode()
        elif input_item == NODES.PAGE_BREAK:
            node = PageBreakNode()
        elif input_item == NODES.TABLE_OF_CONTENTS:
            node = TableOfContentsNode()
        elif input_item == NODES.NAME:
            node = NameNode()
            node.add_attribute({ATTRIBUTE_NODES.NAME: input_specs})
        else:
            node = AttributeNode()
            node.add_attribute(attribute={input_item: input_specs})
        return node
