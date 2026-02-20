"""
OpenReport - YAML-Based Document Generation Tool

OpenReport is a powerful YAML-based tool for creating fully automated and parameterized
documents. It converts YAML specifications into formatted Word (.docx) or PDF documents.

Usage:
    from OpenReport import OpenReportDocumentGenerator

    generator = OpenReportDocumentGenerator(
        yaml_input='specification.yaml',
        output_format='word',
        save_location='./output/'
    )
    generator.process()

For more information, visit: https://openreport.netlify.app/
"""

from OpenReport.base.document_generator import OpenReportDocumentGenerator

__all__ = ["OpenReportDocumentGenerator"]
