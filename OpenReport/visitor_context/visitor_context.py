"""
OpenReport Visitor Context Module

This module contains context classes that maintain state information during
document processing. These contexts store shared data like sources, styles,
and parameters that are used across different visitors and processing stages.

Classes:
    VisitorContext: Base context class for visitor pattern state management
    WordVisitorContext: Specialized context for Word document generation
"""


class VisitorContext:
    """
    Base context class for maintaining state during document processing.
    
    This class serves as a container for shared state information that needs
    to be passed between different visitors and processing stages. It stores
    references to processed sources, styles, and parameters to avoid
    recomputation and ensure consistency.
    
    Attributes:
        super_parent: Reference to the top-level parent node (if applicable)
        all_sources (dict): Cache of processed source content indexed by signature
        all_styles (dict): Cache of processed styles and formatting definitions
        all_parameters (dict): Collection of defined parameters and their values
    """
    
    def __init__(self):
        """
        Initialize a new visitor context with empty collections.
        
        Sets up empty dictionaries for caching sources, styles, and parameters,
        and initializes the super_parent reference to None.
        """
        self.super_parent = None
        self.all_sources = {}
        self.all_styles = {}
        self.all_parameters = {}


class WordVisitorContext(VisitorContext):
    """
    Specialized visitor context for Microsoft Word document generation.
    
    This context inherits from VisitorContext and is specifically designed
    for use with the WordVisitor. Currently provides the same functionality
    as the base class but can be extended with Word-specific context data
    if needed in future versions.
    
    Inherits all attributes and methods from VisitorContext:
        - super_parent: Reference to top-level parent node
        - all_sources: Cache of processed source content
        - all_styles: Cache of processed styles
        - all_parameters: Collection of parameters and values
    """
    pass
