from enum import IntEnum, auto
from dataclasses import dataclass, field

from docx.shared import Cm

# Enum to distinguish which table to create
class TableDefinition(IntEnum):
    DEVELOP_AND_SUPPORT_PROPOSAL = auto()
    EXPECTED_BENEFITS = auto()
    SINGLE_CELL_TEXT_BOX = auto()

@dataclass
class _TextRunData:
    text: str
    italic: bool = False
    bold: bool = False

@dataclass
class _ParagraphData:
    runs: list[_TextRunData] = field(default_factory=list)

@dataclass
class _CellData:
    row: int
    column: int
    paragraphs: list[_ParagraphData] = field(default_factory=list)

@dataclass
class _TableContent:
    definition: TableDefinition
    row_heights: list[Cm] = field(default_factory=list)
    column_widths: list[Cm] = field(default_factory=list)
    cells: list[_CellData] = field(default_factory=list)

# Each table definition should be added here
# Width of table total is 15.9cm currently
TABLE_REGISTRY = [
    _TableContent(
        definition=TableDefinition.DEVELOP_AND_SUPPORT_PROPOSAL,
        row_heights=[Cm(5.08), Cm(1.27), Cm(2.03), Cm(1.27), Cm(1.27), Cm(1.27), Cm(1.27), Cm(1.27)],
        column_widths=[Cm(4.89), Cm(11.01)],
        cells=[
            _CellData(row=1, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Title:",
                        bold=True
                    )
                ])    
            ]),
            _CellData(row=2, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Primary author team:",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=3, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Case contributors and /or reviewers:",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=3, column=2, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="[If applicable, list anyone who helped draft, develop or review the business case before it was submitted for approval. Do not include subject matter experts (SMEs) or SRO/SCS, as these named will be captured separately elsewhere in this document]",
                        italic=True
                    )
                ])
            ]),
            _CellData(row=4, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Project title:",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=5, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Programme title:",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=6, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Portfolio area:",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=7, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Directorate:",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=8, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Type of business case:",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=8, column=2, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Business Justification Case - Procurement",
                        italic=True
                    )
                ])
            ])
        ]
    ),
    _TableContent(
        definition=TableDefinition.EXPECTED_BENEFITS,
        row_heights=[Cm(1.27), Cm(1.27)],
        column_widths=[Cm(5.3), Cm(5.3), Cm(5.3)],
        cells=[
            _CellData(row=1, column=1, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Benefit",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=1, column=2, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="Value/positive impact",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=1, column=3, paragraphs=[
                _ParagraphData(runs=[
                    _TextRunData(
                        text="When will this benefit be realised?",
                        bold=True
                    )
                ])
            ]),
            _CellData(row=2, column=1)
        ]
    ),
    _TableContent(
        definition=TableDefinition.SINGLE_CELL_TEXT_BOX,
        row_heights=[Cm(2.5)],
        column_widths=[Cm(15.9)],
        cells=[
            _CellData(row=1, column=1)
        ]   
    )
]
