from enum import IntEnum

class table_definition(IntEnum):
    proposal_participants_and_organisational_details = 0


# Defines how many Rows and Columns in the corresponding table.
# Key: enum of the named sections table
# Value: tuple representing Row | Column values. (order is the same as the add_table() method in docx)
TABLE_DEFINITIONS = [
        {table_definition.proposal_participants_and_organisational_details: (8, 2)},
]

TABLE_CONTENT = [
    {
        table_definition.proposal_participants_and_organisational_details: [
            {
                "row": 1,
                "column": 1,
                "value": "Primary author name:"
            },
            {
                "row": 2,
                "column": 1,
                "value": "Primary author team:"
            },
            {
                "row": 3,
                "column": 1,
                "value": "Case contributors and /or reviewers:",
            },
            {
                "row": 3,
                "column": 2,
                "italic": True,
                "value": "[If applicable, list anyone who helped draft, develop or review the business case before it was submitted for approval. Do not include subject matter experts (SMEs) or SRO/SCS, as these named will be captured separately elsewhere in this document]"
            }
        ]
    }
]
