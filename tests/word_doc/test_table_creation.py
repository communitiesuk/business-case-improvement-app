from unittest.mock import MagicMock
from apps.word_doc_services.create_document_objects import (
    TABLE_DEFINITION,
    BusinessCaseWordDocument,
    table_footer_word_count
)


def test_create_table_method_adds_table_to_document():
    # arrange
    test_doc = BusinessCaseWordDocument()

    # act
    test_doc.add_generic_table(TABLE_DEFINITION.BREAKDOWN_OF_COST)

    # assert
    assert len(test_doc.doc.tables) == 1


def test_create_table_method_adds_table_footer():
    # arrange
    word_count = "100"
    test_doc = BusinessCaseWordDocument()
    
    # act
    test_doc.add_input_box(word_count)
    test_para = test_doc.doc.paragraphs[0]
    paragraph_count = len(test_doc.doc.paragraphs) + 1
    
    # assert
    assert paragraph_count == 3
    assert test_para.text == table_footer_word_count.format(word_count)


def test_create_table_method_can_add_a_table_for_each_table_definition():
    # arrange
    test_doc = BusinessCaseWordDocument()
    tbl_count: int = 0

    # act 
    for tbl_def in TABLE_DEFINITION:
        print(tbl_def)
        tbl_count +=1
        test_doc.add_generic_table(tbl_def)

    # assert
    assert tbl_count == len(test_doc.doc.tables)

