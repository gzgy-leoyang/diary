import diary 
import pytest

###
def test_get_workbook():
    assert diary.get_workbook( '123' ) == None
    # test_name = ["123","1av","eae","xyz"]
    # for name in test_name :
    #     print(name)
    #     assert diary.get_workbook( name ) == None
    
    # test_name = ["123.xlsx","1av.xlsx","eae.xlsx","xyz.xlsx"]
    # for name in test_name :
    #     print(name)
    #     assert diary.get_workbook( name ) != None

def test_parser_config():
    assert diary.parser_config() != None
    assert diary.parser_config("") == None
    assert diary.parser_config("dd") == None
    assert diary.parser_config("dd.in") == None
    assert diary.parser_config("config.ini") != None


