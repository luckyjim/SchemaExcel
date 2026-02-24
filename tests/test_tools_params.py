import pprint
import json

from schex.tools_params import (
    remove_list_from_value,
    add_double_quote_in_list_string,
    dictstr_2_dicttyped,
    SpreadSheetToParams,
)


def test_remove_list_from_value():
    md = {"a": [""], "b": 12, "c": ["xx", "yy"], "d": [], "e": [23]}
    clear_dict = remove_list_from_value(md)
    pprint.pprint(clear_dict)
    assert clear_dict == {"b": 12, "c": "xx", "e": 23}


def test_add_cote_in_list_string():
    str_l = "[[abb,e,1],[c,3.14]]"
    ret = add_double_quote_in_list_string(str_l)
    assert ret == '[["abb","e","1"],["c","3.14"]]'


def test_dictstr_2_dicttyped():
    dictstr = {
        "int": "10",
        "float": "10.1",
        "str": "toto",
        "list_num": "[10,11]",
        "list2D_num": "[[10,11],[1,2]]",
        "list_str": "[a,b,c]",
        "list2D_str": "[[a,b,c],[e,f]]",
    }
    r_dict = dictstr_2_dicttyped(dictstr)
    t_dict = {
        "int": 10,
        "float": 10.1,
        "str": "toto",
        "list_num": [10, 11],
        "list2D_num": [[10, 11], [1, 2]],
        "list_str": ["a", "b", "c"],
        "list2D_str": [["a", "b", "c"], ["e", "f"]],
    }
    assert t_dict == r_dict


def test_dictstr_2_dicttyped_quote():
    dictstr = {"a_str": "['NEO', 'PEO']"}
    r_dict = dictstr_2_dicttyped(dictstr)
    t_dict = {"a_str": ["NEO", "PEO"]}
    assert t_dict == r_dict


# def test_SpreadSheetToParams():
#     name_xlsx = add_path_current_module(__file__, "params.xlsx")
#     print(name_xlsx)
#     file_pars = SpreadSheetToParams(name_xlsx, "nom", "data")
#     pars = file_pars.get_params_typed()
#     t_pars = {
#         "sheet_a": {"a": 2, "b": "toto", "c": [1, 2]},
#         "sheet_b": {"a": 2, "b": "titi", "c": [1, 2]},
#     }
#     print(pars)
#     assert t_pars == pars


#############################################
# prototype test/exploration
#############################################


def ptest_add_cote_in_list_string():
    sl_01 = "[[a,e,1],[c,3]]"
    add_double_quote_in_list_string(sl_01)

    sl_01 = "[[a,e,1],[c,3],ABB,12.4]"
    add_double_quote_in_list_string(sl_01)

    sl_01 = "[[a,e,1],[c,3],ABB,12.4]"
    ret = add_double_quote_in_list_string(sl_01)
    print(ret)
    print(json.loads(ret))

    sl_01 = "[[a,e,v],[c,a,v],[a,v,f,r]]"
    ret = add_double_quote_in_list_string(sl_01)
    print(ret)
    print(json.loads(ret))
