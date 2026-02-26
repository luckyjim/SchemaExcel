import os.path as osp

import pyexcel as px
import pytest

from schex.convert_schema import ConvertSchemaExcelToJson
import schex.tools_schema_json as tc


DATA_PATH = "data"

FILE_XLSX = "schema_example.xlsx"

dict_OK = {
    "par_enum": "val3",
    "par_path": "ecpi_garage",
    "par_number": 19.99,
    "array_enum": ["ECL-01"],
    "array_2D_number": [[10, 20], [40, 50]],
    "array_2D_enum": [["sta"]],
    "array_2D_energy": [[5.2, 10.4], [20.1, 50.53]],
    "array_number": [5.2, 10.4, 12.5],
}


def test_conversion():
    path_master = osp.join(DATA_PATH, FILE_XLSX)
    obj_px = ConvertSchemaExcelToJson(path_master)
    m_dict = obj_px.get_schema_json("FOR_TEST")
    assert m_dict != {}
    obj_px.write_json(DATA_PATH)


def test_for_test_component():
    path_master = osp.join(DATA_PATH, FILE_XLSX)
    obj_px = ConvertSchemaExcelToJson(path_master)
    m_dict = obj_px.get_schema_json("FOR_TEST")
    assert m_dict != {}
    obj_px.write_json(DATA_PATH)
    path_schema = osp.join(DATA_PATH, "for_test_schema.json")
    assert tc.check_add_default_file(path_schema, dict_OK)
    dict_OK["array_2D_number"] = [[10, 20]]
    assert tc.check_add_default_file(path_schema, dict_OK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["par_enum"] = "val99"
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["par_path"] = 99
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["par_number"] = 20.01
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["par_number"] = -10.01
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["par_number"] = "10.0"
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["array_enum"] = ["ECL-01", "ECL-06"]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_enum"] = "ECL-01"
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["array_2D_number"] = [[12, 21], [51, 5078]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_number"] = [
        [10, 20],
        [30, 40],
        [
            50,
        ],
    ]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_number"] = [[-1, 10]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_number"] = [[10.1, 20]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["array_2D_enum"] = [["sta", 20]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_enum"] = ["sta"]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_enum"] = [[], ["sta"]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_enum"] = [["sta", "reo"]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_enum"] = [["sta", "sta"]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["array_2D_energy"] = [[5.2, 10.4], [20.1, 150.01]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_energy"] = [[5.2, 10.4], [-20.1, 50.53]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_energy"] = [[5.2, 10.4], [20.1, 50.0], [20.1, 150.01]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_energy"] = [
        [5.2, 10.4],
        [
            20.1,
        ],
    ]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_energy"] = [[5.2, 10.4], 20.1]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_2D_energy"] = [[5.2, 10.4], [20.1, "50.0"]]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK = dict(dict_OK.items())
    dict_nOK["array_number"] = [5.2, -10.4, 20.1]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_number"] = [5.2, 10.4, 30.1]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_number"] = [5.2, -10.1, 20.1, 19.99]
    assert not tc.check_add_default_file(path_schema, dict_nOK)
    dict_nOK["array_number"] = [
        5.2,
        -10.1,
    ]
    assert not tc.check_add_default_file(path_schema, dict_nOK)


def nok_test_a_get_component():
    path_master = osp.join(DATA_PATH, "toto.xls")
    with pytest.raises(TypeError):
        obj = ConvertSchemaExcelToJson(path_master)
        obj.get_schema_json("FOR_TEST")
    path_master = osp.join(DATA_PATH, "toto.xlsx")
    with pytest.raises(FileNotFoundError):
        obj = ConvertSchemaExcelToJson(path_master)
        obj.get_schema_json("FOR_TEST")
    with pytest.raises(TypeError):
        path_master = osp.join(DATA_PATH, ".xlsx")
        obj = ConvertSchemaExcelToJson(path_master)
        obj.get_schema_json("FOR_TEST")
