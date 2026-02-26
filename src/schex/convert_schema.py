"""
Module for converting Excel schema definitions to JSON schema format.

This module provides classes to parse Excel spreadsheets containing parameter
definitions and convert them into JSON Schema (draft-04) compliant dictionaries.
The main class ConvertSchemaExcelToJson reads schema descriptions from Excel
worksheets and generates corresponding JSON schemas, while the ParseOneRow class
handles the parsing of individual parameter rows with various data types including
enums, numbers, arrays, and file uploads (PDF, JPEG).

Classes:
    ConvertSchemaExcelToJson: Main class for Excel to JSON schema conversion
    ParseOneRow: Helper class for parsing individual parameter rows

Author: Jm Colley
Created: 8 juin 2021
"""

from astropy.constants.codata2010 import e

import json
import logging
import os.path as osp
import re

import pyexcel as px

import schex.tools_params as tpar

logger = logging.getLogger(__name__)


class ConvertSchemaExcelToJson(object):
    """
    classdocs
    """

    def __init__(self, xlsx_filename):
        """
        Constructor

        :param xlsx_filename: name of xlsx file which describes parameters list
        :type xlsx_filename: string

        :Example:test_MasterParameterToJson.py*
        - to import: from ecpi_garage.pipeline.specs.master_parameter import MasterParameterToJson
        - to init: ws = MasterParameterToJson('master_file_params_ecpi_v2.5.xlsx')

        """

        self.used_column = [
            "description",
            "default_value",
            "unit",
            "type",
            "enum_item",
            "array_nb_item",
            "min_max",
        ]

        self.xlsx_file = xlsx_filename
        self.json_dict = {}
        self.json_dict_prefix = {}
        self.schema = ""
        self.json_file = ""

    def get_schema_json(self, sheet_name=None):
        """
        Create schema json associated to schema

        :param sheet_name: name of sheet where schema parameters are defined
        :type sheet_name: str
        :param to_file: to indicate schema json must be saved in file
        :type to_file: bool
        :return: schema json
        :rtype: dict
        """
        if sheet_name is None:
            book = px.get_book(file_name=self.xlsx_file)
            l_sj = []
            for sheet in book.sheet_names():
                if sheet != "Template":
                    print(f"sheet name: {sheet}")
                    l_sj.append(self._get_one_schema_json(sheet))
            return l_sj
        elif isinstance(sheet_name, str):
            return [self._get_one_schema_json(sheet_name)]

    def _get_one_schema_json(self, sheet_name):
        """
        Create schema json associated to schema

        :param sheet_name: name of sheet where schema parameters are defined
        :type sheet_name: str
        :param to_file: to indicate schema json must be saved in file
        :type to_file: bool return l_sj
        :return: schema json
        :rtype: dict
        """
        # 1 download schema description from sheet and build JSON schema dictionary
        # 2 write JSON schema in file if to_file is True
        self.json_dict_prefix = {
            "$schema": "http://json-schema.org/draft-04/schema#",
            "title": f"JSON SCHEMA of {sheet_name} schema",
            "description": f"Parameters for {sheet_name}",
            "type": "object",
            "properties": {},
            "required": [],
        }
        self.schema = self._get_sheet_schema(sheet_name)
        if isinstance(self.schema, px.sheet.Sheet):
            if self._build_json_dict():
                self.write_json()
                return self.json_dict
        return {}

    def _get_sheet_schema(self, schema):
        """get schema ie worksheet description from Excel master parameter file

        :param schema: name of schema ie name of worksheet
        :type schema: string
        :raises TypeError: error if extension of file is not .xlsx
        :raises FileNotFoundError: error if file does not exist
        :raises KeyError: error if schema ie sheet does not exists in booksheet
        :return: schema description extracted from sheet
        :rtype: instance of pyexcel Sheet class
        """
        self.json_file = f"{schema.lower()}_schema.json"
        if self.xlsx_file.rsplit(".")[-1] != "xlsx":
            logger.error(
                f'File {self.xlsx_file} have not the good type. It must be ".xlsx"'
            )
            raise TypeError(
                f'File {self.xlsx_file} have not the good type. It must be ".xlsx"'
            )
            return
        if not osp.exists(self.xlsx_file):
            logger.error(f"File {self.xlsx_file} does not exist")
            raise FileNotFoundError(f"File {self.xlsx_file} does not exist")
            return

        xlsx_file_path = self.xlsx_file
        book = px.get_book(file_name=xlsx_file_path, read_only=True)

        if schema in book.sheet_names():
            sheet = book.sheet_by_name(schema)
            sheet.delete_rows([1])
            # Transformer une ligne donnée (souvent la première) en en-têtes de colonnes,
            # afin de pouvoir accéder aux données par nom plutôt que par index numérique.
            sheet.name_columns_by_row(0)
            sheet.name_rows_by_column(0)
            logger.info(f"init MasterParameterToJson with {self.xlsx_file}/{schema}")
            return sheet
        else:
            logger.error(f"schema {schema.upper()} does not exist in {xlsx_file_path}")
            raise KeyError(
                f"schema {schema.upper()} does not exist in {xlsx_file_path}"
            )

    def _build_json_dict(self):
        """Build JSON schema for each parameter description in self.schema"""

        self.json_dict = self.json_dict_prefix
        #d_properties = {param: {} for param in self.schema.rownames}
        d_properties = {}
        required = []
        #print(self.schema.rownames)
        for param_name in self.schema.rownames:
            if param_name == "" or param_name[0] == "-":
                continue
            d_values = {
                colname: self._get_cell_value(param_name, colname)
                for colname in self.used_column
            }
            d_values["param_name"] = param_name
            logger.info(f"Parameter {param_name} values: {d_values}")
            #print(f"Parameter '{param_name}'")
            d_params = ParseOneRow(d_values)
            if d_params:
                if not d_values["default_value"] != "":
                    required.append(param_name)
                d_properties[param_name] = d_params.__call__()

        self.json_dict["properties"] = d_properties
        self.json_dict["required"] = required
        return True

    def _get_cell_value(self, rowname, colname):
        """Get cell value of schema sheet with row name and column name

        :param rowname: name of row/parameter
        :type rowname: string
        :param colname: name of column/
        :type colname: string
        :raises KeyError: if rowname not in rownames list
        :raises KeyError: if colname not in colnames list
        :return: value of cell
        :rtype: string
        """
        if rowname not in self.schema.rownames:
            raise KeyError(
                f"Parameter '{rowname}' does not exist in {self.schema.rownames}"
            )
            return

        if colname not in self.schema.colnames:
            raise KeyError(
                f"Column '{colncheck_query.pyame}' does not exist in {self.schema.colnames}"
            )
            return
        colvalue = self.schema[rowname, colname]
        logger.debug(f"{rowname},{colname} val: '{colvalue}', type {type(colvalue)}")
        if isinstance(colvalue, str):
            colvalue.replace("\xa0", " ")
            if colname == "description":
                return colvalue.replace("\n", " ")
            else:
                return colvalue.replace(" ", "")
        return colvalue

    def write_json(self, path_dest="", create_name=True):
        """write JSON schema file"""
        logger.info("writing %s", self.json_file)
        if create_name:
            dest = osp.join(path_dest, self.json_file)
        else:
            dest = path_dest
        logger.debug(f"Write file {dest}")

        with open(dest, "w", encoding="utf-8") as text_file:
            data = json.dumps(self.json_dict, ensure_ascii=False, indent=4)
            text_file.write(data)


class ParseOneRow:
    """
    Specific class to parse values of one row ie parameter

    """

    RE_INT = re.compile(r"[0-9-]+")
    RE_NUMBER = re.compile(r"[0-9.-]+")
    RE_STR = re.compile(r"[a-zA-Z0-9_-]+")
    RA_SEP = "],["
    TPL_SEP = "),("

    def __init__(self, dict_values):
        """

        :param dict_values: one row values dictionary of parameters file
        :type dict_values: dict
        """
        self.d_params = {}
        self.parse_row(dict_values)

    def __call__(self):
        """ """
        return self.d_params

    def parse_row(self, d_values):
        """Parse one row of a parameter description

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict
        """
        #print(f"parse row {d_values}")
        value_type = ParseOneRow.RE_STR.findall(d_values["type"])
        d_values["type"] = value_type
        self.d_params = {
            "type": self._parse_type(value_type[0]),
            "description": self._parse_description(d_values),
        }
        logger.debug(f"{d_values}")
        if value_type[0] == "enum":
            self._parse_enum(d_values)
        elif value_type[0] == "str":
            value_type[0] == "string"
        elif value_type[0] == "string":
            pass
        elif value_type[0] in ["number", "integer", "int"]:
            self._parse_number(d_values)
        elif value_type[0] == "pdf":
            self._parse_pdf(d_values)
        elif value_type[0] == "jpg" or value_type[0] == "jpeg":
            self._parse_jpeg(d_values)
        elif value_type[0] == "csv":
            self._parse_txt(d_values)
        elif value_type[0] == "zip":
            self._parse_zip(d_values)
        elif value_type[0] == "array" and value_type[1] == "enum":
            self._parse_array_enum(d_values)
        elif value_type[0] == "array" and (
            value_type[1] in ["integer", "number", "in"]
        ):
            self._parse_array_number(d_values)
        else:
            msg = f"Type {value_type} unknown. Unable to generate JSON schema for {d_values['param_name']}"
            logger.error(msg)
            # error type => stop
            # code out of pipeline can use raise
            raise (TypeError)

    @staticmethod
    def _parse_description(d_values):
        """Check if "unit" value and add it to description if necessary

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict
        :return: description or description [unit]
        :rtype: string
        """
        ret = d_values["description"]
        if d_values["unit"]:
            ret += ". Unit: [" + d_values["unit"] + "]"
        return ret

    @staticmethod
    def _parse_type(type_param):
        """check/parse type value
        check_query.py
                :param type_param: parameter type from column "type"
                :type type_param: string
                :return: same type except for "enum" type
                :rtype: string
        """
        if type_param == "enum":
            return "string"
        elif type_param == "int":
            return "integer"
        elif type_param == "str":
            return "string"
        else:
            return type_param

    @staticmethod
    def _parse_min_max(min_max, typ_min_max):
        """parse min_max values and return dictionary of keywords/values for JSON dictionary

        :param min_max: values of minimum and maximum of a parameter if exists
        :type min_max: string
        :param typ_min_max: type of "minimum" and "maximum" values - number or integer
        :type typ_min_max: string
        :return: dictionary of "minimum" and "maximum" values
        :rtype: dict
        """
        min_max = ParseOneRow.RE_NUMBER.findall(min_max)
        # logger.debug(typ_min_max)
        if typ_min_max == "number":
            d_min_max = {"minimum": float(min_max[0])}
            if len(min_max) > 1:
                d_min_max.update({"maximum": float(min_max[1])})
        if typ_min_max == "integer" or typ_min_max == "int":
            d_min_max = {"minimum": int(min_max[0])}
            if len(min_max) > 1:
                d_min_max.update({"maximum": int(min_max[1])})
        return d_min_max

    @staticmethod
    def _parse_array_nb_item(array_nb_item):
        """parse array_nb_items values and return dictionary of keywords/values for JSON dictionary

        :param array_nb_item: dimensions of array type
        :type array_nb_item: string
        :return: list of dictionary of "minItems" and "maxItems" values
        :rtype: list
        """
        l_dims = array_nb_item.split(ParseOneRow.TPL_SEP)

        # dimensions for d_params values ie minimum et maximum dimensions of array
        dim_ar0 = ParseOneRow.RE_INT.findall(l_dims[0])
        kws_for_d_params = {"minItems": int(dim_ar0[0])}
        if len(dim_ar0) > 1:
            kws_for_d_params.update({"maxItems": int(dim_ar0[1])})
        l_return = [kws_for_d_params]
        # second tuple for d_items values
        if len(l_dims) > 1:
            # "additionalItems": False -> unused to validate JSON schema ? Same result if with or without keyword
            # kws_for_d_item = {"additionalItems": False}
            kws_for_d_item = {}
            dim_ar1 = ParseOneRow.RE_INT.findall(l_dims[1])
            kws_for_d_item.update({"minItems": int(dim_ar1[0])})
            if len(dim_ar1) > 1:
                kws_for_d_item.update({"maxItems": int(dim_ar1[1])})
            l_return.append(kws_for_d_item)
        # return list of
        return l_return

    @staticmethod
    def _parse_array_default_value(d_values):
        """Parse array default value for integer and number

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict
        :return: default value
        :rtype: list or array of list
        """
        m_dict = {}
        m_dict["def"] = d_values["default_value"]
        t_dict = tpar.dictstr_2_dicttyped(m_dict)
        t_values = t_dict["def"]
        return t_values

    def _parse_enum(self, d_values):
        """Parse parameter type which value must be in an "enum" of values

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict

        :Example:

            d_valuetransports dictionary :
            {'type': 'enum'
            'default_value': 'val1',
            'description': 'exemple de paramètre dont la valeur doit être dans une liste « enum »',
            'enum_item': 'val1, val2, val3, val4'}
        """
        if d_values["enum_item"]:
            # self.d_params["enum"] = d_values["enum_item"].split(',')
            self.d_params["enum"] = ParseOneRow.RE_STR.findall(d_values["enum_item"])
        if d_values["default_value"]:
            self.d_params["default"] = d_values["default_value"]

    def _parse_pdf(self, d_values):
        """Parse parameter type which value must be in an "enum" of values

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict

        :Example:

            d_values dictionary :
        """
        self.d_params["type"] = "string"
        self.d_params["contentEncoding"] = "base64"
        self.d_params["contentMediaType"] = "upload:application/pdf"

    def _parse_zip(self, d_values):
        """Parse parameter type which value must be in an "enum" of values

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict

        :Example:

            d_values dictionary :
        """
        self.d_params["type"] = "string"
        self.d_params["contentEncoding"] = "base64"
        self.d_params["contentMediaType"] = "upload:application/zip"

    def _parse_txt(self, d_values):
        """Parse parameter type which value must be in an "enum" of values

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict

        :Example:

            d_values dictionary :
        """
        self.d_params["type"] = "string"
        self.d_params["contentEncoding"] = "base64"
        self.d_params["contentMediaType"] = "upload:text/plain"

    def _parse_jpeg(self, d_values):
        """Parse parameter type which value must be in an "enum" of values

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict

        :Example:

            d_values dictionary :
        """
        self.d_params["type"] = "string"
        self.d_params["contentEncoding"] = "base64"
        self.d_params["contentMediaType"] = "upload:image/jpeg"

    def _parse_number(self, d_values):
        """Parse parameter if type of value is number (float)

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict

        :Example:

            d_values dictionary:
            {'type': 'number',
            'default_value': '17.4',
            'description': 'real type parameter example',
            'min_max': '-10,20'}
        """
        logger.debug(f"{d_values}")
        logger.debug(f"d_params : {self.d_params}")
        if d_values["min_max"]:
            self.d_params.update(
                self._parse_min_max(d_values["min_max"], d_values["type"][0])
            )

        logger.debug(f"{d_values}")
        if d_values["default_value"] != "":
            if d_values["type"][0] == "number":
                try:
                    default_value = float(d_values["default_value"])
                    self.d_params["default"] = default_value
                except TypeError:
                    logger.erro_parse_typer(
                        f"Unable to convert '{d_values['default_value']}' to a 'float'"
                    )
                    self.d_params["default"] = d_values["default_value"]
            else:
                try:
                    default_value = int(d_values["default_value"])
                    self.d_params["default"] = default_value
                except TypeError:
                    logger.error(
                        f"Unable to convert '{d_values['default_value']}' to a 'int'"
                    )
                    self.d_params["default"] = d_values["default_value"]

        logger.debug(f"d_params : {self.d_params}")

    def _parse_array_enum(self, d_values):
        """build JSON schema of an array enum

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict
        :raises KeyError: error if no value for array_nb_item if type of parameter is "array"

        :Example:

            d_values dictionary:
            {'type': 'array,enum',
             'array_nb_item': '(0:)',
             'description': 'exemple de paramètre dont la valeur doit être un tableau « '
                            'enum ». exemple la liste des produits à sortir',
             'enum_item': 'ECL-01,ECL-02,ECL-03,ECL-04,ECL-05'}
        """
        if d_values["array_nb_item"]:
            l_kws_for_d_params_d_item = self._parse_array_nb_item(
                d_values["array_nb_item"]
            )
            if len(l_kws_for_d_params_d_item) > 1:
                d_item = {
                    "type": "array",
                    "items": {"type": self._parse_type(d_values["type"][1])},
                }
                d_item["items"]["enum"] = ParseOneRow.RE_STR.findall(
                    d_values["enum_item"]
                )
                d_item["uniqueItems"] = True
                d_item.update(l_kws_for_d_params_d_item[1])
                self.d_params["items"] = d_item
            else:
                self.d_params["items"] = {"type": self._parse_type(d_values["type"][1])}
                self.d_params["items"]["enum"] = ParseOneRow.RE_STR.findall(
                    d_values["enum_item"]
                )
            self.d_params.update(l_kws_for_d_params_d_item[0])
        else:
            raise KeyError(
                "A value is mandatory for 'array_nb_item' if 'type' of parameter is 'array'"
            )

        if d_values["default_value"]:
            self.d_params["default"] = self._parse_array_default_value(d_values)

    def _parse_array_number(self, d_values):
        """build JSON schema of an array of number/integer

        :param d_values: dictionary of "column name": "value" for a row ie a parameter
        :type d_values: dict
        :raises KeyError: if type of parameter is "array", error if no value for array_nb_item

        :Example:

            d_values dictionary:
            {'type': 'array,integer',
             'array_nb_item': '(1:),(2:2)',
             'description': 'Exemple de tableau a 2 dimensions. Un tableau de couple d'entier',
             'default_value': '[[0,10],[30,40]]',
             'min_max': '0,1023'}
        """
        if d_values["array_nb_item"]:
            if d_values["min_max"]:
                d_min_max = self._parse_min_max(
                    d_values["min_max"], d_values["type"][1]
                )
            else:
                d_min_max = None

            l_kws_for_d_params_d_item = self._parse_array_nb_item(
                d_values["array_nb_item"]
            )
            if len(l_kws_for_d_params_d_item) > 1:
                d_item = {
                    "type": "array",
                    "items": {"type": self._parse_type(d_values["type"][1])},
                }
                if d_min_max:
                    d_item["items"].update(d_min_max)
                d_item.update(l_kws_for_d_params_d_item[1])
                self.d_params["items"] = d_item
            else:
                self.d_params["items"] = {"type": self._parse_type(d_values["type"][1])}
                if d_min_max:
                    self.d_params["items"].update(d_min_max)
            self.d_params.update(l_kws_for_d_params_d_item[0])
        else:
            raise KeyError(
                "A value is mandatory for 'array_nb_item' if 'type' of parameter is 'array'"
            )

        if d_values["default_value"]:
            self.d_params["default"] = self._parse_array_default_value(d_values)

def main():
    import sys
    
    print(sys.argv)
    obj_px = ConvertSchemaExcelToJson(sys.argv[1])
    obj_px.get_schema_json()    

# class MasterParameterToEcpiSchemaJson(MasterParameterToJson):
#     """
#     classdocs
#     """

#     def __init__(self, xlsx_filename):
#         """
#         Constructor

#         :param xlsx_filename: name of xlsx file which describes parameters list
#         :type xlsx_filename: string
#         """
#         super().__init__(xlsx_filename)

#     def create_all_file_schema_json(self):
#         """
#         Create file schema json for all schemas in ECPI master file parameters
#         """
#         root_c = os.path.join(get_root_eclairs(), "ecpi", "process")
#         # root_p = os.path.join(get_root_eclairs(), "ecpi", "pipeline")
#         root_p = get_path_pipeline()
#         d_cpnt = {'general': os.path.join(root_p, "io", "general_schema"),
#                   'dpco': os.path.join(root_c, "dpco", "io", "dpco_schema"),
#                   'cali': os.path.join(root_c, "cali", "io", "cali_schema"),
#                   'bube': os.path.join(root_c, "bube", "io", "bube_schema"),
#                   'imag': os.path.join(root_c, "imag", "io", "imag_schema")
#                    }
#         for cpnt, cpath in d_cpnt.items():
#             self.get_schema_json(cpnt)
#             os.system(f'mv {cpath}.json {cpath}_old.json')
#             self.write_json(cpath + '.json', False)
