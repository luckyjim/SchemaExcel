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


import json
import logging
import os.path as osp
import re

import pyexcel as px

import schex.tools_params as tpar
from schex.parse_row import ParseOneRow

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
        self.sheet =px.sheet.Sheet()
        self.json_file = ""

    def write_schema_all_sheets(self, sheet_name=None):
        """
        Create schema json associated to schema

        """
        if sheet_name is None:
            book = px.get_book(file_name=self.xlsx_file)
            l_sj = []
            for sheet in book.sheet_names():
                if sheet != "Template":
                    print(f"sheet name: {sheet}")
                    l_sj.append(self.write_schema_one_sheet(sheet))
            return l_sj
        elif isinstance(sheet_name, str):
            return [self.write_schema_one_sheet(sheet_name)]

    def write_schema_one_sheet(self, sheet_name):
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
        self.sheet = self._read_schema_sheet(sheet_name)
        if isinstance(self.sheet, px.sheet.Sheet):
            if self._build_json():
                self._write_json()
                return self.json_dict
        return {}

    def _read_schema_sheet(self, schema):
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

    def _build_json(self):
        """Build JSON schema for each parameter description in self.schema
        
        1. loop on attribut (ie row)
            1. loop on attribut specification (ie on column)
                1. get cell value with row name and column name
            2. build JSON schema with ParseOneRow class and cell values
            3. add attribut to required list if default value is empty
        2. add properties and required list to JSON schema dictionary 
        """

        self.json_dict = self.json_dict_prefix
        #d_properties = {param: {} for param in self.schema.rownames}
        d_properties = {}
        required = []
        #print(self.schema.rownames)
        for param_name in self.sheet.rownames:
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
        # end of attribut loop
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
        if rowname not in self.sheet.rownames:
            raise KeyError(
                f"Parameter '{rowname}' does not exist in {self.sheet.rownames}"
            )
            return

        if colname not in self.sheet.colnames:
            raise KeyError(
                f"Column '{colname}' does not exist in {self.sheet.colnames}"
            )
            return
        colvalue = self.sheet[rowname, colname]
        logger.debug(f"{rowname},{colname} val: '{colvalue}', type {type(colvalue)}")
        if isinstance(colvalue, str):
            colvalue.replace("\xa0", " ")
            if colname == "description":
                return colvalue.replace("\n", " ")
            else:
                return colvalue.replace(" ", "")
        return colvalue

    def _write_json(self, path_dest="", create_name=True):
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


def main():
    import sys
    
    print(sys.argv)
    obj_px = ConvertSchemaExcelToJson(sys.argv[1])
    obj_px.write_schema_all_sheets()    
