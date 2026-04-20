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

import logging
import re

import schex.tools_params as tpar

logger = logging.getLogger(__name__)

# regular expressions for parsing values of parameters
RE_INT = re.compile(r"[0-9-]+")
RE_NUMBER = re.compile(r"[0-9.-]+")
RE_STR = re.compile(r"[a-zA-Z0-9_-]+")
RA_SEP = "],["
TPL_SEP = "),("


class ParseOneRow:
    """
    Specific class to parse values of one row ie parameter

    """

    def __init__(self, dict_values):
        """

        :param dict_values: one row values dictionary of parameters file
        :type dict_values: dict
        """
        self.d_pars = dict_values
        self.d_json = {}
        self.type2func = {
            "string": self._parse_string,
            "str": self._parse_string,
            "integer": self._parse_number,
            "int": self._parse_number,
            "number": self._parse_number,
            "enum": self._parse_enum,
            "pdf": self._parse_pdf,
            "jpg": self._parse_jpeg,
            "jpeg": self._parse_jpeg,
            "csv": self._parse_txt,
            "zip": self._parse_zip,
            "array,enum": self._parse_array_enum,
            "array,integer": self._parse_array_number,
            "array,int": self._parse_array_number,
            "array,number": self._parse_array_number,
        }
        self.parse_row()

    def parse_row(self):
        """Parse one row of a parameter description

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict
        """
        type_attribut = self.d_pars["type"].replace(" ", "")
        self.d_json = {}
        logger.debug(f"{self.d_pars}")
        if type_attribut in self.type2func.keys():
            self.d_pars["type"] = RE_STR.findall(type_attribut)
            d_json = self.type2func[type_attribut]()
            d_json["description"] = self._parse_description()
        else:
            msg = f"Type {type_attribut} unknown. Unable to generate JSON schema for {self.d_pars['param_name']}"
            logger.error(msg)
            raise (TypeError)
        self.d_json = d_json
        return d_json

    def _parse_description(self):
        """Check if "unit" value and add it to description if necessary

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict
        :return: description or description [unit]
        :rtype: string
        """
        ret = self.d_pars["description"]
        if self.d_pars["unit"]:
            ret += ". Unit: [" + self.d_pars["unit"] + "]"
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
        min_max = RE_NUMBER.findall(min_max)
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
        l_dims = array_nb_item.split(TPL_SEP)
        # dimensions for d_params values ie minimum et maximum dimensions of array
        dim_ar0 = RE_INT.findall(l_dims[0])
        kws_for_d_params = {"minItems": int(dim_ar0[0])}
        if len(dim_ar0) > 1:
            kws_for_d_params.update({"maxItems": int(dim_ar0[1])})
        l_return = [kws_for_d_params]
        # second tuple for d_items values
        if len(l_dims) > 1:
            # "additionalItems": False -> unused to validate JSON schema ? Same result if with or without keyword
            # kws_for_d_item = {"additionalItems": False}
            kws_for_d_item = {}
            dim_ar1 = RE_INT.findall(l_dims[1])
            kws_for_d_item.update({"minItems": int(dim_ar1[0])})
            if len(dim_ar1) > 1:
                kws_for_d_item.update({"maxItems": int(dim_ar1[1])})
            l_return.append(kws_for_d_item)
        # return list of
        return l_return

    def _parse_array_default_value(self):
        """Parse array default value for integer and number

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict
        :return: default value
        :rtype: list or array of list
        """
        m_dict = {}
        m_dict["def"] = self.d_pars["default_value"]
        t_dict = tpar.dictstr_2_dicttyped(m_dict)
        t_values = t_dict["def"]
        return t_values

    def _parse_enum(self):
        """Parse parameter type which value must be in an "enum" of values

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict

        :Example:

            d_valuetransports dictionary :
            {'type': 'enum'
            'default_value': 'val1',
            'description': 'exemple de paramètre dont la valeur doit être dans une liste « enum »',
            'enum_item': 'val1, val2, val3, val4'}
        """
        if self.d_pars["enum_item"]:
            # self.d_params["enum"] = self.d_pars["enum_item"].split(',')
            self.d_json["enum"] = RE_STR.findall(self.d_pars["enum_item"])
        if self.d_pars["default_value"]:
            self.d_json["default"] = self.d_pars["default_value"]

    def _parse_pdf(self):
        """Parse parameter type which value must be in an "enum" of values

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict

        :Example:

            self.d_pars dictionary :
        """
        self.d_json["type"] = "string"
        self.d_json["contentEncoding"] = "base64"
        self.d_json["contentMediaType"] = "upload:application/pdf"

    def _parse_zip(self):
        """Parse parameter type which value must be in an "enum" of values

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict

        :Example:

            self.d_pars dictionary :
        """
        self.d_json["type"] = "string"
        self.d_json["contentEncoding"] = "base64"
        self.d_json["contentMediaType"] = "upload:application/zip"

    def _parse_string(self):
        """ """
        self.d_json["type"] = "string"

    def _parse_txt(self):
        """Parse parameter type which value must be in an "enum" of values

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict

        :Example:

            self.d_pars dictionary :
        """
        self.d_json["type"] = "string"
        self.d_json["contentEncoding"] = "base64"
        self.d_json["contentMediaType"] = "upload:text/plain"

    def _parse_jpeg(self):
        """Parse parameter type which value must be in an "enum" of values

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict

        :Example:

            self.d_pars dictionary :
        """
        self.d_json["type"] = "string"
        self.d_json["contentEncoding"] = "base64"
        self.d_json["contentMediaType"] = "upload:image/jpeg"

    def _parse_number(self):
        """Parse parameter if type of value is number (float)

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict

        :Example:

            self.d_pars dictionary:
            {'type': 'number',
            'default_value': '17.4',
            'description': 'real type parameter example',
            'min_max': '-10,20'}
        """
        logger.debug(f"{self.d_pars}")
        logger.debug(f"d_params : {self.d_json}")
        if self.d_pars["min_max"]:
            self.d_json.update(
                self._parse_min_max(self.d_pars["min_max"], self.d_pars["type"][0])
            )

        logger.debug(f"{self.d_pars}")
        if self.d_pars["default_value"] != "":
            if self.d_pars["type"][0] == "number":
                try:
                    default_value = float(self.d_pars["default_value"])
                    self.d_json["default"] = default_value
                except TypeError:
                    logger.error(
                        f"Unable to convert '{self.d_pars['default_value']}' to a 'float'"
                    )
                    self.d_json["default"] = self.d_pars["default_value"]
            else:
                try:
                    default_value = int(self.d_pars["default_value"])
                    self.d_json["default"] = default_value
                except TypeError:
                    logger.error(
                        f"Unable to convert '{self.d_pars['default_value']}' to a 'int'"
                    )
                    self.d_json["default"] = self.d_pars["default_value"]

        logger.debug(f"d_params : {self.d_json}")

    def _parse_array_enum(self):
        """build JSON schema of an array enum

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict
        :raises KeyError: error if no value for array_nb_item if type of parameter is "array"

        :Example:

            self.d_pars dictionary:
            {'type': 'array,enum',
             'array_nb_item': '(0:)',
             'description': 'exemple de paramètre dont la valeur doit être un tableau « '
                            'enum ». exemple la liste des produits à sortir',
             'enum_item': 'ECL-01,ECL-02,ECL-03,ECL-04,ECL-05'}
        """
        if self.d_pars["array_nb_item"]:
            l_kws_for_d_params_d_item = self._parse_array_nb_item(
                self.d_pars["array_nb_item"]
            )
            if len(l_kws_for_d_params_d_item) > 1:
                d_item = {
                    "type": "array",
                    "items": {"type": self._parse_type(self.d_pars["type"][1])},
                }
                d_item["items"]["enum"] = RE_STR.findall(self.d_pars["enum_item"])
                d_item["uniqueItems"] = True
                d_item.update(l_kws_for_d_params_d_item[1])
                self.d_json["items"] = d_item
            else:
                self.d_json["items"] = {
                    "type": self._parse_type(self.d_pars["type"][1])
                }
                self.d_json["items"]["enum"] = RE_STR.findall(self.d_pars["enum_item"])
            self.d_json.update(l_kws_for_d_params_d_item[0])
        else:
            raise KeyError(
                "A value is mandatory for 'array_nb_item' if 'type' of parameter is 'array'"
            )

        if self.d_pars["default_value"]:
            self.d_json["default"] = self._parse_array_default_value()

    def _parse_array_number(self):
        """build JSON schema of an array of number/integer

        :param self.d_pars: dictionary of "column name": "value" for a row ie a parameter
        :type self.d_pars: dict
        :raises KeyError: if type of parameter is "array", error if no value for array_nb_item

        :Example:

            self.d_pars dictionary:
            {'type': 'array,integer',
             'array_nb_item': '(1:),(2:2)',
             'description': 'Exemple de tableau a 2 dimensions. Un tableau de couple d'entier',
             'default_value': '[[0,10],[30,40]]',
             'min_max': '0,1023'}
        """
        if self.d_pars["array_nb_item"]:
            if self.d_pars["min_max"]:
                d_min_max = self._parse_min_max(
                    self.d_pars["min_max"], self.d_pars["type"][1]
                )
            else:
                d_min_max = None

            l_kws_for_d_params_d_item = self._parse_array_nb_item(
                self.d_pars["array_nb_item"]
            )
            if len(l_kws_for_d_params_d_item) > 1:
                d_item = {
                    "type": "array",
                    "items": {"type": self._parse_type(self.d_pars["type"][1])},
                }
                if d_min_max:
                    d_item["items"].update(d_min_max)
                d_item.update(l_kws_for_d_params_d_item[1])
                self.d_json["items"] = d_item
            else:
                self.d_json["items"] = {
                    "type": self._parse_type(self.d_pars["type"][1])
                }
                if d_min_max:
                    self.d_json["items"].update(d_min_max)
            self.d_json.update(l_kws_for_d_params_d_item[0])
        else:
            raise KeyError(
                "A value is mandatory for 'array_nb_item' if 'type' of parameter is 'array'"
            )

        if self.d_pars["default_value"]:
            self.d_json["default"] = self._parse_array_default_value()
