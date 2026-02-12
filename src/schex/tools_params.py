"""
Manage spreadsheet format, tools conversion for string value
"""

import json
import logging
import pprint

import pyexcel as px

s_logger = logging.getLogger(__name__)


def cfgparser_to_dicttyped(cfg, section):
    """
    Convert config parser object with section to typed dictionary

    :param cfg: object configparser with section where values are string
    :type cfg: dict
    :param section: section name to convert
    :type section: str
    :return: typed dictionary
    """
    dict_section = cfg._sections[section]
    s_logger.debug(dict_section)
    return dictstr_2_dicttyped(dict_section)


def dictstr_highlevel_to_dicttyped(dict_hl):
    """
    Convert high level dictionary (ie with different component) where
    value are string to typed dictionary

    Example of dict_hl:
     {'CALI': {'array_2D_number': '[[0,10],[30,40]]',
               'array_enum': '[a,b,c]',
               'par_number': '10.9'},
      'DPCO': {'par_path': 'd2'}}

    return value
     {'CALI': {'array_2D_number': [[0, 10], [30, 40]],
               'array_enum': ['a', 'b', 'c'],
               'par_number': 10.9},
     'DPCO': {'par_path': 'd2'} }

    :param dict_hl: high level dictionary with different sections
    :type dict_hl: dict
    :return: typed dictionary
    """
    for key in dict_hl.keys():
        dict_hl[key] = dictstr_2_dicttyped(dict_hl[key])
    return dict_hl


#####################################################################
#
#  spreadsheet features
#
#####################################################################


s_list_col = []


def remove_list_from_value(p_dict):
    """
    Remove list from value and used the first element. Empty list are removed of dictionary

    Motivation : get_sheet() function of pyexcel returns always a list as value, example:
       {'par_enum': ['i1'], 'par_path': ['i2']}
    so this function returns:
       {'par_enum': 'i1', 'par_path': 'i2'}

    :param p_dict: dictionary with list as value.
    :type p_dict: dict
    """
    ret_dict = {}
    for key, value in p_dict.items():
        if key == "":
            continue
        if isinstance(value, list):
            if len(value) > 0:
                value = value[0]
            else:
                value = ""
        if isinstance(value, str):
            value = value.replace(" ", "")
        if value != "":
            ret_dict[key] = value
    return ret_dict


def select_column_gen(current_index, start, limit=-1):
    """
    skip_column_func for get_sheet() function

    :param current_index:
    :type current_index: int
    :param start:
    :type start: int
    :param limit:
    :type limit: int
    """
    import pyexcel_io.constants as constants

    if current_index in s_list_col:
        return constants.TAKE_DATA
    else:
        return constants.SKIP_DATA


def spreadsheet_to_dict(pars_file, sheet_name, l_column_key_val):
    """
    Extraction a dictionary from spreadsheet file.

    :param pars_file: path and name file parameters
    :type pars_file: str
    :param sheet_name:
    :type sheet_name: str
    :param l_column_key_val: list with 2 values : key index, value index
    :type l_column_key_val: list
    """
    global s_list_col
    s_list_col = l_column_key_val
    pars_dict = {}
    try:
        cpnt_dict = px.get_sheet(
            file_name=pars_file,
            sheet_name=sheet_name,
            start_row=1,
            skip_column_func=select_column_gen,
            skip_empty_rows=True,
        )
    except:
        s_logger.warning(f"[{sheet_name}] isn't present")
        return pars_dict
    if l_column_key_val[0] < l_column_key_val[1]:
        cpnt_dict.name_rows_by_column(0)
    else:
        cpnt_dict.name_rows_by_column(1)
    cpnt_dict = cpnt_dict.dict
    cpnt_dict = dict(cpnt_dict)
    s_logger.debug(cpnt_dict)
    cpnt_dict = remove_list_from_value(cpnt_dict)
    if len(cpnt_dict) > 0:
        s_logger.info(f"add section [{sheet_name}]")
        pars_dict[sheet_name] = cpnt_dict
    return pars_dict


class SpreadSheetToParams(object):
    """
    Extract parameters value from spreadsheet to dictionary
    """

    def __init__(self, pars_file, col_key, col_value):
        self.pars_file = pars_file
        self.book = px.get_book(file_name=pars_file)
        self.col_name_key = col_key
        self.col_name_value = col_value
        self.pars_dict = {}

    def get_params_str(self):
        """
        Extract all parameters in all sheet where col_key, col_value are present
        """
        self.pars_dict = {}
        for sheet_name in self.book.sheet_names():
            s_logger.info(f"Examen sheet {sheet_name}")
            sheet = self.book[sheet_name]
            if sheet.number_of_columns() == 0:
                s_logger.warning(f"skip empty sheet.")
                continue
            sheet.name_columns_by_row(0)
            l_colnames = sheet.colnames
            s_logger.debug(l_colnames)
            if self.col_name_key not in l_colnames:
                s_logger.warning(
                    f"Column '{self.col_name_key}' isn't present in sheet :"
                    " {sheet.name}."
                )
                continue
            if self.col_name_value not in l_colnames:
                s_logger.warning(
                    f"Column '{self.col_name_value}' isn't present in sheet :"
                    " {sheet.name}."
                )
                continue
            l_keep_col = [
                l_colnames.index(self.col_name_key),
                l_colnames.index(self.col_name_value),
            ]
            s_logger.debug(l_keep_col)
            dict_pars = spreadsheet_to_dict(self.pars_file, sheet_name, l_keep_col)
            s_logger.info(f"{len(dict_pars)} params in sheet {sheet_name}")
            s_logger.debug(dict_pars)
            self.pars_dict.update(dict_pars)
        s_logger.debug(f"Final pars_dict:\n{pprint.pformat(self.pars_dict)}")
        return self.pars_dict

    def get_params_typed(self):
        return dictstr_highlevel_to_dicttyped(self.get_params_str())


#####################################################################
#
#  low level function conversion
#
#####################################################################


def add_double_quote_in_list_string(str_list, rm_space=False):
    """
    Automaton adding double quote in string list for each element, example
     * '[[abb,e,1],[c,3.14]]'
    returns
     * '[["abb","e","1"],["c","3.14"]]'

     ...note::
        Motivation: with double quote we can apply json.loads()
        function to convert string list in list.

    :param str_list: representation of a list as a string
    :type str_list: str
    :param rm_space: remove space character before processing
    :type rm_space: bool
    """
    if rm_space:
        str_list.replace(" ", "")
    list_char = ["[", ",", "]"]
    new_str = ""
    in_quote = False
    for let in str_list:
        if not in_quote:
            if let not in list_char:
                new_str += f'"{let}'
                in_quote = True
            else:
                new_str += let
        elif let not in list_char:
            new_str += let
        else:
            new_str += f'"{let}'
            in_quote = False
    return new_str


def dictstr_2_dicttyped(dict_str):
    """
    Convert a dictionary where all values are a string to typed dictionary
    if conversion is possible with json.loads() function

    Motivation: python library configparser gives value as string

    :param dict_str: dict where values are string
    :type dict_str: dict
    :return: typed dictionary
    """
    dict_typed = {}
    for key, pval in dict_str.items():
        dict_typed[key] = pval
        if not isinstance(pval, str):
            # no conversion necessary !?
            continue
        pval = pval.replace(" ", "")
        new_val = pval
        try:
            if "." in pval:
                new_val = float(pval)
            else:
                new_val = int(pval)
        except:
            pass

        if pval[0] == "[":
            if pval.count("[") != pval.count("]"):
                s_logger.error(
                    f"Number of '[' and ']' is different in parameters {key}: '{pval}'"
                )
                return None
            try:
                new_val = json.loads(pval)
            except:
                new_val = None

            if not new_val:
                # json.loads manage only double quote
                pval_m = pval.replace("'", '"')
                try:
                    new_val = json.loads(pval_m)
                except:
                    new_val = None

            if not new_val:
                pval_quote = add_double_quote_in_list_string(pval)
                try:
                    new_val = json.loads(pval_quote)
                except:
                    s_logger.error(
                        f"Problem to convert parameters {key}: '{pval}'."
                        " Check '[' and ']' order ?"
                    )
                    return None
        dict_typed[key] = new_val
    return dict_typed
