#! /usr/bin/env python3

import sys

from schex.convert_schema import ConvertSchemaExcelToJson

print(sys.argv)

obj_px = ConvertSchemaExcelToJson(sys.argv[1])
m_dict = obj_px.get_schema_json()

