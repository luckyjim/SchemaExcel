#! /usr/bin/env python3

import sys

from schex.convert_schema import ConvertSchemaExcelToJson

if __name__ == "__main__":
    print(sys.argv)
    obj_px = ConvertSchemaExcelToJson(sys.argv[1])
    obj_px.get_schema_json()