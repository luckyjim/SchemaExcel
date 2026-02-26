# SchemaExcel

## Installation

* Upgrading pip and setuptools packages can help

```bash
pip install --upgrade setuptools pip
```

* Installation from GitHub

```bash
python -m pip install git+https://github.com/luckyjim/SchemaExcel.git 
```

* Update version

```bash
python -m pip uninstall SchemaExcel
python -m pip install git+https://github.com/luckyjim/SchemaExcel.git 
 ```


## Features

SchemaExcel allows you to define schemas using an Excel sheet or other spreadsheet software in XLSX format.
Each sheet defines a schema, and the `schex_convert` executable will generate JSON schema files for each sheet in the XLSX file.

Writing is simpler and more concise, similar to defining multimedia files (MIME type). However, not all JSON schema syntax is supported—see the Limitations section for details.

## Example

Here a example file [schema_example.xlsx](example/schema_example.xlsx) which defined one schema with the sheet "FOR_TEST". Launch conversion with `schex_convert <file.xlsx>` :


```bash
$ schex_convert example/schema_example.xlsx
sheet name: FOR_TEST
 ```

the result is the file [for_test_schema.json](example/for_test_schema.json)


## Sheet syntax

In your sheet define 8 columns:
* name : name of property
* description : string
* default_value : The absence of a default value implies a mandatory property
* unit : string
* type : string, number, integer, array, enum, csv, pdf, jpeg
* enum_item : 
* array_nb_item : specific syntax see file example
* min_max: two number


## Limitations

Array must has same type and the type can be only basic type and not user type.

## Similar project

[schemasheets](https://linkml.io/schemasheets/)

