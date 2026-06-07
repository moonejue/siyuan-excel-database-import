[简体中文](https://github.com/moonejue/siyuan-excel-database-import/blob/main/README_zh_CN.md)

# Excel Database Import

Import `.csv`, `.xlsx`, or `.xls` files directly into a SiYuan database.

The logo follows the Moon Teacher brand style: moon-white, misty sage, warm gold, soft contrast, and generous breathing room.

## Features

- Open the importer from the top-right corner of a database.
- Select a worksheet and map columns before importing.
- Match existing database columns by header name.
- Create missing text columns or skip columns manually.
- Skip rows with an empty primary key.
- Import in batches of 100 rows.

## Usage

1. Open a SiYuan document containing a database.
2. Click `Excel 导入` in the top-right corner of the database.
3. Select a file and worksheet.
4. Review the mappings and click `开始导入`.

Excel files are read with [SheetJS Community Edition](https://sheetjs.com/) `0.18.5`. See `LICENSE.xlsx.txt`.
