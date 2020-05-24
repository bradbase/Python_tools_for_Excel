.. _software_file_tools:


File Tools
==========

There are a number of tools for Python which help manage reading and writing Excel files. Most have been around for a long time and are mature.

The below list of them came from the website http://www.python-excel.org/.


openpyxl
--------

`openpyxl homepage <https://openpyxl.readthedocs.io/en/stable/>`_

Licence: MIT/Expat

Openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

It was born from lack of existing library to read/write natively from Python the Office Open XML format.

It is a comprehensive library to create, modify and save Excel files using operations akin to Excel itself.


Features:

- `Cells <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.cell.html>`_
- `Charts <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.chart.html>`_
- `Chartsheets <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.chartsheet.html>`_
- `Comments <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.comments.html>`_
- `Descriptors <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.descriptors.html>`_
- `Drawing <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.drawing.html>`_
- `Formatting <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.formatting.html>`_
- `Pivot Tables <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.pivot.html>`_
- `Read Excel files <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.reader.html>`_
- `Styles <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.html>`_
- `Workbook operations <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.workbook.html>`_
- `Worksheet operations <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.html>`_
- `Write Excel files <https://openpyxl.readthedocs.io/en/stable/api/openpyxl.writer.html>`_




Pandas
------

`pandas homepage <https://pandas.pydata.org/docs/user_guide/index.html#user-guide>`_

Licence:

pandas is an open source, BSD-licensed library providing high-performance, easy-to-use data structures and data analysis tools for the Python programming language.

It is comprehensive for data analyisis and plays well with numpy and scipy. It facilitates a huge range of operations for data analysis. The focus of this section of this review is Excel file operations so I will only list the related aspect. pandas may well be mentioned elsewhere in this review if it does things related to that section.

The read_excel() method can read Excel 2003 (.xls) files using the xlrd Python module. Excel 2007+ (.xlsx) files can be read using either xlrd or openpyxl. Binary Excel (.xlsb) files can be read using pyxlsb. The to_excel() instance method is used for saving a DataFrame to Excel. Generally the semantics are similar to working with csv data. See the cookbook for some advanced strategies.

To write a DataFrame object to a sheet of an Excel file, you can use the to_excel instance method. The arguments are largely the same as to_csv, the first argument being the name of the excel file, and the optional second argument the name of the sheet to which the DataFrame should be written. Files with a .xls extension will be written using xlwt and those with a .xlsx extension will be written using xlsxwriter (if available) or openpyxl.

Features:

* `IO tools (xls, xlsx, xlsb, text, CSV, HDF5, ...) <https://pandas.pydata.org/docs/user_guide/io.html>`_

  * `Excel Files <https://pandas.pydata.org/docs/user_guide/io.html#excel-files>`_

    * `Reading Excel files <https://pandas.pydata.org/docs/user_guide/io.html#reading-excel-files>`_

      * `Specifying sheets <https://pandas.pydata.org/docs/user_guide/io.html#specifying-sheets>`_
      * `Reading a MultiIndex <https://pandas.pydata.org/docs/user_guide/io.html#reading-a-multiindex>`_
      * `Parsing specific columns <https://pandas.pydata.org/docs/user_guide/io.html#parsing-specific-columns>`_
      * `Parsing dates <https://pandas.pydata.org/docs/user_guide/io.html#parsing-dates>`_
      * `Cell converters <https://pandas.pydata.org/docs/user_guide/io.html#cell-converters>`_
      * `Dtype specifications <https://pandas.pydata.org/docs/user_guide/io.html#dtype-specifications>`_

    * `Writing Excel files <https://pandas.pydata.org/docs/user_guide/io.html#writing-excel-files>`_

      * `Writing Excel files to disk <https://pandas.pydata.org/docs/user_guide/io.html#writing-excel-files-to-disk>`_
      * `Writing Excel files to memory <https://pandas.pydata.org/docs/user_guide/io.html#writing-excel-files-to-memory>`_
      * `Excel writer engines <https://pandas.pydata.org/docs/user_guide/io.html#excel-writer-engines>`_
      * `Style and formatting <https://pandas.pydata.org/docs/user_guide/io.html#style-and-formatting>`_


Pyxlsx
------

`pyxlsx homepage <https://github.com/fortfall/pyxlsx>`_

Licence:

A package to read/write xlsx worksheet like dictionary, based on openpyxl.

* Create a new xlsx file and write to it
* Open an existing xlsx file
* Append rows to a worksheet
* Read from / write to a worksheet by row
* Read from a worksheet by column
* Read cell directly from Worksheet, Header, ContentRow
* Read adjacent cells of a certain cell


xlrd
----

`xlrd homepage <https://xlrd.readthedocs.io/en/latest/>`_

Licence:

xlrd is a library for reading data and formatting information from Excel files, whether they are .xls or .xlsx files.

* `Handling of Unicode <https://xlrd.readthedocs.io/en/latest/unicode.html>`_
* `Dates in Excel spreadsheets <https://xlrd.readthedocs.io/en/latest/dates.html>`_
* `Named references, constants, formulas, and macros <https://xlrd.readthedocs.io/en/latest/references.html>`_
* `Formatting information in Excel Spreadsheets <https://xlrd.readthedocs.io/en/latest/formatting.html>`_
* `Loading worksheets on demand <https://xlrd.readthedocs.io/en/latest/on_demand.html>`_
* `XML vulnerabilities and Excel files <https://xlrd.readthedocs.io/en/latest/vulnerabilities.html>`_
* `API Reference <https://xlrd.readthedocs.io/en/latest/api.html>`_

XlsxWriter
----------

`xlsxwriter homepage <https://xlsxwriter.readthedocs.io/>`_

Licence:

XlsxWriter is a Python module that can be used to write text, numbers, formulas and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file. It supports features such as formatting and many more, including:

* 100% compatible Excel XLSX files.
* Full formatting.
* Merged cells.
* Defined names.
* Charts.
* Autofilters.
* Data validation and drop down lists.
* Conditional formatting.
* Worksheet PNG/JPEG/BMP/WMF/EMF images.
* Rich multi-format strings.
* Cell comments.
* Textboxes.
* Integration with Pandas.
* Memory optimization mode for writing large files.

It supports Python 2.7, 3.4+ and PyPy and uses standard libraries only.

* `Workbook <https://xlsxwriter.readthedocs.io/workbook.html>`_
* `Worksheet <https://xlsxwriter.readthedocs.io/worksheet.html>`_
* `Worksheet (Page Setup) <https://xlsxwriter.readthedocs.io/page_setup.html>`_
* `Format <https://xlsxwriter.readthedocs.io/format.html>`_
* `The Chart <https://xlsxwriter.readthedocs.io/chart.html>`_
* `The Chartsheet <https://xlsxwriter.readthedocs.io/chartsheet.html>`_
* `Working with Cell Notation <https://xlsxwriter.readthedocs.io/working_with_cell_notation.html>`_
* `Working with and Writing Data <https://xlsxwriter.readthedocs.io/working_with_data.html>`_
* `Working with Formulas <https://xlsxwriter.readthedocs.io/working_with_formulas.html>`_
* `Working with Dates and Time <https://xlsxwriter.readthedocs.io/working_with_dates_and_time.html>`_
* `Working with Colors <https://xlsxwriter.readthedocs.io/working_with_colors.html>`_
* `Working with Charts <https://xlsxwriter.readthedocs.io/working_with_charts.html>`_
* `Working with Object Positioning <https://xlsxwriter.readthedocs.io/working_with_object_positioning.html>`_
* `Working with Autofilters <https://xlsxwriter.readthedocs.io/working_with_autofilters.html>`_
* `Working with Data Validation <https://xlsxwriter.readthedocs.io/working_with_data_validation.html>`_
* `Working with Conditional Formatting <https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html>`_
* `Working with Worksheet Tables <https://xlsxwriter.readthedocs.io/working_with_tables.html>`_
* `Working with Textboxes <https://xlsxwriter.readthedocs.io/working_with_textboxes.html>`_
* `Working with Sparklines <https://xlsxwriter.readthedocs.io/working_with_sparklines.html>`_
* `Working with Cell Comments <https://xlsxwriter.readthedocs.io/working_with_cell_comments.html>`_
* `Working with Outlines and Grouping <https://xlsxwriter.readthedocs.io/working_with_outlines.html>`_
* `Working with Memory and Performance <https://xlsxwriter.readthedocs.io/working_with_memory.html>`_
* `Working with VBA Macros <https://xlsxwriter.readthedocs.io/working_with_macros.html>`_
* `Working with Python Pandas and XlsxWriter <https://xlsxwriter.readthedocs.io/working_with_pandas.html>`_


xlutils
-------

`xlutils homepage <https://xlutils.readthedocs.io/en/latest/>`_

Licence:

This package provides a collection of utilities for working with Excel files. Since these utilities may require either or both of the xlrd and xlwt packages, they are collected together here, separate from either package. The utilities are grouped into several modules within the package, each of them is documented below:

* `xlutils.copy <https://xlutils.readthedocs.io/en/latest/copy.html>`_

  * Tools for copying xlrd.Book objects to xlwt.Workbook objects.

* `xlutils.display <https://xlutils.readthedocs.io/en/latest/display.html>`_

  * Utility functions for displaying information about xlrd-related objects in a user-friendly and safe fashion.

* `xlutils.filter <https://xlutils.readthedocs.io/en/latest/filter.html>`_

  * A mini framework for splitting and filtering existing Excel files into new Excel files.

* `xlutils.margins <https://xlutils.readthedocs.io/en/latest/margins.html>`_

  * Tools for finding how much of an Excel file contains useful data.

* `xlutils.save <https://xlutils.readthedocs.io/en/latest/save.html>`_

  * Tools for serializing xlrd.Book objects back to Excel files.

* `xlutils.styles <https://xlutils.readthedocs.io/en/latest/styles.html>`_

  * Tools for working with formatting information expressed the styles found in Excel files.

* `xlutils.view <https://xlutils.readthedocs.io/en/latest/view.html>`_

  * Easy to use views of the data contained in a workbook’s sheets.


xlwt
----

`xlwt homepage <https://xlwt.readthedocs.io/en/latest/>`_

Licence:

xlwt is a library for writing data and formatting information to older Excel files (ie: .xls)

* `add_sheet <https://xlwt.readthedocs.io/en/latest/api.html#xlwt.Workbook.Workbook.add_sheet>`_
* `save <https://xlwt.readthedocs.io/en/latest/api.html#xlwt.Workbook.Workbook.save>`_
* `write <https://xlwt.readthedocs.io/en/latest/api.html#xlwt.Worksheet.Worksheet.write>`_

Formatting

* Number format
* Font
* Alignment
* Border
* Background
* Protection
