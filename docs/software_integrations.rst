.. _software_integrations:



Integrations
============



DataNitro
---------

According to Tony Roberts of PyXll fame `in his blog post <https://www.pyxll.com/blog/tools-for-working-with-excel-and-python/#datanitro>`_ of August 2018, DataNitro is no longer under active development and is not available to license anymore.


ExcelPython
-----------

`ExcelPython homepage <https://github.com/ericremoreynolds/excelpython>`_

ExcelPython is a lightweight COM library which enables you to call Python code and manipulate Python objects from Excel VBA (or indeed any language supporting COM).

https://www.codeproject.com/Articles/639887/Calling-Python-code-from-Excel-with-ExcelPython


ExPy
----

`ExPy homepage <http://www.bnikolic.co.uk/expy/>`_

Licence: BSD-3

The ExPy add-in allows easy use of Python directly from within an Microsoft Excel spreadsheet, both to execute arbitrary code and to define new Excel functions. Features:

* Based on the standard Python interpreter (i.e., not IronPython or other alternatives). Therefore it is fully compatible with all standard Python extensions
* Easy installation -- just unpack and add the DLL to Excel as an Add-In. No registry modification, no installation to system directories
* Define new Excel function at run-time directly from the Excel worksheets
* No COM - based on the pure C-language Excel API

The ExPy add-in is made available to you free-of-charge on this web-site, under the licensing terms detailed below. If you need to integrate Excel and Python we can help! For all enquiries please contact us at webs@bnikolic.co.uk.

The source code: https://github.com/bnikolic/ExPy


pywin32
-------

`pywin32 homepage <https://github.com/mhammond/pywin32>`_

Licence:

PyWin32 uses the Common Object Model (COM) to communicate with MS Windows applications. COM is the Windows infrastructure for intra-application communication. It's been around for a very, very long time (even in "people" years). And it's a two-way conduit.

COM allows a program to "tap" Windows on the shoulder and say, "Hey, if you know a thing called 'Excel', can you start that application? Cheers. Now that it's started, can you please run the method that does X, Y, Z?".

Due to this, it's the reason most of the other options in this category leverage it do do their magic.

It's entirely possible to use this library to integrate Python and Excel without using one of the wrappers but you'll have to do a lot of things manually that the other solutions provide. Stackoverflow has a lot of help on this.

http://timgolden.me.uk/pywin32-docs/contents.html


PyXll
-----
`PyXll homepage <https://www.pyxll.com/>`_

Licence:

PyXLL is an Excel Add-In that enables developers to extend Excel's capabilities with Python code.

PyXLL makes Python a productive, flexible back-end for Excel worksheets, and lets you use the familiar Excel user interface to interact with other parts of your information infrastructure.

With PyXLL, your Python code runs in Excel using any common Python distribution(e.g. Anaconda, Enthought's Canopy or any other CPython distribution from 2.3 to 3.8).

Because PyXLL runs your own full Python distribution you have access to all third party Python packages such as NumPy, Pandas and SciPy and can call them from Excel.

It has a great strength in being able to manage custom buttons on the Ribbon menu.

Example use cases include:

* Calling existing Python code to perform calculations in Excel
* Data processing and analysis that’s too slow or cumbersome to do in VBA
* Pulling in data from external systems such as databases
* Querying large datasets to present summary level data in Excel
* Exposing internal or third party libraries to Excel users

Features:

* Worksheet Functions

  * Argument and Return Types
  * Array Functions
  * Asynchronous Functions
  * Handling Errors
  * Function Documentation
  * Variable Arguments
  * Interrupting Functions

* Using Pandas in Excel
* Menu Functions
* Customizing the Ribbon
* Context Menu Functions
* Macro Functions
* Real Time Data
* Reloading and Rebinding
* Error Handling
* Python as a VBA Replacement
* Distributing Python Code


xlwings Community Edition (CE)
------------------------------

`xlwings CE homepage <https://www.xlwings.org/>`_

Licence:

xlwings CE is a BSD-licensed Python library that makes it easy to call Python from Excel and vice versa:

* `Scripting <https://docs.xlwings.org/en/stable/udfs.html#the-vba-keyword>`_: Automate/interact with Excel from Python using a syntax close to VBA.
* `Macros <https://docs.xlwings.org/en/stable/vba.html#call-python-with-runpython>`_: Replace VBA macros with clean and powerful Python code.
* `UDFs <https://docs.xlwings.org/en/stable/udfs.html>`_: Write User Defined Functions (UDFs) in Python (Windows only).
* `REST API <https://docs.xlwings.org/en/stable/rest_api.html>`_: Expose your Excel workbooks via REST API.

xlwings is a sophisticated COM wrapper.

Numpy arrays and Pandas Series/DataFrames are fully supported. xlwings-powered workbooks are easy to distribute and work on Windows and Mac.

* `Top-Level functions <https://docs.xlwings.org/en/stable/api.html#module-xlwings>`_

  * `view <https://docs.xlwings.org/en/stable/api.html#xlwings.view>`_

* `Object Model <https://docs.xlwings.org/en/stable/api.html#object-model>`_

  * `Apps <https://docs.xlwings.org/en/stable/api.html#apps>`_
  * `App <https://docs.xlwings.org/en/stable/api.html#app>`_
  * `Books <https://docs.xlwings.org/en/stable/api.html#books>`_
  * `Book <https://docs.xlwings.org/en/stable/api.html#book>`_
  * `Sheets <https://docs.xlwings.org/en/stable/api.html#sheets>`_
  * `Sheet <https://docs.xlwings.org/en/stable/api.html#sheet>`_
  * `Range <https://docs.xlwings.org/en/stable/api.html#range>`_
  * `Range Rows <https://docs.xlwings.org/en/stable/api.html#rangerows>`_
  * `Range Columns <https://docs.xlwings.org/en/stable/api.html#rangecolumns>`_
  * `Shapes <https://docs.xlwings.org/en/stable/api.html#shapes>`_
  * `Shape <https://docs.xlwings.org/en/stable/api.html#shape>`_
  * `Charts <https://docs.xlwings.org/en/stable/api.html#charts>`_
  * `Chart <https://docs.xlwings.org/en/stable/api.html#chart>`_
  * `Pictures <https://docs.xlwings.org/en/stable/api.html#pictures>`_
  * `Picture <https://docs.xlwings.org/en/stable/api.html#picture>`_
  * `Names <https://docs.xlwings.org/en/stable/api.html#names>`_
  * `Name <https://docs.xlwings.org/en/stable/api.html#name>`_

* `Extensions <https://docs.xlwings.org/en/stable/extensions.html>`_

  * `In-Excel SQL <https://docs.xlwings.org/en/stable/extensions.html#in-excel-sql>`_

* `MatPlotLib <https://docs.xlwings.org/en/stable/matplotlib.html#id1>`_



xlwings Pro
-----------

`xlwings Pro homepage <https://www.xlwings.org/>`_

Licence:

xlwings PRO adds the following features to xlwings CE:

* Dedicated support via email, phone, screenshare
* Access to video training
* Additional features like embedded code (see below)
* Build zero-configuration installers for easy deployment (see below)
* Access to the reports add-on

Help on the features;

* `Embedded Code <https://docs.xlwings.org/en/stable/deployment.html#embedded-code>`_: Store your Python source code directly in Excel for easy deployment.
* `One-Click Zero-Config Installer <https://docs.xlwings.org/en/stable/deployment.html#one-click-zero-config-installer>`_: Guarantees that the end user does not need to know anything about Python.
* `xlwings Reports <https://docs.xlwings.org/en/stable/api.html#module-xlwings.pro.reports>`_: A template based reporting mechanism, allows business users to change the layout of the report whithout having to change Python code.
* `Plotly static charts <https://docs.xlwings.org/en/stable/matplotlib.html#plotly-static-charts>`_: Support for Plotly static charts.
