.. _software_integrating_calculation:

Integrations with Calculation
=============================

These solutions integrate Excel and Python but have an added feature-set due to providing access to a library that can evaluate Excel formulas and functions in Python.

FlyingKoala
-----------

FlyingKoala is a purpose built Python library and Excel Add-In which adds functionality to the xlwings Excel/Python integration by adding calculation through xlcalculator.

This provides the ability to evaluate Excel formulas in Python (eg; in pandas/numpy/scipy) while you are in Excel and also evaluate Excel formulas in Python where Excel isn't installed.

Key Features;

* Adds the ability to define the behaviour of a UDF using an Excel formula

  * Makes a particular calculation transparent to managers, domain experts and potentially other companies (where you need to share techniques but can't be certain of coding skill).
  * `UDF worked example <https://flyingkoala.readthedocs.io/en/latest/worked_example_horticulture.html>`_

* Unit testing of Excel models either "in Excel" (while it's running) or without Excel (doesn't need to be installed/server)

  * Comprehensively exercising a model written in Excel helps provide evidence that the model operates as you claim
  * Use unit tests to help ensure a Python coded replica of an Excel workbook has correctly replicated the model
  * `Unit test example <https://github.com/bradbase/flyingkoala/tree/master/examples/unit_test_formulas>`_

* Is a great repository for generic UDFs that use specialist Python libraries

  * PVLib for Photovoltaic analysis
  * Pandas for timeseries transformations
  * numpy for differential equations
  * numpy.finance for financial modelling
  * Harvest and Xero for timesheeting, invoicing and accounting integrations


PyCel with PyXll
----------------

It appears to be possible to get Excel/Python integration and calculation working with PyXll using PyCel. There are some notes on the PyCel GitHub page to say this works and provides an explaination as to how.

From the presentation it seems as though is something that contributors to Pycel have created due to having used PyXll. So the level of integration is such that PyXll doesn't need to know about the existence of PyCel.

There's no discernable list of features or benefits except what you can figure for yourself.

From the PyCel GitHub page;
It's possible to run pycel as an excel addin using PyXLL. Simply place pyxll.xll and pyxll.py in the lib directory and add the xll file to the Excel Addins list as explained in the pyxll documentation.
