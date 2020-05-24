.. _software_calculation:


Calculation (evaluating Excel formulas in Python)
=================================================

These tools help calculate the result of an Excel equation within Python, using the functions as defined in an Excel file without the need for Excel to be installed.

Formulas
--------

Formulas has come from an engieering background.

`Formulas homepage <https://github.com/vinci1it2000/formulas>`_

Licence:

formulas implements an interpreter for Excel formulas, which parses and compile Excel formulas expressions.

Moreover, it compiles Excel workbooks to python and executes without using the Excel COM server. Hence, Excel is not needed.

Supported functions;

* AND
* ARABIC
* ASIN
* ASINH
* ATAN
* ATAN2
* ATANH
* AVERAGE
* AVERAGEA
* AVERAGEIF
* BIN2DEC
* CEILING
* CEILING.MATH
* CEILING.PRECISE
* COLUMN
* CONCAT
* CONCATENATE
* COS
* COSH
* COT
* COTH
* COUNT
* COUNTA
* COUNTBLANK
* COUNTIF
* CSC
* CSCH
* DATEVALUE
* DAY
* DEC2BIN
* DEC2HEX
* DEC2OCT
* DECIMAL
* DEGREES
* EVEN
* EXP
* FACT
* FACTDOUBLE
* FALSE
* FIND
* FLOOR
* FLOOR.MATH
* FLOOR.PRECISE
* GCD
* HEX2DEC
* HLOOKUP
* HOUR
* IF
* IFERROR
* INDEX
* INT
* IRR
* ISERR
* ISERROR
* ISO.CIELING
* LARGE
* LCM
* LEFT
* LEN
* LOG10
* LOG
* LOOKUP
* LOWER
* LN
* MATCH
* MINUTE
* MAX
* MID
* MIN
* MOD
* MONTH
* MROUND
* NA
* NOT
* NOW
* NPV
* OCT2DEC
* ODD
* OR
* PI
* POWER
* RADIANS
* RAND
* RANDBETWEEN
* REPLACE
* RIGHT
* ROMAN
* ROUND
* ROUNDDOWN
* ROUNDUP
* ROW
* SEC
* SECH
* SECOND
* SIGN
* SIN
* SINH
* SMALL
* SQRT
* SQRTPI
* SUMPRODUCT
* SUM
* SUMIF
* SWITCH
* TAN
* TANH
* TODAY
* TIME
* TIMEVALUE
* TRIM
* TRUE
* TRUNC
* UPPER
* VLOOKUP
* XIRR
* XNPV
* XOR
* YEAR
* YEARFRAC


Koala
-----
`Koala homepage <https://github.com/vallettea/koala>`_

Licence:

Koala converts any Excel workbook into a python object that enables on the fly calculation without the need of Excel.

Koala parses an Excel workbook and creates a network of all the cells with their dependencies. It is then possible to change any value of a node and recompute all the depending cells.

You can read more on the origins of Koala `here <https://github.com/vallettea/koala/blob/master/doc/presentation.md>`_.

Supported Functions;

* ALL
* AND
* ARRAY
* ARRAYROW
* ATAN2
* AVERAGE
* CHOOSE
* COLUMNS
* CONCAT
* CONCATENATE
* COUNT
* COUNTA
* COUNTIF
* COUNTIFS
* DATE
* EOMONTH
* GAMMALN
* IF
* IFERROR
* INDEX
* IRR
* ISBLANK
* ISNA
* ISTEXT
* LINEST
* LOG
* LOOKUP
* LN
* MATCH
* MAX
* MID
* MIN
* MOD
* MONTH
* NPV
* OFFSET
* OR
* PI
* PMT
* POWER
* RAND
* RANDBETWEEN
* RIGHT
* ROUND
* ROUNDUP
* ROWS
* SLN
* SQRT
* SUM
* SUMIF
* SUMIFS
* SUMPRODUCT
* TAN
* TODAY
* VALUE
* VDB
* VLOOKUP
* XIRR
* XLOG
* XNPV
* YEAR
* YEARFRAC

Pandas
------

Pandas does not do this. To do this you need to write code to read the functions and map them to Pandas, numpy, numpy.finance or scipy functions which is the service the other solutions in this category offer.


PyCel
-----

`PyCel Homepage <https://github.com/dgorissen/pycel>`_

Licence:

Pycel is a small python library that can translate an Excel spreadsheet into executable python code which can be run independently of Excel.

The python code is based on a graph and uses caching & lazy evaluation to ensure (relatively) fast execution. The graph can be exported and analyzed using tools like Gephi. See the contained example for an illustration.

The full motivation behind pycel including some examples & screenshots is described in this `blog post <https://dirkgorissen.com/2011/10/19/pycel-compiling-excel-spreadsheets-to-python-and-making-pretty-pictures/>`_.

It's possible to run pycel as an excel addin using `PyXLL <https://www.pyxll.com/>`_.

Supported Functions;

* Abs
* Atan2
* Average
* Averageif
* Averageifs
* Cieling
* Cieling.Math
* Cieling.Precise
* Count
* Countif
* CountIfs
* Even
* Fact
* FactDouble
* Floor
* Floor.math
* Floor.precise
* Int
* IsErr
* IsError
* IsEven
* IsText
* IsNa
* IsOdd
* IsNumber
* Large
* Linest
* Ln
* Log
* Max
* Maxifs
* Min
* Minifs
* Mod
* NPV
* Odd
* Power
* Round
* Rounddown
* Roundup
* Sign
* Small
* Sum
* Sumif
* Sumifs
* SumProduct
* Trunc


xlcalculator
------------

`xlcalculator homepage <https://github.com/bradbase/xlcalculator>`_

Licence:

xlcalculator converts a given Excel workbook into a Python object (model) that enables calculation (evaluation) without the need of Excel.

It uses the xlfunctions library for the Python implementation fo Excel functions.

* Loading an Excel file into a Python compatible state
* Saving Python compatible state
* Loading Python compatible state
* Ignore worksheets
* Extracting sub-portions of a model. "focussing" on provided cell addresses or defined names
* Evaluating cells, ranges defined names and formulas
