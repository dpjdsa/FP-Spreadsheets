# [Functional Programming Spreadsheets Main.py] in Python

A program to translate functional definitions written in Python to
a CSV file which can be read into an Excel spreadsheet. The spreadsheet
will replicate in Excel formulas what the Python code achieves. The
spreadsheet is also dynamic in that input cells when changed will
automatically lead to recalculation of output cells and values.

## Dependencies

* [Python >=3.6](https://www.python.org/)
* [Treelib](https://treelib.readthedocs.io)


## Evaluation

```
python3 "Functional Programming Spreadsheets Main.py" input_file.txt output
```
Where input_file.txt contains the Python function definitions to be translated
of the the form def f(x): Return(list(1,x)), one function per line. Lines can
be excluded from translation by preceding them with the # character.

Output is written to file output.csv

Parameters are contained in params.py which contains the following 4 lines
NUMFOLDS=50     # Set number of unfolds
Argrow=3        # Set initial row to output variables
Argcol="B"      # Set initial column to output variables
MAXCOL="L"      # Set Column boundary for functions



## Problems

The program has been tested on the following types of function:

* `simple`: Function returning simple expression of its arguments. 
            Example: simple_function(w,x,y,z):return(-w+(x**y/x)%z)
* `range`:  Function returning a list of range derived from input parameters
            Example: def range_function(x,y,z):return(list(range(x,x+y*2,z)))
* `map`:    Function mapping a Lambda function to a list.
            Example: def squareseq(x): return list(map(lambda y:y*y,list(range(1,x))))
* `factors`:Function returning a list of values filtered from a range according to a Lambda function.
            Example: def factors(x): return list(filter(lambda y:(x%y==0),list(range(1,x))))
* `multi`:  Multiple Function with dependencies.
            Example:    def f(x):return x+2
            def g(y): return list(range(10,f(y+1)*f(y)))

Several functions can be defined, one for each line of the input file and these
will be sequentially translated into a section of the CSV file
The CSV file can then be read into Microsoft Excel. The file may need to be reformatted
to enlarge columns and remove wrapping where necessary. However, the file
will dynamically change its results when the cells representing each of the input parameters
is changed.

