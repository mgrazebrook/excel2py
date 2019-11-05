# Create a Python module with the same calculation as an Excel file
This converts an Excel spreadsheet, which may contain many sheets, external 
functions and database look-ups, into a Python program. It's good for things 
like actuarial or quant models.

It creates a class library implementing the spreadsheet's calculation. 
Customise it by subclassing. It's designed so that if the original 
spreadsheet changes and you regenerate the code, it affects the customisation 
as little as possible. 

## How to use it

Excel must be running: excel2py uses the .Net interface not only to extract 
the content but also to try out functions it doesn't know 
(see "Working with in-house library code")

### Changes you could make to the input spreadsheet

By default, it looks for sheets called "Inputs" and "Results". You can 
use these names in your spreadsheet or name them on the command line. 

The generated code uses cell names where they exist, otherwise it generates
names from cell references and sheet names. The generated code is much nicer 
to work with if you do. In Excel, the 'Define Name' function in the 
"Name manager" section of the "Formulas" tab can generate names from 
descriptions.

Neither is necessary but it makes the generated code nicer to use.

### Customising the calculation
Common reasons for customisation include replacing hard-coded tables with 
database lookup or a section of the calculation with a library 
implementation of that functionality. 

The generated code creates a property method for each cell with a 
calculation. When run, the code calculates each cell exactly once. 

So if you have a library function which sets many values, override
all the values you set with the same custom function. The custom function
calls the library function and sets all the output variables.

### The configuration file

TODO:

### Working with in-house library code

If excel2py finds functions it doesn't know, it creates a default 
implementation and a set of unit tests.

The default implementation is a lookup table for the values in the sample
spreadsheet. The lookup table can be extended by sub-classing. This is 
useless for production but handy for demos. When run, it prints a warning.

It also experiments with possible inputs by calling the function in Excel.
It uses this to create a default set of unit tests for test driven 
development.

### Limitations

It's currently designed for a single calculation, rather than running a 
calculation on each of a table of values. It would be practical to enhance 
it to do that, I just didn't need to when I wrote it.

To re-run the calculation, "reset()" resets everything. If you change a 
single input value, it does not recalculate. 