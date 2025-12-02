# Excel
Formulas and VBA code are kept here

Abbreviate-Millions-to-M.txt
----
The formula formats a number on the order of millions 
that it appears as currency, in number of millions, and
ends with a capital M. 

For example, 3860000 would be converted to $3.86M.

To apply, select the cell or cells you want to re-format. 
Then right-click to get a menu and choose Format Cells.

In the Format Cells popup, choose the Number tab, then 
choose Custom from the Category list. Enter the formula
**$0.00,,"M"** in the Type box, then press OK. This should
apply your formula. 

Phone-Number-Formula.txt
----
The formula assumes 10 digits and any combination of 
non-digits to create a value in (999) 999-9999 format, 
which I will call the 'nice' format. This is commonly 
used in the US, and other countries in the so-called
North American Numbering Plan (NANP).

On a side note, calling any country that is part of 
the North American Numbering Plan (NANP) from the US 
is the done the same way as making a domestic 
long-distance call. No international exit code 
(like 011) is needed. 

So the 'nice' format would be valid for Canada, Jamaica, 
and the Dominican Republic for example. 

The formula was created using a Claude.ai prompt. 

This was used in a worksheet with column headers in row
1 and data beginning in row 2. Column C has 'raw' phone 
number data; column D has the desired format. 

I applied the formula to 100 rows by saving the formula 
to D2, then pressing Ctrl-C to put that cell's content 
in the clipboard. I then selected 99 more rows of cells 
in column D so that the formula would be applied from D3 
to D101. 

If a cell in column C was blank, then the corresponding
cell in column D would also be blank. 

If your raw phone number data does not start in C2, then
edit Phone-Number-Formula.txt and replace every occurrance 
of C2 with the first cell designation containing the raw 
phone number. 

I have not exhaustively tested this. The most common
phone number formats I get are:

999999999 | 999.999.9999 | 999-999-9999

So far they get formatted to the nice format as I wanted, 
so I saved this for re-use. 
