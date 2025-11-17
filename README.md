# Excel
Formulas and VBA code are kept here

Phone-Number-Formula.txt
________________________________________________________________
The formula assumes 10 digits and any combination of 
non-digits to create a value in (999) 999-9999 format, 
which I will call the 'nice' format.

The formula was created using a Claude.ai prompt. 

This was used in a worksheet with column headers in row
1 and data beginning in row 2. Column C has 'raw' phone 
number data; column D has the desired format. 

Put the formula in D2 and then copy/paste it down to as 
many rows as you like. 

I did that by saving the formula to D2, then pressing 
Ctrl-C to put that cell's content in the clipboard. I 
then selected 99 more rows of cells in column D so that
the formula would be applied from D3 to D101. 

If a cell in column C was blank, then the corresponding
cell in column D will also be blank. 

If your raw phone number data does not start in C2, then
edit Phone-Number-Formula.txt and replace every occurrance 
of C2 with the first cell designation containing the raw 
phone number. 

I have not exhaustively tested this. The most common
phone number formats I get are:

999999999
999.999.9999
999-999-9999

So far they get formatted to the nice format as I wanted, 
so I saved this for re-use. 
