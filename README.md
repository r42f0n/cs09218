java c
Advanced Excel
Module 2 Lecture
Module 2   –lookup   function and data   Table management
Outline
Part 1 Lookup Functions
-            Vlookup
•          Exact lookup
•         Approximate lookup
-            HLookup
-            XLookup
Part 2 Sorting and Filtering Using   Tables
-            Sort by one or more columns
-            Filter by color
-            Filter   with conditions
-            Text Filter
-            Format   table as Table and rename
Part 1 Lookup   functions
File: Part 1 - Basic lookup   functions.xlsx
VLOOKUP is a powerful Excel   function   that allows   you   to search for a   value in   the   first column of   a   table and return a   value in   the same row   from a specified column.
Vlookup   will lookup the Lookup_value in   the 1st column of the Table_array, and   try   to   find   the   value   that matches   with   the Lookup_value , once   found, return   the cells content in   the
col_index_num   with an approximate or an exact match (Range_lookup)
VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)
Lookup_Value:
•          Is   the   value   to be   found   Table_Array:
•          Is a   table of data in   which data is retrieved
Table_array can be a reference to a range or   to a   table name   Col_Index_num:
•          Is   the column number in   table_array   from   which   the matching   value   should be returned
e.g.   the   first column of   values in   the   table is column 1   Range_lookup:
•          Is a logical   value:   to   find   the closest match in the   first column of   the   table (sorted in   ascending order) =   TRUE or 1 or to   find an exact match = False or 0
Worksheet: Vlookup (Exact   match)
Task 1: Retrieve Name and Price information corresponding to   the ID using   vlookup
o   C17: =VLOOKUP(C15,A3:C12,2,FALSE)
o   C18: =VLOOKUP(C15,A3:C12,3,0)
Worksheet: Vlookup (Approximate match)
Task 2: Find Discount rate (D4:D13)   for each   price   from   the discount table (H3:I8)   o   D4: =VLOOKUP(C4,$H$4:$I$8,2,TRUE)
Note:
•         there is no exact match of   the price   to   the discount rate   table.
-          Example: For Price = 675.18,   excel   will go through   the discount   table and try   to   find   the interval   that contains   this   value. First, it   tries   to   find   the step   value (600)   that is         lower   than   the Price, and   the next step   value (900) is higher than the Price, then
returns   the corresponding discount rate   within this interval,   that is 10%.
•          If   the price is -100$,   the   output   will be #N/A since Excel cannot   find a step   valueless   than -   100$.
•          Remember   to lock   the   table_array   with $ sign
or
You may also use a named   table,   Table_Discount   for H5:I8
-            Method 1: Format an excel   Table
         Select any cell   within   the   table, or range of cells   you   want   to   format as a   table
         On   the Home   tab, click Format as   Table
            Keep one cell active on   the   table
         Table tab > change the name in   Table Name field
-            Method 2: Define name
         Select   the cell, range of cells that   you   want   to name
         On   the Formulas   tab, click Define name
         Type   the name in Name field
Task 3: Retrieve Grade (D17:D27) for each mark   from grade table (H17:I27)
o   Name grade table H17:I27 as TableGradeScale
o   D17: =VLOOKUP(C17,TableGradeScale,2,TRUE)
Note:   When implementing an approximate match,   you need   to define   your searching table_array   in ascending order!
If   the grade scale   table is   in descending order,   you'll get   the   following:
because excel is unable   to search   for approximate   values in any order other   than ascending order
   
Now   the   table is horizontal.   We can use Hlookup
HLOOKUP(lookup_value, table_array, row_index_num,[range_lookup])   Task   4: Retrieve Name and Price information given   the ID using Hlookup
o   C7: =HLOOKUP(C6,$B$2:$K$4,2,FALSE)
o   C8: =HLOOKUP(C6,$B$2:$K$4,3,FALSE)
Worksheet:   Xlookup
X   lookup   will research   the Lookup_value in the range of   the Lookup_array, and   try   to find   the   value   that matches   with   the Lookup_value ,   once   found, return   the cells content   within   the
Return_array   when the Lookup_value is   not   found, If_not_found   value is returned   and the
research will be done   with a specific Match_mode method and   with a specific   Search_mode
XLOOKUP(lookup_value, Lookup_array, Return_array, If_not_found, Match_mode)   Lookup_value:
•          Is   the   value   to search   for   Lookup_array:
•          Is   the array or range   to search   Return_array:
•          Is the array or range   to   search   If_not_found:
•          Is returned if no match is   found   Match_mode:
•          Specify how   to match Looup_value against   the   values in Lookup_array   0:   exact match, -1: exact match or next smaller, 1: exact or next larger,   2: wildcard character match)
Search_mode:
•         Specify   the search mode   to use
1: search   first   to last,   -1: search last   to   first
Benefits of using   Xlookup:
            It can use a lot of different match mode
            It can search   for both horizontal and   vertical data
            It can perform. a reverse search
            It can return entire rows and columns of data instead of a single   value
            It can include   the "if not   found" argument
Task 5: Retrieve Name and Country information   given   the ID using   Xlookup
o   Method 1: using range
B3: =XLOOKUP(A3;C6:C15;A6:B15;"ID not found";0;1)
o  代 写Advanced Excel Module 2 –lookup function and data Table management
代做程序编程语言 Method 2: using   table array
Assuming   that the cell range   A5:D15 has been   formatted as a   table   with   the name
Table_Employee
B3: =XLOOKUP(A3;Table_Employee[Emp ID];Table_Employee[[Employee   Name]:[Country]];"ID not   found";0;1)
Part 2   Sorting, filtering, Using   Tables
File: Part 2 - Sorting,   filtering,   tables.xlsx

Task 1:   Sort   by   CustomerID   in   descending   order
o   Method   1:
•          Select   the   entire   table   (Ctrl+A)   >   Data tab>   in   Sort      Filter   Click   on   Sort   
•         Select   CustomerID,   on   what   you   want   to   sort   and   the   order
   
Note:   Excel   will   automatically   identify   if   there   is   a   header.
☑My   data   has   header
be   sorted.
o   Method   2:
•          Format   range   A1:H202   as   a   table:
•          Select   one   cell   in   the   table
•          In   Home   tab,   click   Format   as   table   button
you   get   automatically   into   the   header   row   the   filter   button   you   can   sort   the   column   with   the   order   you   want
•          Click   filter   button   of   the   column   CustomerID
•         Then   select   Sort Largest   to Smaller
   
Task 2: Sort   table   with two columns (levels): LastName ascending, FirstName descending
Note:   you must use Sort dialog box (see   Task1, Method1) because   when   you use   twice   the Filter   button Excel   will retain only   the last sort   you did
   
Task 3: Sort or Filter by color
o   Click on   the   ‘Filter’ button of   the column CustomerID
o   Select Sort by color or Filter by   color
   
Note:   The rows are not in consecutive   numbers.   The   funnel icon      in   the filter button shows   that   filtering has been done in   that column.
Task   4: Clear Filter
o   Method 1:
•         Click   the   funnel icon      of   the column   for   which   you   want   to clear   the   filter
•         Select Clear Filter
This method clears only   the   filter on   this column
   
o   Method 2:
•          Data   tab> in Sort  Filter Click on Clear   限clear   This method clear all columns   filters
Important: it is possible   to clear a   filter BUT it is not possible   to clear a sort.   Task 5: Number Filter   with CustomerID greater than 100
o   Click on   the   ‘Filter’ button of   the column CustomerID
o   Select Number Filters and Greater   Than   …
o   Then   type 100
   
Task 6:   Text Filter all customers living in New   York
o   Click on   the   ‘Filter’ button of   the column City
o   Type New   York in the Searchbox
   
Task 7:   Text Filter all customers   with LastName end   with ‘on’
o   Click on   the   ‘Filter’ button of   the column LastName
o   Select   Text Filters and Ends   with   …
o   Type on
   
Task 8:   Text Filter all customers   with   Address contains 'road'
o   Click on   the   ‘Filter’ button of   the column   Address
o   Type road in Searchbox
   
Task 9:   Text Filter all customers   with 'a' as second letter of LastName
o   Click on   the   ‘Filter’ button of   the column LastName
o   Select   Text Filters and Begins   with   …
o   Type   ?a
Note: here, ? is   what   we call a   wildcard, it represent any single character
   
Task 10: Format the ‘Customers table’   as   Table   If it's not already done:
o   Select one cell into   the table
Warning: DO NOT select   the   whole   worksheet,   the   table range is   A1:I202
o   Select   the   table region > Ribbon   ‘   Home’   >   ‘Styles’   >   ‘Format as   Table’
Now if   you click on any cell, a new ribbon ‘Table Design’   will appear.   Try   the   following operation:
•          Rename   the   Table as ‘Table_Customer’
You'll   find   Table name on   the left side of the   Table Design   tab
   
•         Add a new column 'FullName' at   the end of   the table   Excel resizes automatically   the   table
•          In 'FullName' column, Create a   formula   to concatenate FirstName and LastName   columns,   you'll add a space as separator
Note:   when   you click into   the cell, Excel   table recognizes automatically column names   you should get   the   following   formula: =[@FirstName]" "[@LastName]
•         Add   to   the bottom of   the   table a   total row   to count   the number of customers
Tips: Check   Total row option in   Table Design   tab, and   you get automatically   the count at   the bottom of   the last column
Note:   you could change   the calculation using   the   drop-down list at   the right side of   the   cell

Worksheets   ‘Products’, ‘Orders’   and   ‘Orders   products list’
Task 9: Change   Table design   for   worksheets   ‘Products’, ‘Orders’   and ‘Orders’ products list’, and   name   them ‘TableProducts’, ‘TableOrders’   and ‘TableOrderDetails’ respectively.
Worksheet   ‘Orders   products list’
Task 10: Add new column 'Price' after   the last column   Note   that   the   table is automatically resized,
and calculate   the 'Price' =   ‘Price per Unit’   *   ‘Quantity’
   
Note   that   the   formula   takes into account   the   table column names,
[@[Price per Unit]]*[@Quantity] and not   the cell references   when clicking into   the cell, C2*D2   You may also notice   that   the   formula is automatically copied   to   the bottom of   the   table
Task 11:   Add new column 'Final Price' after   the last column and calculate as   ‘Price’   –   ‘Price’   *   ‘Discount’
Task 12: In Ribbon   ‘Table Design’, Remove duplicates      with   the same ProductName
   
Task 13: In Ribbon   ‘Table Design’, insert Slicer      of ProductName and select all 'Chocolate'   products
   

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
