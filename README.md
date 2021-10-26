# stock-analysis
stock-analysis

## Project Goals

1) Finding a single stock's total daily volume and return
2) Finding the total daily volume and yearly return for multiple stocks
3) Creating performance analysis on AllStocksAnalysis()


## VBA Learnings Goals

- Creating subroutings (aka macros)
- Knowing how to activate a worksheet
- Understanding how to create headings
- Initializing variables and arrays
- Counting the number of cells in a given row
- Processing table data using for-loops and nested for-loops
- Calculating stock volume and returns
- Formatting analysis for readability
- Creating Command Buttons to run VBA code
- Clearing a worksheet

## Subroutines/Macros 
DQAnalysis() - Generating a report containing a single stocks volume and return.
SkillDrillNo1() - Filling cells of a spreadsheet with values in relationship to the cell
AllStocksAnalysis() - Processing multiple stocks to find the volume and return of each and report the results.
SkillDrillNo2() - Practicing nested for loops using VBA
formatAllStocksAnalysis() - Using VBA to format the All Stocks Analysis
ClearWorksheet() - Clears any worksheet

*** Each sheet contains buttons to clear the worksheet and run the corresponding VBA code

## Stock Calculations: 
Volume: Total number of shares traded within a given period of time.

Return: Percentage difference in price from the beginning of a given
        period to the end of another given period.
Objects: "Things" that can be manipulated by methods.
Example: A cell in Excel.

Properties: predefined variables holding values about the object.
Example: A cell in Excel can have a value. 

Methods: A collection of instructions used to "do things" to objects.
Example: Add a value to a cell.


## Methods Used: 
Worksheets().Activate
Range().Value
Range().NumberFormat
Range().Borders().LineStyle
Range().Font.Bold
Range().Font.Color
Range().Font.Italic
Range().Font.Size
Cells().End().Row
Cells().Value
Cells().ClearContents
Cells().Interior.Color
Cells.Clear
MsgBox ()

MsgBox()

### Code Stuctures Use
#### Loops
For Loops
  
#### Conditionals
If-Then, If-Then-Else and If-Then-Else-If structures.

#### Logical operators

### Arrays - AKA lists
Nested Loops (The Skill Drill was actually fun!)

Reinforcement of Best Practices
- Documentation: "Code is read more than it is written."
- Whitespace, give everyone's eyes a break.
- Code Reuse
- Planning and understanding the problem (AKA Requirements)
- Readability of Workbooks
- Removal of any Hardcoding in VBA code
- Code Performance
- Code Effeciency : Refactoring Code 
How fast something runs is very important in several industries. 
There isn't one needing their data faster and faster. 
It is an amazing race. With no further ado, I say, knowing how
to get my code to run faster than anyone else's is important to me. 
With the directions given, I did what I do.  


### Resources

Juniper: [Juniper Website] (https://www.juniper-design.com/)
Design Patterns: [The GoF Design Patterns Reference] http://w3sdesign.com/GoF_Design_Patterns_Reference0100.pdf
Code Smells and Anti-patterns: [Sniffing Out Success: Identifying Smells and Anti-Patterns in Your Code by Patrick Delancy's:] (https://patrickdelancy.com/2013/02/sniffing-out-success-identifying-smells-and-anti-patterns-in-your-code/)

Create a variable with single or long data types.
Writing for loops.
Writing if-then statements.
Using design patterns.
Using logical and comparison operators.
Using an index to access data in an array.
Using nested loops.
Reusing code.
Debugging and comment on code.
Using visual and numeric formatting.
Using conditional formatting.
Measuring code performance.

