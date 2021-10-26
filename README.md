# stock-analysis
stock-analysis

## Project Overview

### This project was programmed in VBA. Each module's contents is described below.

## Subroutines/Macros 
```
### module1: DQAnalysis()
 Generate report containing a single stocks volume and return.

### module2: SkillDrillNo1() and SkillDrillNo2()
 1) Fill cells of a spreadsheet with values in relationship to the cell.
 2) Fill cells of a spreadsheet with results from an algorithm created to give the sum of cell's column and row values.

### module3: AllStocksAnalysis()
 Process multiple stocks to find the volume and return of each and report the results.

### module4: formatAllStocksAnalysis()
 Use VBA to format the "All Stocks Analysis" (reuse is possible)

### module5 ClearWorksheet()
 Subroutine created to clear any spreadsheet (reuse is possible)

### module6: AllStocksAnalysisRefactored()
 Refactoring completed with instructions given (adding arrays to store and report)

### module7: AllStocksAnalysisRefactorAgain()
 Refactoring based on my own understanding of how processing occurs 

### (As Sonic says,_"Faster, faster, faster, faster, faster.")

### *** Each sheet contains buttons to clear the worksheet and run the corresponding VBA code
```
## _KEY TAKE AWAYS_
### Within each module completing multiple stocks calculations, preformance analysis is completed.
### It was clear that by utilizing arrays and variables, processing time is improved drastically.
### And, nested for loops are expensive.


1) Finding a single stock's total daily volume and return
2) Finding the total daily volume and yearly return for multiple stocks
3) Creating performance analysis on AllStocksAnalysis() and refactoring 
   code so that it runs as fast as it can.


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

#Independant Study
#What makes code run faster?
##Multiple Cores?
https://www.newcmi.com/blog/how-many-cores#:~:text=When%20a%20computer%20multi-tasks%2C%20because%20a%20single-core%20processor,quicker%20transfer%20of%20data%20at%20any%20given%20time.
Points: 
1) Approximately 75% of CPU time is used waiting for memory access results.
2) Multiple cores allow PCs to run multiple processes at the same time with greater ease, increasing your performance when multitasking or under the demands of powerful apps and programs.
3) When a computer multi-tasks, because a single-core processor can manage one thread at a time, the system must move between the threads quickly to process the data.
4) A high clock speed means faster processor. For instance, a quad-core processor may support a clock speed of 3.0GHz, while a dual-core processor may hold a clock speed of 3.5 GHz for every processor. This means that a dual-core processor can run 14% faster. So, if you have a single-threaded program, the dual-core processor is indeed more efficient. On the flip side, if your program can use all 4 processors, then the quad-core will then be about 70% quicker than the dual-core processor.
5) When multiple cores work concurrently on instructions, at a lower rate than the single-core, they achieve an immeasurable processing rate. Multi-core processors produce high-performance computing (HPC). HPC will take complex computations and break them into smaller pieces. 
6) HPC can therefore enable users to manage difficult tasks at relatively lower energy, which is a significant factor in devices like laptops, mobile phones or laptop, which run on batteries. This kind of energy saving – and ultimately cost saving – is one way in which your business could benefit.
7) For database management, scientific analysis or anything that requires processing huge volumes of data at high speeds, HPC, enabled by multi-core processing, is also essential.

[4 Areas Where Multi-core Processing Really Matters]
https://blog.storagecraft.com/4-areas-multi-core-processing-really-matters/

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

