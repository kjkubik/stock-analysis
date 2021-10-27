# stock-analysis
stock-analysis

## Project Overview

### This project was programmed in VBA. Each module's contents is described below.

## Subroutines/Macros 

```
module1: DQAnalysis()
 DQAnalysis() processes a single stock's data to create a report containing total 
 volume and year's return. The report is generated in an Excel spreadsheet using VBA.

module2: SkillDrillNo1() and SkillDrillNo2()
 Skill Drills are more-less what Dale Carnegy would call "sharpening the saw". 
 
 For the first skill drill, we dive into filling in cells systematically so that the 
 sum of the row number and column number is displayed for each cell in a 10x10 range. 
  
 [put a picture of this here!] 
 
 1: Practice working with nested forloops to generate values in Fill cells of a spreadsheet with values in relationship to the cell.
 2) Fill cells of a spreadsheet with results from an algorithm created to give 
    the sum of cell's column and row values.

module3: AllStocksAnalysis()
 Process multiple stocks to find the volume and return of each and report the results.

module4: formatAllStocksAnalysis()
 Use VBA to format the "All Stocks Analysis" (reuse is possible).

module5 ClearWorksheet()
 Create subroutine to clear any spreadsheet (reuse is possible).

module6: AllStocksAnalysisRefactored()
 Refactor by adding arrays to store and report volume and returns. 

module7: AllStocksAnalysisRefactorAgain()
 Refactor based on understanding of processing and how performance can 
 be improved by 'tweeking' the code just a little. 
 (As Sonic says,_"Faster, faster, faster, faster, faster.")
```

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
- Utilizing buttons with in a worksheet to execute subroutines.

## Methods Used: 
- Worksheets().Activate
- Range().Value
- Range().NumberFormat
- Range().Borders().LineStyle
- Range().Font.Bold
- Range().Font.Color
- Range().Font.Italic
- Range().Font.Size
- Cells().End().Row
- Cells().Value
- Cells().ClearContents
- Cells().Interior.Color
- Cells.Clear
- MsgBox ()
-?????????????????????????????????????????

## Results and _KEY TAKE AWAYS_

### Within each multiple stock analysis, preformance analysis is completed. It was clear 
### that by utilizing arrays and variables, processing time is improved drastically.
### And, nested for loops are expensive.

### I seen this as an introduction to bigger and better things. The editor within Excel is lame. 
### It pained me greatly. Instead of using it, use another, more powerful editor. I was ashamed 
### to use it. 

### you have to say something about the differences between running 2017 and 2018!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

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

