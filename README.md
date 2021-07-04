# VBA Challenge/stock-analysis
Module 2 - VBA

## Overview of Project
Analyze stock data for two different years to show annual volume and return.
### Purpose
Steve has requested our help with this data to analyze annual volumes and returns for multiple stocks.  Using VBA, we will automate analyses for Steve to use to assist his clients.  After automating the analysis first, we will then refractor the VBA code to make it run faster, to allow Steve to use it for thousands of stocks.
### Background
Steve has just graduated with his finance degree and will be using our help to analyze stocks for his first clients, his parents.  They are passionate about green energy, particularly DAQO New Energy Corp ($DQ).  Steve has some concerns about them needing to diversity their portfolio and will be using this data to assist in that decision.
![Steve_Parents_2_85](https://user-images.githubusercontent.com/74840026/124373378-c07a6400-dc46-11eb-9512-1293b1dfc6a1.png)

## Analysis and Challenges
## Results
Refactoring the code using `for` loops with multiple arrays resulted in a faster run time for the subroutine.
### Original results
Using VBA to analyze the data, our first attempt took roughly 0.75 seconds to operate.  With this method, we used nested `for` loops to run through the data multiple times to collect the volumes and prices.  Although this method worked correctly, using it to analyze thousands of stocks will take significant time.

![2017 before](https://user-images.githubusercontent.com/74840026/124373481-68902d00-dc47-11eb-96ef-c8ed4d28a43a.PNG)
![2018 before](https://user-images.githubusercontent.com/74840026/124373482-6af28700-dc47-11eb-9bcd-59be0bc733ae.PNG)

### Refactored results
Once refactored, our code took approximately 0.12 seconds to operate, giving us our results in less time.  With the refactored code, we used a `for` loop to go through the data once, collecting the volumes and prices for each stock and storing them into arrays.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/74840026/124373500-91182700-dc47-11eb-8c88-d9407de0fcb2.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/74840026/124373501-92e1ea80-dc47-11eb-840a-aabc6d050023.PNG)

### Challenges and Difficulties Encountered
Not having a good basic understanding of arrays in VBA, I had quite some challenge getting them to work correctly.  I struggled to understand exactly what each step was asking for.  After starting over multiple times, I wrote down each step by hand to work my way through and did multiple Google searches to gain a better understanding of the proper use of arrays in VBA.  Taking advantage of Office Hours and help from classmates through Slack, I was able to write them to perform as desired.

- [Excel-Easy.com](https://www.excel-easy.com/vba/array.html)
- [Using arrays-Microsoft](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays)
- [ExcelMacroMastery.com](https://excelmacromastery.com/excel-vba-array/)

## Summary
Refactoring code has both advantages and disadvantages in this example.  
- Advantages
  - Using more complex methods, a faster run time was gained, giving the user the ability to run this subroutine for thousands of lines of stock data quickly.

- Disadvantages
  - A thorough understanding of how the original code was written is needed to find where the code can be changed to produce faster results.
 
 - Pros/Cons in refactoring original VBA script
    - Working through the instructions to create the original program, I gained a good understanding of `for` loops and nested `for` loops.
    - Cluttered code with minimal notes will pose difficulties to refactoring.



