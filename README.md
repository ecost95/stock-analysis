# stock-analysis

## Overview of Project

The purpose of this project is to create a refactored VBA script in order to analyze stocks datasets. The script will loop through either the 2017 or 2018 (based on the user input), then will generate an summary of the total daily volumn and the returns for each stock ticker. This table will be places on the sheet "All Stocks Analysis", and the return column will be color coordinated. Finally, the script will tell us how fast our code ran for the chosen year. 


## Results 

- The VBA refactored script ran very fast on our 2017 and 2018 worksheets. The 2017 summary generated at 0.44 seconds and the 2018 summary generated at 0.49 seconds. 

![Refactored 2017](https://raw.githubusercontent.com/ecost95/stock-analysis/main/VBA_Challenge_2017.PNG)


![Refactored 2018](https://raw.githubusercontent.com/ecost95/stock-analysis/main/VBA_Challenge_2018.PNG)

- If you examine the original script in the "greenstocks (1).xlsm" file, the 2017 script ran at 0.55 seconds. This means our refactored code slightly improved upon our original.

![Original 2017](https://raw.githubusercontent.com/ecost95/stock-analysis/main/2017Original.png)


- In our refacored code, we had two seperate for loops - one that initialzed ticker volume to 0, and one that increased the ticker volume and set the starting/ending prices. 

![Refactored code](https://raw.githubusercontent.com/ecost95/stock-analysis/main/Codesnippet.PNG)

- This is different from the original code, where we nested the for loops and added the values to the initial table, and created a new subroutine for the formatting. These changes seemed to have affected our run times.
  
  
## Summary: 

- What are the advantages or disadvantages of refactoring code?
* Generally, the advantage of refactoring code is that it will make our scripts more efficient and easier to use. Howver, there is a risk that we will break the script or add unnecessary edits.

- How do these pros and cons apply to refactoring the original VBA script?
* The advantage for our refactored code is that that the code runs faster. However, because the code is in one for loop, there is a risk that a small error could break the entire script and would be difficult to fix.
    
