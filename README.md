# Stocks Analysis
## Project Overview
The client, Steve, requested that we refactor the stocks analysis workbook prepared for him. To ensure that the workbook is robust and efficient independent of the dataset size, the workbook was refactored to look for speed and code efficiency improvements.

## Results
### 2017 and 2018 Comparison
As shown in the two images below, it is clear that overall the stocks analyzed collectively performed better in 2017 than 2018. Eleven out of twelve stocks had positive returns in 2017, compared to two in 2018.

![2017 results](/Resources/VBA_Challenge_results_2017.png)

![2018 results](/Resources/VBA_Challenge_results_2018.png)

### Code Refactor
Initially, the stocks analysis code executed in 0.7 seconds as shown in the figure below.

![Original results](/Resources/InitialRun_2018.jpg)

After refactoring, the 2017 and 2018 run times improved, as shown in the two image below.

![2017 refactor results](/Resources/VBA_Challenge_2017.png)

![2018 refactor results](/Resources/VBA_Challenge_2018.png)

The primary refactor modification was to implement arrays for the output variables, looping through and outputting as arrays. This is different than the output methodology shown below. 
![Code results](/Resources/VBA_Challenge_outputCode.png)

## Summary
The advantages to refactoring code are that it can improve run time efficiency, aid in organization and identify potential bugs. Refactoring the original VBA script, the author found one potential bug where the code comments suggested iterating the ticker index in two separate spots. That would have led to over-indexing the output arrays and ultimately killing the code.
