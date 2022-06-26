# Stock_Analysis

### Overview of Project
   Steve loves the workbook that was prepared for him. At the click of a button, he can analyze an entire dataset and clear it to restart. Now, to do a little more          research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although the code works well for a dozen stocks,it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.  There needs to be a refactor of the original code to enhance the speed at which the data is run for more efficiency.

### Results

 After running the refactored script, we see that it was faster for both years being run below.  Updated script included formatting to show were gains/losses were made for Steve to quickly see what needs to be addressed from an investing standpoint.  Additionally, the ease of using a button to reset and clear data and change years makes this a quick tool to view different years at different times. 

Runtime for 2017

   ![2017%20Macro%20Runtime_GJ.png](Resources/2017%20Macro%20Runtime_GJ.png)

Data for 2017

   ![2017%20Data%20_GJ.png](Resources/2017%20Data%20_GJ.png)

Runtime for 2018

   ![2018%20Macro%20Runtime_GJ.png](Resources/2018%20Macro%20Runtime_GJ.png)

Data for 2018

   ![2018%20Data%20_GJ.png](Resources/2018%20Data%20_GJ.png)

### Summary

  Advantages / Disadvantages of refactoring code in general
  
  1 - Refactoring gives you a place to start and apply to your data set or problem / This can cause and issue because you have to make changes in potential variables         and have to change it in numerous places
  
  2 - Refactoring can help expediate a project and meet deadlines / Your code is not unique and can cause more debugging since it's being applied to a new data set 
  
  Advantages / Disadvantages of refactoring code for this VBA script
  
  1 - Original script was a quick script to loop through all the data as it was originally just the 11 stocks without having to worry about more years being added /       Original script was limited by not having a way to input the year you wanted to pull and having to pull it year by year.
  
  2 - Refactoring this made this specific process faster for the client and he can add tabs in the same format to the workbook and select years as needed / This VBA         script would stil need to be updated if more stocks being tracked are added there by having to update the tickerIndex.
