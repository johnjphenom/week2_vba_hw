# week2_vba_hw
Week 2 VBA Homework for Boot Camp
# Kickstarting with Excel

## Overview of the Project

### Purpose

The goal of this project is to assist Steve with an analysis of stocks that he is going to provide to his parents. Initially his parents were looking at investing in a green energy company Daqo’s (DQ). Our goal was to write code in VBA to analyze this for the year 2018 and after doing so successfully it turned out DQ would be an unwise investment based on past performance. From here Steve wanted to the analysis to the performance of a dozen stocks to see will be better suited for investment. The differing stocks were analyzed for both 2017 and 2018. From this analysis there still did not seem to be as many well performing stocks as they would like to make a well diversified portfolio. 
The goal of the challenge is to refactor the code so that the entirety of the stock market can be efficiently and effectively analyzed. The current code was adequate for the twelve stocks but the repetitive nature of the program in its current form would necessitate copying and pasting repetitive code. This would make the program much slower than a refactored optimized program could. The final program was such that it looped through all the data one time instead of the inefficient manner of the original code. This will allow the better and more timely analysis or the stock market.

## Results

After the challenge was completed and the refactored code was implemented it was time to test the functionality and speed of the program. At first glance one could easily surmise that the program ran much faster without even setting up the message box to show concrete results. Yet, we want to have quantitative facts to show this to be the case. The below image shows the difference in computing speeds between the original program and the refactored code for both the 2017 analysis and the 2018 analysis.

![This is and image](https://github.com/johnjphenom/week2_vba_hw/blob/main/Resources/Original%20Code/Original_Code_2017_Time_Trial.png)
Time Trial Results for the Original Code for 2017


![This is and image](https://github.com/johnjphenom/week2_vba_hw/blob/main/Resources/VBA_Challenge_2017.png)
Time Trial Results for the Refactored Code for 2017


![This is and image](https://github.com/johnjphenom/week2_vba_hw/blob/main/Resources/Original%20Code/Original_Code_2018_Time_Trial.png)
Time Trial Results for the Original Code for 2018


![This is and image](https://github.com/johnjphenom/week2_vba_hw/blob/main/Resources/VBA_Challenge_2018.png)
Time Trial Results for the Refactored Code for 2018


When looking at the results it shows that the refactored code was about a whole second faster than the original code. Another way to look at it is that based on the ratio comparing the times, the refactored code is about five times more efficient. This clearly shows definitively that the refactored code had the effect of improving the efficiency of the original program. By not having to loop through the program an unnecessary amount of times with repetitive code and instead using arrays to streamline this, the program was much improved. The final product is shown to be up to the standards to allow Steve to assist his parents in their financial investments well.  

## Advantages and Disadvantages
