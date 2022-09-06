# Stock Analysis
##Analyzing green energy stock to see how funds should be disbursed. 
### The purpose of this analysis is to assist Steve in determining if and how Steve's parent's money should be distributed across multiple green energy stocks rathan than just DAQO. I assisted Steve by analyizing the total number of shares traded daily, and how actively a stock is traded. By analyzing the data and looking at these variables, I can get a better understanding of which stocks are appropriate for Steve's parent's to invest in. 

##Analysis of Green Stock Data

###Analysis of DQ Stock for 2018
- Steve's parents wanted to know how actively DQ was traded in 2018. To find the total daily volume of the DQ stock, I had to create a for loop that would only increase the total colume for DQ's volume. First I had to create an inital value for the total volume to zero. Therefore, if the stock value was equal to DQ, then it would add to total volume. After running the analysis, DQ traded 107,873,900 shares in 2018. However, that number doesn't provide us with information in regards to how well DQ performed. Therefore, I also needed to look at the yearly return in 2018 for DQ to understand whether or not DQ performed well for 2018. 
- First, I needed DQ's first closing price and last closing price for the return analysis. I created conditionals in the VBA code to find these values. If the current row's ticker was DQ and if the previous row's ticker was not DQ, it provided our starting price. Whereas, if the current row's ticker was DQ and the next row's ticker was not DQ, then it provided our ending price. After we ran the analysis, DQ's return dropped about 63%. This indicates that DQ experienced a 63% financial loss in 2018. Therefore, investing in this green enegery stock would not be a financially effective choice for Steve's parent's. 
- After completing the analysis, I generated a button so that Steve could run the analysis and look at the data himself. This allowed Steve to run the analysis himself. In the script I also included a Timer function so that when Steve runs larger datasets in the future, he will know how fast the VBA code will compile the results. The first image shows the code time for the 2017 output. The second image shows the code time for the 2018 output. 

![Timer for 2017](https://github.com/mrma2318/stock-analysis/blob/16443aa5397beee98936b9b31e599fa763eaf287/Code%20Time%20for%202017.png)

![Timer for 2018](https://github.com/mrma2318/stock-analysis/blob/e664db94761f9cf23a6865878d49b12728574082/Code%20Time%20for%202018.png)

###Analysis of All Stock 
- Then Steve wanted us to expand our analysis to include the entire stock market from the last few years. Since this analysis could take a long to execute, I had to refactor the code to loop through all the data one time to collect the same information and make it more efficient. Refactoring the code will make it easier for Steve to read in the future as it decreases the amount of steps, uses less memory, and improves the logic of the code. 
- By using the script I created to analyize the DQ Analysis, I could refactor it so that I could loop through all the data at one time and collect all the information for the stocks. First, I needed to create a ticker index and set it equal for zero as my starting value for all the tickers. Then I needed to create three output arrays for the values I was analyzing: ticker volumes, ticker starting price, and ticker ending price. Similar to the DQ Analysis, I needed to increase the current ticker volume by using the ticker index. For the ticker starting and ending price, I used the same script as I did for my DQ Analysis, but adjusted the script to reflect the all stock analysis. 

##Results
