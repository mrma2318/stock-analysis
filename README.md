# Stock Analysis
## Analyzing green energy stock to see how funds should be disbursed. 
### The purpose of this analysis is to assist Steve in determining if and how Steve's parent's money should be distributed across multiple green energy stocks rathan than just DAQO. I assisted Steve by analyizing the total number of shares traded daily, and how actively a stock is traded. By analyzing the data and looking at these variables, I can get a better understanding of which stocks are appropriate for Steve's parent's to invest in. 

## Analysis of Green Stock Data

### Analysis of DQ Stock for 2018
- Steve's parents wanted to know how actively DQ was traded in 2018. To find the total daily volume of the DQ stock, I had to create a for loop that would only increase the total colume for DQ's volume. First I had to create an inital value for the total volume to zero. Therefore, if the stock value was equal to DQ, then it would add to total volume. After running the analysis, DQ traded 107,873,900 shares in 2018. However, that number doesn't provide us with information in regards to how well DQ performed. Therefore, I also needed to look at the yearly return in 2018 for DQ to understand whether or not DQ performed well for 2018. 
- First, I needed DQ's first closing price and last closing price for the return analysis. I created conditionals in the VBA code to find these values. If the current row's ticker was DQ and if the previous row's ticker was not DQ, it provided our starting price. Whereas, if the current row's ticker was DQ and the next row's ticker was not DQ, then it provided our ending price (see Image 1 for how the code was created for this). After we ran the analysis, DQ's return dropped about 63%. This indicates that DQ experienced a 63% financial loss in 2018. Therefore, investing in this green enegery stock would not be a financially effective choice for Steve's parent's. 

### Image 1: 

![DQ Analysis Code](https://github.com/mrma2318/stock-analysis/blob/b2e2009c43a6defa8ee5e97813c064a5b90ef4dd/Screen%20Shot%202022-09-07%20at%202.32.37%20PM.png)


- After completing the analysis, I generated a button so that Steve could run the analysis and look at the data himself. This allowed Steve to run the analysis himself. In the script I also included a Timer function so that when Steve runs larger datasets in the future, he will know how fast the VBA code will compile the results. The first image shows the code time for the 2017 output. The second image shows the code time for the 2018 output. 

### Image 2:

![Timer for 2017](https://github.com/mrma2318/stock-analysis/blob/16443aa5397beee98936b9b31e599fa763eaf287/Code%20Time%20for%202017.png)

### Image 3:

![Timer for 2018](https://github.com/mrma2318/stock-analysis/blob/e664db94761f9cf23a6865878d49b12728574082/Code%20Time%20for%202018.png)


### Analysis of All Stock 
- Then Steve wanted us to expand our analysis to include the entire stock market from the last few years. Since this analysis could take a long to execute, I had to refactor the code to loop through all the data one time to collect the same information and make it more efficient. Refactoring the code will make it easier for Steve to read in the future as it decreases the amount of steps, uses less memory, and improves the logic of the code. 
- By using the script I created to analyize the DQ Analysis, I could refactor it so that I could loop through all the data at one time and collect all the information for the stocks. First, I needed to create a ticker index and set it equal for zero as my starting value for all the tickers. Then I needed to create three output arrays for the values I was analyzing: ticker volumes, ticker starting price, and ticker ending price. Similar to the DQ Analysis, I needed to increase the current ticker volume by using the ticker index. For the ticker starting and ending price, I used the same script as I did for my DQ Analysis, but adjusted the script to reflect the all stock analysis. 
- Once the script and code had been refactored, I then had to adjust the button I created previously to run the new refactored code for all the stocks. This way, when Steve clicks the "Run Analysis for All" button, it would give him the total daily volume and return for all the stocks, and he could choose the year he wants to review. By refactoring the code, it also provided a faster code run time. It improved the code run time for 2017 from 0.61 seconds to 0.08 seconds. For 2018, it improved the code run time from 0.55 seconds to 0.08 seconds. The new refactored code run times are provided in images 5 and 6. 

### Image 4: 

![Refactored Analysis Code](https://github.com/mrma2318/stock-analysis/blob/1aef69c928d18589838845ccfc282237fcfca14f/Screen%20Shot%202022-09-07%20at%202.37.34%20PM.png)

### Image 5:

![Refactored Timer for 2017](https://github.com/mrma2318/stock-analysis/blob/f75c2b77b8bc7e98b83d7e786bbae3b8ed3ebec0/VBA_Challenge_2017.png)

### Image 6:

![Refactored Timer for 2018](https://github.com/mrma2318/stock-analysis/blob/f75c2b77b8bc7e98b83d7e786bbae3b8ed3ebec0/VBA_Challenge_2018.png)


## Results
- There are two conclusions that we can draw by looking at the analysis. First, the original stock that Steve's parents want to invest in was DQ. In 2017 the stock had a financial gain of 199%, however, in 2018 DQ had a 63% financial loss. While the financial gain was more significant than the financial loss they experienced, the DQ stock is probably not the most reliable for Steve's parents to invest their money in. 
- However, there are two stocks that both had financial gain in 2017 and 2018, that would be greater for Steve's parents to invest their money. The first stock is ENPH, in 2017 the stock had a financial gain of 130%. Then in 2018, ENPH stock had another financial gain of 82%, showing that the ENPH stock continues to do well over the years. The other stock doing financially well is the RUN stock. In 2017 the stock had a finacial gain of 6%, and then in 2018 the stock had a financial gain of 84%, also showing that RUN continues to do well over the years. Recommending ENPH and RUN for Steve's parents, would be a better option than investing in DQ, since ENPH and RUN are consistent with returning financial gain every year. 
- There are some limiations to this dataset though. The first one is that we only have stock data from 2017 and 2018. To have a better understanding on how well a stock is performing, it would be beneficial to have data going back even further than 2017. A stock could do well two years in a row, but that doesn't mean they haven't experienced high fincancial loss at some time either. Even though DQ experienced a 63% financial loss in 2018, they might have done well the last few years, but we can't see data beyond 2017. 
- Another limiation is that we don't understand why the stocks that experienced a financial loss, had that loss. So understanding the external factors that impacted those stocks. Just because a stock had a bad year, doesn't mean it's because the company or stock didn't do well, there could be external factors that impacted that financial situation. Understanding that information as well, could assist Steve in helping his parents make smart financial decisions on where to dispurse their money. 

##Conclusion
- There are some advantages and disadvantages to refactoring code. Some advantages to refactoring code is that it provides better comprehensibility or understanding of the code. This way if other programmers review the code, it's easier for them to understand what you did. Refactoring also removes redundancies or duplicate scripts by making the code more combined into one overall code instead of multiple individual codes. For example, instead of having multiple for loops, we can have a nested for loop with output arrays for a more comprehensive script. This also allows us to have cleaner code and better testability. 
- Some disadvantages to refactoring code is that it could introduce more errors into the code. Another disadvantage is that the customer, so in our analysis our customer was Steve, they might not recognize the difference between the original code and a refactored code since the functionality and results stay the same. 
- There are also some disadvantages between the original script and the refactored VBA script. 

