# stock-analysis

### Overview
The goal for this project was to help Steve make recommendations to his parents for purchasing stock in green energy companies. The data, presumably, came from public exchanges. Steve's parents are keen to invest solely in DAQO New Energy Corporation (ticker DQ), a company that makes silicon wafers for solar panels. Steve believes that they would be in a better position by diversifying, and wants a throughough analysis of several green energy companies to prove the point. Thus, my tasks were to review trading data from the passed few years to learn how the market value each of the given companies' stock performed, and to provide a concise report with the findings.

#### Methodology
The first part of my analysis consisted of writing a subroutine in Visual Basic for Applications (VBA) to analyze DAQO's stock performance in 2017 and 2018. Once the subroutine produced correct and accurate results, I could then extrapolate and refactor the code in order to analyze each stock in the data set. After verifying the results, I wrote additional subroutines to format the data into a table that was easy to use and read. Directly in Microsoft Excel, I created the buttons that would dispaly results with a single click of the mouse.

#### Results
In 2017, DQ had outstanding performance, increasing in value from the first trade of the year to the last by 199.45%. This proved to be the best-performing stock in the data set. I suspect that this performance is the reason Steve's parents heard of the company and wanted to invest in it.

![Analysis results for 2017](https://github.com/veachk90/stock-analysis/blob/main/Screenshot%20(8).png)

Unfortunately, the data from 2018 tells a different story. The market value of DQ's stock fell 62.6% from the beginning to the end of year, the worst performance of any company in the data set. In fact, for that year, only two green energy companies of the twelve in our set showed favorable results: ENPH and RUN, with roughly 82% and 84% increases in market value, respectively. 

![Analysis results for 2018](https://github.com/veachk90/stock-analysis/blob/main/Screenshot%20(10).png)

Thus, unless Steve's parents had decided to invest prior to DQ's tremendous increase in value in 2017, they would do much better by diversifying. This would allow their portfolios to benefit from some companies that may be doing well during market downturns and losing excessive value. Of course, the trade-off is that large increases in value of a single company are mitigated by having dedicated funds to the stocks of other companies.  

#### Analysis
This analysis was conducted in VBA by first creating an array of each of the tickers. One loop was created to increment through each element of this array. This ensured that the analysis compared data from each company for the given years. Then, another loop was nested in the first that determined the opening and closing prices for each company for each year, as well as tracking the total daily trading volume. 

![Nested For loops to extract necessary data](https://github.com/veachk90/stock-analysis/blob/main/Screenshot%20(11).png)

Having scoured each entry in the data for these significant figures, I could then consider a ratio of each stock's starting and ending market price for the given year and express it as a percentage. Finally, putting each company's percentage in a simple table made comparing their stock performances straightforward. A person who invested in DQ at the end of 2017 would have been sorely disappointed in 2018!

![Formatting the data for presentation in Excel](https://github.com/veachk90/stock-analysis/blob/main/Screenshot%20(12).png)

#### Take-aways
The process of breaking the problem in simpler, smaller problems proved to be an effective strategy. Figuring out how to analyze a single company's performance first, and then extrapolating the methodology to consider a set of multiple companies, prevented a daunting task from becoming a giant, confusing, and frustrating coding process. Another benefit of having completed this project is that I now have this tool written and available. With some simple re-working, it could be used to provide a similar analysis of any set of tickers. During this process, I had to search for error messages on internet search engines in order to identify problems. Understanding the problems as they arose allowed me to consider possible solutions, rather than feeling stuck and frustrated. 

#### Summary
To summarize, Steve's parents are most likely better to diversify their available funds across multiple green energy companies, rather than choosing a single company for investment. This strategy will protect their assets from rapid market downturns. 

#### Additional notes
Refactoring code offers many advantages. In a large project, organization is paramount to success. If multiple programmers will be working on or with the same program, it helps everyone involved to keep each section as simple and concise as possible (while still meeting the objectives, of course). It is also helpful for efficiency. One example would be including unnecessary calculations in a For loop over a large amount of data, which poses many extra steps for the computer as it processes the code. Refactoring would allow these calculations to be performed after all of the looping is completed and the necessary data is stored in variables. However, if not conducted correctly, refactoring could introduce new bugs or errors that then have to be solved. While it can offer an improvement in efficiency, mismanagement of the source code could possibly result in introducing addional steps that actually take more time or resources. 

In this particular example, the differences in efficiency were rather small. This is primarily due to the size of the dat set. In cases with larger sets, efficiency becomes a higher priority. However, with the refactored code, all of the tabulating and formatting in Excel was done right away when the program was executed, so the user would not have to push multiple buttons in Excel to produce the final result. Considering the experience for the end user is almost always a high priority, if not the top one, so I rather liked this outcome.
