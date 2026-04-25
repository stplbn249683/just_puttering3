The main reason why I have a database of OHLC stock prices is to be able to do backtesting.

This program uses global optimization of backtested stock price data to try to find the optimized 
variables that give the best results for the strategy.  The advantage of global optimization is that it does 
not try to predict the stock price for very day; it just tries to find the trades that have a higher probability 
of making a profit.

Disclaimer: This software includes code that I developed for my own personal use and I have included it 
as a source of information. I am not recommending that anyone else use the code exactly as written. 
This software reflects my current understanding and may contain errors. In fact, the program is complex 
enough that there could be errors that I am not aware of. Since the code was developed for my own use, 
I did not attempt to make it elegant or efficient.  See the additional statement at the bottom of the file.

In this example of a program that does global optimization, I have deliberately chosen a few strategy 
examples that probably will not work all that well because I do not want to provide a strategy that I have 
actually tried to use.  So, this program would need to be modified to add different strategies to be useful.

Note that, unlike previous programs, the entire database connection string is specified in 
InitializeDefaults (or read in from the file connection_string.ini which is assumed to be located in the 
same folder as the executable).  This allows me to use different connection strings on different 
computers depending on whether the database is on the current computer or a different one.

This program would probably not be considered very well written because I have made extensive use of 
public variables that bypass the subroutine argument lists.  I did this in part to get the variables into the 
optimization routine (there is structure argument that could be used to do that but I did not see much to 
be gained by using it). Also, that allowed me to reduce some the changes that I needed to make when I 
added a new strategy.  I put most of the public variables in the structure named pb and the class named 
bp1 to make it easier to know where they are being used.

This program does not do walk-forward optimization.  Walk-forward optimization might help in providing 
clues about the robustness of the strategy but you would still need to pick which set of optimization 
parameters to use and you would probably pick the optimization parameters for the most recent time 
segment.  This program does not try to short the market since I have no interest in doing that. If fact, the 
results for SH (an inverse ETF) for various strategies show how hard it is to make a profit doing that over 
the long term which is not surprising since the stock market always goes up over the long term. The 
program does not add to existing positions.  It buys the whole position in the stock and sells the whole 
position.  The program includes fractional shares since I am just trying to figure out whether the strategy 
works and not to match exactly what I would do if I wanted to buy whole shares.

The program assumes that the stock is purchased at the end of the day and sold at the end of the day.  
Of course, in practice you cannot sell at the exact end of the day so I normally use the stock prices from 
20 or 30 minutes before the close as an approximation of the closing prices when deciding whether to 
buy or sell the stock.  That gives me time to figure out what I want to do.  So, I temporarily update the 
database 20 or 30 minutes before the market closes so I can run program strategy_buy and then 
overwrite the stock prices with the correct values after the close (and run strategy_buy again so that the 
long term results are accurate).  Strategy_buy is a separate program that uses the results from 
strategy_optimize to tell me when to buy, hold or sell the stock.

This program uses 3 Nuget packages that are not included in the project file: DocumentFormat.OpenXml, 
NLoptNet and Skender.Stock.Indicators.  The DocumentFormat.OpenXml package is used to read in the 
input xlsx file in subroutine ReadExcelFileSAX. Note that this does not work correctly if the Excel input 
file contains more than one worksheet.  Since I only intended that subroutine to read a very simple Excel 
file with one worksheet, I did not take the time to investigate why.  I have a slight preference for using an 
xlsx file for input instead of a csv file because it remembers the column widths.

The NLoptNet package is used for the global optimization.  I did not spend much time looking at the 
various optimization methods in NLoptNet.  Several of them exited quickly with terrible solutions. 
However, The ESCH algorithm worked well, maybe in part because it keeps on trying and does not give 
up.  However, it gives a slightly different answer each time it is run so, since the execution time does not 
need to be extremely quick because the optimization only needs to be done once, I set up the program 
to try several attempts and choose the attempt that had the best result. The optimization returns 
decimal values and not discrete integers.  So, if the optimization parameter needed to be a discrete 
integer, I rounded the returned value to the nearest integer.  I thought that this might affect the 
convergence but it does not seem to.  I had to rename the 64-bit version of nlopt.dll to nlopt_x64.dll and 
copy it to the folder containing the executable to use it.

The displayed form has inputs for  Number of Attempts  and  Max Solver Iterations .  The  Number of 
Attempts  is the number of times that the optimization is repeated to try to find the best result. The Max 
Solver Iterations  is an input into NLoptSolver. If the optimization parameters are only used as multipliers 
and comparison limits and the indicators just use standard parameters, then the indicators can be 
computed before the loop for the global optimization and 10 attempts with 90000 max solver iterations 
can be done in a reasonable amount of time. If the optimization parameters are input parameters for the 
indicators, then the indicators need to be computed for each loop of the global optimization and even 3 
attempts with 50000 max solver iterations will take a fair amount of time.

There is a checkbox for including the interest that is earned on the money when it is not invested in a 
stock.  Including the interest earned probably results in a fairer comparison to a buy and hold but it 
makes it more difficult to compare to the results from other back testing programs.  The interest rate 
percentage is also used in the Sharpe ratio calculation.

The goal of the optimization is likely to be to find the maximum percent gain per day (while owning the 
stock), the maximum Sharpe ratio, the maximum win rate, the maximum gain/loss ratio, etc.  The 
program calculates the Sharpe ratio based on the days that the stock or ETF is owned so it is basically the 
same as maximizing the percent gain per day.  Optimizing based on percent gain per day seems to me to 
maybe be a little more stable but I have really spent hardly any time comparing the two. 

However, the maximum gain per day or the maximum Sharpe ratio is likely to occur on one (or very few) 
anomalous trades. So, some other condition such as a minimum number of trades per year is needed to 
ensure that the results are based on enough trades to be useful.  But I was reluctant to force the 
optimization to find a specified number of trades if the optimization found enough trades without it.  So, 
I had the program start with a small weight for the minimum number of trades per year and increase it 
on every iteration until enough trades are found (or it decided to stop trying). This can significantly 
increase the time it takes to find a solution.  The number of iterations from increasing the weight for the 
minimum number of trades is displayed in the box labelled Repeat.  This is probably somewhat quirky 
but since I wrote the program for my own use, I can do as many quirky things as I want.

The program allows each stock or ETF to be analyzed individually or from a file containing the ticker 
symbol, the category, and the number of days in the time segment (ending at the most recent day) for 
each stock or ETF.  The reason for the category is that some strategies work better on ETFs than on 
stocks.  The number of days should allow for enough additional days in the database before the time 
segment for errors in the calculations to die out.  The input file should be an xlsx file with 3 columns.  I 
have included s1_input.csv as an example where I have provided a csv file instead of an xlsx file to avoid 
concerns about viruses in the xlsx file.   The output files s<strategy number>_results.csv and s<strategy 
number>_transactions.csv are only created when the stock or ETF is analyzed individually (not from an 
input file).

The results are stored in the output file s<strategy number>_summary.csv.  The optimization parameters 
are stored as columns x(0), x(1), x(2) etc.

I put the parts of the program that need to be modified when a new strategy is added in 
 IndivdualStrategies.vb . To add a new strategy, the names of the lists containing any new indicators (or 
quantities that can derived from the indicators such as slopes) need to be added to public class 
PublicAccessType1.  Function number_of_opt_variables sets the number of optimization parameters 
that are used.  Subroutine set_bounds sets the upper and lower bounds for the optimization parameters 
and the initial values. Note that for the SuperTrend indicator, I used pb.x_type(0) = "i" to tell the 
program that the first optimization parameter needs to the rounded to an integer when it is written to 
the summary file.

Function calculate_indicators sets the minimum number of trades per year and the weights for the 
various constraints, and calculates the indicators that can be calculated before the optimization loop 
starts.  The weights are used to add or subtract additional terms to the sum to be maximized.  If the 
minimum number of trades per year is set to zero, then the minimum number of trades is not used as a 
criterion. The win rate can be increased some by increasing the weight but that will probably cause 
something else such as the gain/loss ratio to decrease.

If all the indicators (and derived quantities) can be calculated before the optimization loop starts, then 
ResizeLists can be used to resize all the lists to the shortest length.  This ensures that all the lists cover 
the same time interval.  If the time of the first point in a list is different from the other lists, then the 
calculations will be erroneous.

Function calculate_indicators1 sets the minimum number of trades per year and the weights for the 
various constraints, and calculates the indicators that need to be calculated inside the optimization loop.  
It may not be desirable to keep resizing the lists repeatedly for each optimization loop. So, function 
MinNumPoints can be used to calculate to minimum number of points from all the lists.  This makes the 
code for the strategy in find_action a little more complicated.

Function results_file_heading can be used to add column headings to the file s<strategy 
number>_results.csv that are specific to an individual strategy. Function results_file_text can be used to 
add lists to the file s<strategy number>_results.csv that are specific to an individual strategy.

Function find_action contains the individual strategies.  The code for the SuperTrend indicator is a little 
more complicated because the parameters for the indicator need to be calculated inside the 
optimization loop.  Consequently, the lists have different lengths and an offset based on the length of the 
smallest list is used to find the starting point in each list.

I reduced the number of strategies to 3.  I have tried far more but I don't want to include the strategies 
that I have tried myself except for 3 examples that I would not expect to work all that well.

Strategy 1 just uses the MACD histogram, RSI and MFI as can be seen in find_action.  I normalized the 
MACD histogram by dividing by EMA(26) in calculate_indicators so that the MACD histogram 
optimization parameters for different stocks and ETFs would be of similar magnitude.  The optimization 
parameters are x(0), x(1), etc.  The optimization parameters are just used as multipliers and comparison 
values and the indicators use standard parameters so the indicators can be calculated before the 
optimization loop starts.

Strategy 2 just uses the VIX.  It assumes that the VIX is stored in the database as $VIX, the symbol used 
by Schwab.

Strategy 3 uses the SuperTrend indicator.  It needs to be calculated in calculate_indicators1 instead of 
calculate_indicators because it needs to be calculated inside the optimization loop.  It may not be 
desirable to keep resizing the lists over and over again inside the optimization loop, so I used 
MinNumPoints to calculate to minimum number of points from all the lists.  Then I calculated the offsets 
inside find_action that would cause the times for all of the lists to line up.  This made the code slightly 
more complicated than when ResizeList was used.  Notice that this also caused the code inside 
results_file_text to be slightly different than it would have been if ResizeList had been used.  Note that, 
as indicated earlier, since this strategy is inside the optimization loop, the number of attempts and the 
number of max solver iterations will need to be reduced in order to get an answer in a reasonable 
amount of time.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


