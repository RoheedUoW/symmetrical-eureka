Question 1: 
Select Three Stocks from the S&P 500
Choose three individual stocks from the S&P 500 index.
Download their historical closing prices from 20 August 2024 to 20 February 2025.
Clearly state which stocks you selected and the source of the data (Yahoo Finance,
Bloomberg, etc.).

Answer
For this assignment, the three companies picked from the S&P 500 are:
1. Tesla (TSLA)- Stock Exchange (a member of the S&P 500)
2. Nvidia (NVDA) - NASDAQ (a member of the S&P 500)
3. Apple (AAPL) - NASDAQ (a component of the S&P 500)
The historical closing price data for these stocks, covering the period from August 20, 2024, to February 20, 2025, was obtained from Bloomberg Terminal. The data is shown on Microsoft Excel.


Question 2: 
Compute Key Statistical Measures
For each of the three stocks, compute:
• Daily Return
• Daily Volatility
• Compounded Annual Growth Rate (CAGR)*
• Annualized Volatility*

Answer
The daily return for each stock was calculated using the formula:
Daily Return = ln (Price today / Price yesterday) 
Log returns were used due to their desirable statistical properties, as discussed in Benninga and Mofkadi (2021).

Daily Return Calculation
•	Formula: Daily Return=Current Closing Price−Previous Closing PricePrevious Closing Price\text{Daily Return} = \frac{\text{Current Closing Price} - \text{Previous Closing Price}}{\text{Previous Closing Price}}Daily Return=Previous Closing PriceCurrent Closing Price−Previous Closing Price
•	Excel Formula (assuming Closing Prices are in column B and start from row 2):
excel
CopyEdit
= (B3 - B2) / B2
•	Drag this formula down to calculate the daily return for all rows.
________________________________________
Daily Volatility (Standard Deviation of Daily Returns)
•	Formula: σdaily=STDEV.S(Daily Returns)\sigma_{\text{daily}} = \text{STDEV.S}(\text{Daily Returns})σdaily=STDEV.S(Daily Returns)
•	Excel Formula (assuming Daily Returns are in column C from row 3 to row N):
excel
CopyEdit
= STDEV.S(C3:CN)
________________________________________
Compounded Annual Growth Rate (CAGR)
•	Formula: CAGR=(End PriceStart Price)1Years−1\text{CAGR} = \left( \frac{\text{End Price}}{\text{Start Price}} \right)^{\frac{1}{\text{Years}}} - 1CAGR=(Start PriceEnd Price)Years1−1
•	Excel Formula (assuming the start price is in B2, end price in B(N), and total years calculated in cell F2):
excel
CopyEdit
= (B[N] / B2) ^ (1 / F2) - 1
o	Total Years (F2) formula:
excel
CopyEdit
= (MAX(A:A) - MIN(A:A)) / 252
(Assumes column A contains dates)
________________________________________
Annualized Volatility
•	Formula: σannual=σdaily×252\sigma_{\text{annual}} = \sigma_{\text{daily}} \times \sqrt{252}σannual=σdaily×252
•	Excel Formula (assuming daily volatility is in cell G2):
excel
CopyEdit
= G2 * SQRT(252)
________________________________________
Excel Table Layout
A (Date)	B (Closing Price)	C (Daily Return)	D (Formula Applied)
2024-08-20	454.70		
2024-08-21	453.67	=(B3-B2)/B2	(453.67 - 454.70) / 454.70
2024-08-22	459.77	=(B4-B3)/B3	(459.77 - 453.67) / 453.67
...	...	...	...
Total Days	(MAX(A:A)-MIN(A:A))/252	Daily Volatility	=STDEV.S(C3:C100)
CAGR	=(B100/B2)^(1/F2)-1	Annualized Volatility	=G2*SQRT(252)






How to Use VBA:
1.	Open the file in Excel.
2.	Enable Macros (Excel will ask for permission).
3.	Press ALT + F8, select ComputeCAGR_AnnualVolatility, and click Run.
4.	The script will compute and fill in the CAGR and Annualized Volatility for all stocks.

VBA script:
Function CalculateCAGR(startPrice As Double, endPrice As Double, years As Double) As Double CalculateCAGR = (endPrice / startPrice) ^ (1 / years) - 1 End Function Function CalculateAnnualVolatility(dailyVolatility As Double) As Double CalculateAnnualVolatility = dailyVolatility * Application.WorksheetFunction.Sqrt(252) End Function Sub ComputeCAGR_AnnualVolatility() Dim ws As Worksheet Dim startPrice As Double, endPrice As Double, totalDays As Double, years As Double Dim dailyReturns As Range, dailyVolatility As Double Dim lastRow As Integer For Each ws In ThisWorkbook.Worksheets lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row startPrice = ws.Cells(2, 2).Value endPrice = ws.Cells(lastRow, 2).Value totalDays = ws.Cells(lastRow, 1).Value - ws.Cells(2, 1).Value years = totalDays / 252 ws.Cells(2, 5).Value = "CAGR" ws.Cells(3, 5).Value = (endPrice / startPrice) ^ (1 / years) - 1 Set dailyReturns = ws.Range(ws.Cells(3, 3), ws.Cells(lastRow, 3)) dailyVolatility = Application.WorksheetFunction.StDev_S(dailyReturns) ws.Cells(2, 6).Value = "Annualized Volatility" ws.Cells(3, 6).Value = dailyVolatility * Application.WorksheetFunction.Sqrt(252) Next ws MsgBox "CAGR and Annualized Volatility Calculated!", vbInformation, "Calculation Complete" End Sub


Question 3:
Compute the Covariance Matrix
Calculate the covariance matrix for the three stocks

Answer
Excel Formula to Compute Covariance
•	To compute the covariance between NVDA and AAPL:
excel
CopyEdit
=COVARIANCE.S(B:B, C:C)

Question 4:
Find the Global Minimum Variance Portfolio (GMVP)
Clearly show the optimal weights of the three stocks


Answer
Step 1: You can calculate the Global Minimum Variance Portfolio (GMVP) in Excel using the covariance matrix and matrix algebra.
Arrange your data with dates and daily returns for each stock in columns (e.g., NVDA, AAPL, TSLA) in an Excel sheet.
If you don't have returns, you can compute them using:
mathematica
CopyEdit
= (Current Price - Previous Price) / Previous Price
Apply this formula to all rows for each stock.

Step 2: Use Excel’s COVARIANCE.P() function to compute pairwise covariances:
•	Example for NVDA-AAPL covariance:
CopyEdit
=COVARIANCE.P(NVDA_Returns, AAPL_Returns)
•	Compute this for all stock pairs (NVDA-NVDA, NVDA-AAPL, NVDA-TSLA, etc.).
•	Arrange these values into a 3x3 covariance matrix.

Step 3: Compute GMVP Weights Using Matrix Algebra
1.	Create a 3×3 covariance matrix (Σ)
o	Use the covariance values you just calculated.

2.	Create a column vector of ones (1)
Example: If your covariance matrix is in range B2:D4, create a column vector {1,1,1} in another range.




3.	Compute the inverse of the covariance matrix (Σ⁻¹)
o	Use Excel’s MINVERSE() function:
makefile
CopyEdit
=MINVERSE(B2:D4)
o	This outputs the inverse covariance matrix.


Step 4: Multiply Σ⁻¹ by the ones vector
•	Use Excel’s MMULT() function:
php
CopyEdit
=MMULT(MINVERSE(B2:D4), F2:F4)
•	This gives an intermediate column vector

Step 5. Compute the scalar denominator (1' * Σ⁻¹ * 1)
•	Multiply the row vector {1,1,1} with the previous result:
php
CopyEdit
=MMULT(TRANSPOSE(F2:F4), MMULT(MINVERSE(B2:D4), F2:F4))
•	This returns a single scalar value.

6.	Compute the final GMVP weights
•	Divide each element of Σ⁻¹ * 1 by (1' * Σ⁻¹ * 1):
swift
CopyEdit
=G2/$G$5
•	Copy this formula for all three stocks.
The calculated GMVP weights should sum to 1 (or very close due to rounding). These represent the optimal allocation percentages for each stock in the Global Minimum Variance Portfolio.
Instead of manually computing the GMVP weights, you can use Solver:
Set an objective function to minimize portfolio variance:
mathematica
CopyEdit
Portfolio Variance = MMULT(TRANSPOSE(Weights), MMULT(Covariance_Matrix, Weights))
Set constraints:
Weights sum to 1.
No short-selling (if required, set weights >= 0).
Below is the summary:
Function	Purpose
COVARIANCE.P()	Computes covariance between two stocks
MINVERSE()	Computes the inverse of the covariance matrix
MMULT()	Performs matrix multiplication
TRANSPOSE()	Converts a column vector into a row vector
SUM()	Ensures weights sum to 1


Question 5:
Find a T-bill and Use It as the Risk-Free Rate
• Identify a current U.S. Treasury Bill (T-bill) rate to use as the risk-free rate.
• Provide the source for the risk-free rate (Federal Reserve, U.S. Treasury
website, Bloomberg, etc.)

Answer
March 5, 2025, the U.S. Department of the Treasury reported the following Treasury Bill (T-bill) rates:
T-bill Maturity	Bank Discount Rate	Coupon Equivalent Rate
4 weeks	4.59%	4.68%
8 weeks	4.64%	4.75%
13 weeks	4.76%	4.90%
26 weeks	4.87%	5.03%
52 weeks	4.98%	5.19%

The 13-week U.S. Treasury Bill (T-bill) is a popular choice for the risk-free rate because of its short-term nature and high liquidity. Based on the latest data, the current 13-week T-bill rate is around 4.90%.
These rates were obtained from the U.S. Department of the Treasury's official website (home.treasury.gov). However, T-bill rates change frequently, so it's always a good idea to check the latest numbers directly from the Treasury’s website or other reliable financial sources.

Question 6:
Find the Optimal Risky Portfolio
Select two of the three stocks and compute the tangency portfolio using the risk-
free rate.

Answer
 	Optimal Weights
NVDA	0
AAPL	1

How You Can Solve This in Excel
1.	Compute the mean daily return for each stock using =AVERAGE(range_of_daily_returns).
2.	Compute the covariance matrix using =COVARIANCE.P(range1, range2).
3.	Set up Solver:
Objective Function: Maximize Sharpe Ratio = (Portfolio Return - Risk-Free Rate) / Portfolio Volatility
Changing Cells: Portfolio Weights (w1, w2)
Constraints:
w1 + w2 = 1
0 <= w1, w2 <= 1
Use the GRG Nonlinear solving method



Question 7: 
Find the Optimal Complete Portfolio and Create a Graph
Compute the Optimal Complete Portfolio (OCP) by allocating funds between:
• The risk-free asset (T-bill).
• The Optimal Risky Portfolio (ORP) from Question 6


Answer
 	Optimal Complete Portfolio Weights
Risk-Free Asset	-1.43638
NVDA	0
AAPL	2.436385

Metric	Value
Expected Return	0.005472
Standard Deviation	0.041942
Sharpe Ratio	0.125827

 
Step 1: Computing the Proportion of Wealth Invested in the Risky Portfolio (y∗y^*y∗)
Formula:
= (Portfolio_Return - Risk_Free_Rate) / (Risk_Aversion * Portfolio_Variance)
Step 2: Compute the Risk-Free Asset Weight

= 1 - y_star 
(Cell for the risk-free weight
Check if y∗>1y^* > 1y∗>1 (indicating leverage is used))


Step 3: Compute the Expected Return of the OCP
= (Risk_Free_Weight * Risk_Free_Rate) + (y_star * Portfolio_Return)

Step 4: Compute the Risk (Standard Deviation) of the OCP
= y_star * Portfolio_Standard_Deviation

Step 5: Compute the Sharpe Ratio of the OCP
= (OCP_Return - Risk_Free_Rate) / OCP_Standard_Deviation

Step 6: Create the Capital Allocation Line (CAL) Graph (as shown above)
Excel Formula for a series of returns (for different risk levels):
excel
CopyEdit
= Risk_Free_Rate + Sharpe_Ratio * Risk

Step 7: Interpret the Results
Explain the investor’s choice based on risk aversion:
If y∗<1y^* < 1y∗<1 → The investor is conservative and holds more T-bills.
If y∗>1y^* > 1y∗>1 → The investor is aggressive and uses leverage to invest more in ORP.

Compare OCP vs. ORP Performance:
Portfolio	Expected Return	Risk (Std Dev)	Sharpe Ratio
ORP	E(R_P)	σ_P	S_P
OCP	E(R_C)	σ_C	S_C


The Optimal Complete Portfolio (OCP) offers a more balanced investment approach compared to the Optimal Risky Portfolio (ORP) by incorporating a mix of the risk-free asset (T-bill) and the ORP. While the ORP has a higher expected return, it also carries greater risk since it is fully invested in risky assets. 
In contrast, the OCP adjusts risk exposure based on the investor’s risk tolerance by allocating a portion to the risk-free asset, thereby reducing volatility while maintaining strong returns. The proportion of funds allocated to ORP, denoted as y∗y^*y∗, determines how much risk an investor takes—higher y∗y^*y∗ means more exposure to risky assets, while lower y∗y^*y∗ means a safer, more conservative portfolio.
 The Sharpe Ratio of the OCP is typically equal to or higher than that of the ORP, indicating an improved risk-adjusted return. This means that by properly diversifying between ORP and the risk-free asset, an investor can achieve a better return per unit of risk taken. Graphically, the OCP lies on the Capital Allocation Line (CAL), showing how different combinations of risk-free assets and the ORP provide the best possible return for a given level of risk. Ultimately, the OCP is more flexible and tailored to individual risk preferences, making it a more practical choice for investors compared to the fixed-risk ORP.


References:
Frontline Systems, Inc. (2021) GRG nonlinear solver - An introduction, Frontline Systems. Available at: https://www.solver.com/grg-nonlinear-solver
U.S. DEPARTMENT OF THE TREASURY (2019). Front page | U.S. Department of the Treasury. [online] Treasury.gov. Available at: https://home.treasury.gov/.
