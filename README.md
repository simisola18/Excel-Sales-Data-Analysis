# Excel-Sales-Data-Analysis
Data Cleaning and Validation, Calculations and Analysis, Sales Performance Evaluation. I developed a project in Excel, creating pivot tables to analyze the data. This process involved several stages, including data preprocessing, cleaning, and visualization.
## Project Objective
The dataset from Kaggle needed the dataset to be cleaned and certain cleaning and validation steps taken, calculations and analysis on the sales dataset done and a sales performance evaluation done also. The tasks needed to be done and questions to be answered are already uploaded as a file on this repository and can be reviewed to understand all that was done on the dataset and why and how I got to the answers below.
## Dataset used
Raw Sales_Data_Expanded.xlsx
Results Sales_Data_SimiOduba.xlsx

## Section 1: Data Cleaning and Validation
1. Identify and Correct Errors:
A. The date column had 12 rows with missing dates, and 1 of them was a date that had an error put down as the 31st of February, which doesn’t exist. I Formatted the cells to align with the other filled date cells based on the custom rule used and inputted this formula to find the average between the cell above and below to get the answers =IF(AND(B8<>"", B10<>""), (B8 + B10) / 2, "").
B. Using a filter to identify the tabs with typos or inconsistent names in the "Region" and "Salesperson" columns, I could then use the find and replace function Ctrl + H to correct this. In the region column, we had 66 rows with typos or inconsistent names: Wset(29 replacements), Westt(1 replacement), Easst(34 replacements), Eastt(1 replacement), and Soth(1 replacement).
In correcting the Salesperson column, I noticed we had Ali Kahn and Ali Khan. To determine the correct name, I used the filter option to identify which Ali name was the most, the data showed we had 37 inputs for Ali Kahn and 45 inputs for Ali Khan. Since each transaction's ID is unique, we can assume that these are not duplicated sales recorded and that the typos were in the salesperson's name. Ali Khan had higher inputs than Ali Kahn, hence the replacement for Ali Kahn with Ali Khan. Using Ctrl + H. Also, we had 1 record for J. Smith, which should be Jane Smith, as the first name was recorded as an Initial.
2. Validate Data:
A. Using this formula =IF(OR(ISBLANK(F2), F2=0, ISBLANK(G2), G2=0), "Check Quantity/Price", "") to Identify rows where "Quantity" or "Unit Price" is blank or zero and then flagged them in a new column as "Check Quantity/Price". We have 38 rows that were flagged with this.
## Section 2: Calculations and Analysis
3. Sales Amount Calculation:
A. Using this formula =IF(AND(F2<>0, F2<>"", G2<>0, G2<>""), F2 * G2, 0)  to calculate Total Sales for valid rows (non-blank Quantity and Unit Price Columns). The total sales were $498,577.
4. Sales by Product and Region: 
A.  The East Region had the most sales in the entire region with a total sale of $201,184 while Furniture had the most sales by Product with a total sale of $184,242.
## Section 3: Advanced Analysis and Logical Reasoning
5. To get the top 3 Salesperson, I first copied out the data in the Salespersons and Total sales Column to a new sheet, Then used duplicate function to highlight the 4 salesperson available and this formula =SUMIF(A:A, E2, B:B) to calculate the total sales for each person. I then used the RANK function with this formula to rank them =RANK(F2, F:F). I then used the Index match formula =INDEX(E:E, MATCH(1, G:G, 0)), subsequently for 2 and 3 to get the top 3 salesperson and =INDEX(F:F, MATCH(1, G:G, 0)) to get the total sales respectively. Ali Khan was the Top salesperson with a total sale of $220,890.
6. Dynamic Discount Calculation: 
A. Using this formula =IF(I2>=2000, I2*0.15, IF(I2>=1000, I2*0.10, 0)) to calculate the discounts to the sales, the discount was applied to 128 sales to a total of $68,906.90 applied.
7. Error Check: We had discounts exceeding $300 applied to over 90 sales totaling $63,349.2. This occurred also because the majority of the sales were over $1000 to almost $10,000 and the average sale in this range was $3734.
This is a very high discount and may need to be reviewed for Promotions to ensure we are not reducing our profit margins and assess if it is justifiable in terms of profitability. It also needs to be reviewed to ensure it aligns with our pricing policies.
## Section 4: Decision-Making Scenario
8. Sales Performance Evaluation:
A. To do the Sales Performance Evaluation, I used this formula =IF(C2>=B2, "Met", "Did Not Meet") to calculate the Performance based on the values for the Actual sales and the target sales. All the Regions met and passed their targets but looking at the relative performance amongst the regions, the North region is the region that needs the most improvement as it has the smallest surplus compared to the other regions and its performance is the lowest compared to East and West which has higher actual sales compared to their targets.
