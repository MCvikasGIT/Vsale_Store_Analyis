
Create Table VSales_Store(
transaction_id Varchar(25),
customer_id Varchar(20),
customer_name Varchar(25),
customer_age int,
gender Varchar(25),
product_id float(20),
product_name Varchar(25),
product_category Varchar(25),
quantiy int,
prce int,
payment_mode Varchar(25),
purchase_date	Date,
time_of_purchase	Time,
status Varchar(25)
);

Alter table VSales_Store
Alter Column product_id Varchar(20)

Select  *From VSales_Store

SET dateformat dmy
Bulk insert  VSales_Store
From 'D:\Vikas\Excel projects\Store data.csv'
with (
Firstrow=2,
fieldterminator=',',
rowterminator='\n'
)
----date formate in excel is differt in sql its YYYYMMDD

--Now data cleaning
Select * From VSales_Store

-- data set to be duplicated so that it will be easy to cleaning and analysis so that original data will remain same

Select * From VSales_Store

Select * Into CopyVSales_Store From VSales_Store

Select * From CopyVSales_Store 
Select * From VSales_Store

---Data cleaning step 1
--transaction Id will be unique when ever you purchase any thing it will be give new unique iD so we will check that first

Select  transaction_id,count(*) from CopyVSales_Store
Group by transaction_id
Having COUNT(transaction_id)>1

  ---now we will check one more method by window function by using row number
with CTE as(
 select*,
 ROW_NUMBER() Over (partition by transaction_id order by transaction_id) As Row_Num
 From CopyVSales_Store
)
Select* From CTE
where Row_Num>1




--which bases it is saying duplicate i will check now

with CTE as(
 select*,
 ROW_NUMBER() Over (partition by transaction_id order by transaction_id) As Row_Num
 From CopyVSales_Store
)
Select* From CTE
where transaction_id in('TXN240646','TXN342128','TXN855235','TXN981773')

--- befor deleting we will check once how many rows are present

Select * From CopyVSales_Store
----2005 rows


with CTE as(
 select*,
 ROW_NUMBER() Over (partition by transaction_id order by transaction_id) As Row_Num
 From CopyVSales_Store
)
--Delete from CTE
---where Row_Num=2

---4 rows deleted

Select* From CTE
where transaction_id in('TXN240646','TXN342128','TXN855235','TXN981773')

---So step 1 completed

-- Now step 2 in data cleaning
--Correction of headers quantiy & Prce


Select * From CopyVSales_Store

Execute sp_rename'CopyVSales_Store.Quantiy','Quantity','COLUMN'

Execute sp_rename'CopyVSales_Store.prce','Price','COLUMN'

---its changed 

-- Now step 3 in data cleaning
--To check data type

Select Column_Name, Data_Type
From INFORMATION_SCHEMA.Columns
where Table_Name ='CopyVSales_Store'

--it will provide data type

---step 4 to check null values

Select * from CopyVSales_Store
where transaction_id is Null 
or
customer_id	is Null
or
customer_name	is Null
or
customer_age	is Null
or
gender	is Null
or
product_id	is Null
or
product_name	is Null
or
product_category	is Null
or
Quantity	is Null
or
Price	is Null
or
payment_mode	is Null
or
purchase_date	is Null
or
time_of_purchase	is Null
or
status	is Null

--we will delete null value

delete from CopyVSales_Store
where transaction_id is null

---now we wil see treating Null values
Select * from CopyVSales_Store
where customer_name='Ehsaan Ram'

--Ehsaan Ram have shoped so many time so customer id will be same so we will update here

update CopyVSales_Store
set customer_id='CUST9494'
where transaction_id='TXN977900'

--- customer id will be replaced inplace of Null
--now for another customer


Select * from CopyVSales_Store
where customer_name='Damini Raju'

update CopyVSales_Store
set customer_id='CUST1401'
where transaction_id='TXN985663'

--now we have 1  null left and only customer id is there to check other are null so we try to find and replace the null value
Select * from CopyVSales_Store
where customer_id='CUST1003'

--Mahika saini ,35 age,male no we will replace null

update CopyVSales_Store
set customer_name='Mahika Saini',customer_age='35',gender='Male'
where transaction_id='TXN432798'

---No null value present everything is cleaned

Select * from CopyVSales_Store

--Step 5 of cleaning 
--in gender there different type of formate like F,Female,M, male

Select Distinct gender
from CopyVSales_Store

Update CopyVSales_Store
set gender='Male'
where gender='M'

Update CopyVSales_Store
set gender='Female'
where gender='F'

--Updated and now only 2 types F and M

--Now we check on payment mode

Select Distinct payment_mode
from CopyVSales_Store

Update CopyVSales_Store
set payment_mode='Credit Card'
where payment_mode='CC'

---all data is cleaned

---Now we solve business quaries
--Data analysis

---Question 1
--what are the top 5 most selling products by quantity?
Select * from CopyVSales_Store

select TOP 5  product_name,Sum (Quantity) as total_quanity_sold   from CopyVSales_Store
where status='delivered'
group by product_name
order by total_quanity_sold dESC

Select Distinct status from CopyVSales_Store

---we should only consider deliverd status for top 5 product sold
--Solved

--what was the business problem here?
--Solution = they were not knowing which products are in demand after solving here we got one idea 
--Impact= it helps to priorirtize stock and boost sales through this products

--Question 2 which is most cancelled product?
Select top 5 product_name,Count(*) as total_cancelled from CopyVSales_Store
where status='cancelled'
group by product_name
order by total_cancelled Desc

--what was the business problem here?
--Frequent cancellations affect revenue and customer trust
--impact- they can improve quality of products which are frequenly cancelled products or completely remove from store


---Question 3 what time od the day has the highest number of purchases?

--we use time of purchase by bucketing

Select*from CopyVSales_Store

Select
Case
when Datepart(hour,time_of_purchase) Between 0 and 5 then 'night'
when Datepart(hour,time_of_purchase) Between 6 and 11 then 'Morning'
when Datepart(hour,time_of_purchase) Between 12 and 17 then 'Afternoon'
when Datepart(hour,time_of_purchase) Between 18 and 23 then 'Evening'
End as time_of_day,
Count(*) As total_orders
From VSales_Store
where time_of_purchase is not null
Group By
Case
when Datepart(hour,time_of_purchase) Between 0 and 5 then 'night'
when Datepart(hour,time_of_purchase) Between 6 and 11then 'Morning'
when Datepart(hour,time_of_purchase) Between 12 and 17 then 'Afternoon'
when Datepart(hour,time_of_purchase) Between 18 and 23 then 'Evening'
End 
order By total_orders DESC

---what is the business problem here
---soultion----we can find peak sales time 
--business impact  --we can optimize the maximun staff or promtion

---Question 4 are the top 5 highest spending customers?

Select * from CopyVSales_Store

select top 5 customer_name,
format (Sum (Price*Quantity),'C0','en-IN') As total_spend
from CopyVSales_Store
group by customer_name
order by Sum (Price*Quantity) Desc

--C0 (C stands for currency) it will automatically come with $ if want indian we can change to IN its indian currency

---what was the business problem
--solution - we identified our VIP Customers
--Business impact: we can give personalized offers,or loyality rewards 

---Question 5 which product categories generate the heighest revenue

Select * from CopyVSales_Store

Select product_category,
format (Sum (Price*Quantity),'C0','en-IN') As Revenue
from CopyVSales_Store
Group by product_category
order by Sum (Price*Quantity) Desc

--what was the business problem 
--Solution -- we identified top performing product categories
--business impact ;we can define our product Stratergy and improve our supplychain and we can include some promotions so that it improve our business with high margin

-- Question 6 what is the return/cancellation rate per product category?

Select * from CopyVSales_Store
-- for cancelltion
select product_category,
Format(count(Case when status='cancelled' then 1 end)*100.0/count(*),'N3')+' %' as cancelled_percent
from CopyVSales_Store
Group by product_category
order by cancelled_percent Desc

-- for Return
select product_category,
Format(count(Case when status='returned' then 1 end)*100.0/count(*),'N3')+' %' as Returned_percent
from CopyVSales_Store
Group by product_category
order by Returned_percent Desc

--what was business problem 
--soultion--we can monitor customer dissatisfaction trnds per category
--impact -- we will know we should reduct return product,we can improve prodcut discussion it will help to identfy and fix product issue

--question 7  what is the most preferred payment mode?

Select * from CopyVSales_Store

Select payment_mode,count(payment_mode) as total_count
from CopyVSales_Store
Group by payment_mode
order by total_count Desc

--what was business problem
--solution-- we will know which payment mode most of customers prefer
--impact---we will know which streamline payment should be given more preferness and give multiple option and smooth transction without error (we got credit card)

--question 8 how does age group affect purchasing behavior?
Select * from CopyVSales_Store

Select 
      Case 
           when customer_age between 18 and 25 then'18-25' 
           when customer_age between 26 and 35 then'26-35'
           when customer_age between 35 and 50 then'35-50'
           Else '51+'
           End as Customer_age,
           format(Sum(Price*Quantity),'C0','en-IN')as total_purchase
           from CopyVSales_Store
           Group by case
           when customer_age between 18 and 25 then'18-25' 
           when customer_age between 26 and 35 then'26-35'
           when customer_age between 35 and 50 then'35-50'
           Else '51+'
           End
           order by Sum(Price*Quantity) DESC

     --- what was the business problem
     ---solution;; we understood customer demographics which age is intrested 
     --impact;; we can target this group and market so that we can get more profit

     --Question 9  what 's the monthly sales trend?


     Select * from CopyVSales_Store
--method 1

Select 
format(purchase_date,'yyyy-MM') as Month_year,
format(sum(Price*Quantity),'C0','en-IN') as total_sales,
Sum(Quantity) as total_quantity
from CopyVSales_Store
Group by Format(purchase_date,'yyyy-MM')

---method 2 by using year and month function
     Select * from CopyVSales_Store

     Select 
     --year(purchase_date)as years,
     Month(purchase_date) as months,
     format(sum(Price*Quantity),'C0','en-IN') as total_sales,
Sum(Quantity) as total_quantity
    from CopyVSales_Store
    Group by month(purchase_date)
    order by months

    --2023	1	₹ 46,28,608
---2024	1	₹ 3,39,442

Select(4620608+339442) --4960050


     --- what was the business problem
--solution ; we were not knowing which month had how much sales was done
--impact--; we can plan inventory and marketing according to seasonal trends


---Question 10 are certain genders buying more specific product categories?

     Select * from CopyVSales_Store
     --method 1
Select gender,product_category,count(product_category) as total_purchase
from CopyVSales_Store
group by gender,product_category
order by gender

--method 2
Select *
from(
select gender,product_category
from CopyVSales_Store )
as source_table
pivot(
count (gender)
for gender in( [Male],[Female])
) as pivot_table
order by product_category

--what was businees problem?
--solution; gender base product preference
--impact:we can do personlized ads, gender focused campaigns so that sales can be incresed

--Project analysis done
