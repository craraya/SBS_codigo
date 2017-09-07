proc print data=sashelp.class;
run;
proc means data=sashelp.class;
run;

* creamos la libreria orion con los datos de la clase;
%let path=D:\workshop\PG1;
libname orion "D:\workshop\PG1";

* eliminamos la libreria orion;
libname orion clear;

proc print work.donations;
run;

proc sort data=orion.employee_payroll
	out=work.employee_payroll;
	by Salary;
run;

proc sort data=work.employee_payroll;
	by Employee_Gender descending Salary;
run;

proc print data=work.employee_payroll;
	by Employee_Gender;
run;

* ejercicion 2;
proc sort data=work.employee_payroll
	out=work.sort_sal;
	where Salary > 65000 and Employee_Term_Date=.;	
	by Employee_Gender;
run;

proc print data=work.sort_sal noobs;
	var Employee_ID Salary Marital_Status;
	by Employee_Gender;
	sum Salary;
run;

* Titulos y notas al pie;
title1 'Australian Sales Employees';
title2 'Senior Sales Representatives';
footnote1 'Job Title: Sales Rep. IV';
proc print data=orion.sales;
	var Employee_ID First_Name Last_Name Gender salary;
	where Country='AU' and Job_Title contains 'Rep. IV';
run;

* Tabla con Label;
title 'Entry-level Sales Representatives';
footnote 'Job_Title: Sales Rep. I';

proc print data=orion.sales noobs label;
	where Country='US' and Job_Title='Sales Rep. I';
	var Employee_ID First_Name Last_Name Gender Salary;
	label Employee_ID='Employee ID'
			First_Name='First Name'
			Last_Name='Last Name'
			Salary='Anual Salary';
run;

* Tabla con label y split;
proc print data=orion.sales noobs split='*';
	where Country='US' and Job_Title='Sales Rep. I';
	var Employee_ID First_Name Last_Name Gender Salary;
	label Employee_ID='Employee*ID'
			First_Name='First*Name'
			Last_Name='Last*Name'
			Salary='Anual*Salary';
run;

title;
footnote;

*;
proc sort data=orion.employee_addresses;
	by State City Employee_Name;
run;

proc print data=orion.employee_addresses split='*';
	by State;
	var Employee_ID Employee_Name City Postal_Code;
	label Employee_ID='Employee*ID'
			Employee_Name='Name'
			Postal_Code='Zip*Code';
run;

* Formatos;
proc print data=orion.employee_payroll;
	format Salary dollar8. Birth_Date mmddyy10. Employee_Hire_Date date9.;
	var Employee_ID Salary Birth_Date Employee_Hire_Date;
run;

title1 'US Sales Employees';
title2 'Earning Under $26.000';
proc print data=orion.sales split='*' noobs;
	where Salary < 26000;
	var Employee_ID First_Name Last_Name Job_Title Salary Hire_Date;
label First_Name='First*Name'
		Last_Name='Last Name'
		Job_Title='Title'
		Hire_Date='Date*Hire';
format Salary dollar8. Hire_Date date9. First_Name MSGCASE10.;
run;

* Formatos 2;
data Q1Birthdays;
   set orion.employee_payroll;
   BirthMonth=month(Birth_Date);
   if BirthMonth le 3;
run;

proc format;
	value $gender 	'F'='Female'
					'M'='Male'
					other='Invalid code';
	value mname 1='January'
				2='February'
				3='March';
	value salrange 	20000-<100000='Below $100.000'
					100000-<500000='$100.000 or more'
					.='Missing salary'
					other='Invalid salary';
run;

proc print data=work.q1birthdays;
	var Employee_ID Employee_Gender BirthMonth;
	format Employee_Gender $gender.
			BirthMonth mname.;
run;

*Formatos 3;
proc print data=orion.nonsales;
   var Employee_ID Job_Title Salary Gender;
   title1 'Salary and Gender Values';
   title2 'for Non-Sales Employees';
   format Gender $gender.
   			Salary salrange.;
run;

*;
data work.youngadult;
	set orion.customer_dim;
	where Customer_Gender = 'F' and Customer_Age between 18 and 36 and Customer_Group contains 'Gold';
	Discount = .25;
run;

proc print data=work.youngadult noobs;
	var Customer_ID Customer_Name Customer_Age Customer_Gender Customer_Group Discount;
run;

*;
data work.assistant;
	set orion.staff;
	where Job_Title contains 'Assistant' and Salary < 26000;
	Increase = Salary * 0.10;
	New_Salary = Salary + Increase;
run;

proc print data= work.assistant;
	var Employee_ID Job_Title Salary Increase New_Salary;
	format Increase dollar8. New_Salary dollar8.;
run;

* p106e04;
* 1;
data work.increase;
   set orion.staff;
   where Emp_Hire_Date > '01Jul2010'd;
   Increase=Salary*0.10;
   NewSalary=Salary+Increase;
   if Increase > 3000;
   keep Employee_ID Emp_Hire_Date Salary Increasse NewSalary;
   label Employee_ID='Employee ID' Emp_Hire_Date='Hire Date' NewSalary='New Annual Salary';
   format Employee_ID 12. Emp_Hire_Date date9. NewSalary dollar10.2 Salary dollar10.2 Increasse comma5.;
run;

proc print data=work.increase;
run;

proc contents data= work.increase;
run;

proc print data=work.increase split='*';
	var Employee_ID Salary Emp_Hire_Date Increasse NewSalary;
	label Employee_ID='Employee*ID' Salary='Employee*Annual*Salary' Emp_Hire_Date='Hire*Date' Increasse NewSalary='New*Annual*Salary';
run;

* 2;
data work.delays;
	set orion.orders;
	where Delivery_Date - Order_Date > 4 and Employee_ID = 99999999;
	Order_Month = month(Order_Date);
	if Order_Month = 8;
	keep Employee_ID Customer_ID Order_Date Delivery_Date Order_Month;
	label Employee_ID='Employee ID' Customer_ID='Customer ID' Order_Date='Date Ordered' Delivery_Date='Date Delivered' Order_Month='Month Ordered';
	format Order_Date mmddyy10. Delivery_Date mmddyy10.;
run;

proc contents data=work.delays;
run;

proc print data=work.delays;
run;

data work.bigdonations;
	set orion.employee_donations;
	call nmiss(Qtr1,Qtr2,Qtr3,Qtr4);
run;

* Archivos - librerias - Excel;
* 1;
options validvarname=v7; 
libname custfm excel "&path\custfm.xlsx";

proc contents data=custfm._all_;
run;

data work.males;
	set custfm.'Males$'n;
	keep First_Name Last_Name Birth_Date;
	format Birth_Date year4.;
	label Birth_Date='Birth Year';
run;

proc print data=work.males label;
run;

libname custfm clear;

* 2;
options validvarname=v7; 
libname prod excel "&path\products.xlsx";

proc contents data=prod._all_;
run;

data work.golf;
	set prod.'Sports$'n;
	where Category = 'Golf';
	drop Category;
	label Name='Golf Products';
run;

libname prod clear;

proc print data=work.golf;
run;

* 3;

 data salesemps;
    length First_Name $ 12 Last_Name $ 18
           Job_Title $ 25;
    infile "&path\newemps.csv" dlm=','; 
    input First_Name $ Last_Name $
          Job_Title $ Salary;
run;

proc export data=work.salesemps outfile="&path\salesemps.xlsx";
	sheet='sales';
run;

* Importación de archivos;
* 1;
data work.nonsales2;
	length Employee_ID 8 First $ 12 Last $ 18 Gender $ 2 Salary 8 Title $ 25 Country $ 2;
	infile "&path\nonsales.csv" dlm=',';   
	input Employee_ID First $ Last $ Gender $ Salary Title $ Country $;
run;

proc print data=work.nonsales2;
	var First Last Title Salary;
run;

* 2;
data work.qtrdonation;
	length IdNum $ 6 Qtr1 8 Qtr2 8 Qtr3 8 Qtr4 8;
	infile "&path\donation.dat" dlm=' ';   
	input IdNum $ Qtr1 Qtr2 Qtr3 Qtr4;
run;

proc print data=work.qtrdonation;
run;

* Challenger;
data work.managers2;
	length ID 8 First $ 12 Last $ 18 Gender $ 2 Salary 8 Title $ 25;
	infile "&path\managers2.dat" dlm='\t';   
	input ID First $ Last $ Gender $ Salary Title $;
run;

* List imput - colon format;
   /* Part 1 - using colon format modifiers*/
data work.salaries;
	infile "&path\salary.dat";
	input Name $ HireDate :date. Age State $ Salary :comma.;
run;

proc print data=work.salaries;
run;

  /* Part 2 - omit the colon format modifier for Salary */
data work.salaries;
	infile "&path\salary.dat";
	input Name $ HireDate :date9. Age State $ Salary dollar.;
run;

proc print data=work.salaries;
run;

* Ex 1;
data work.canada_customers;
	
	infile "&path\custca.csv" dsd;
	input First :$20. Last :$20. ID Gender :$1. BirthDate :ddmmyy10. Age AgeGroup :$12.;
	label BirthDate='Birth*Date';
run;

proc print data=work.canada_customers split='*';
	var First Last Gender AgeGroup BirthDate;
run;

* 2;
data work.prices;	
	infile "&path\pricing.dat" dlm='*';
	input ProductID StartDate :date9. EndDate :date9. Cost :dollar8. Sales_Price :dollar8.;
run;

proc print data=work.prices;
	format StartDate ddmmyy10. EndDate ddmmyy10.;
run;

* Ex;
data work.donations;
	infile "&path\donation.csv" dsd missover;
	input EmpID Q1 Q2 Q3 Q4;
run;

proc print data=work.donations;
run;

* 2;
data work.prices;
	infile "&path\prices.dat" dlm='*' dsd missover;
	input ProductID StartDate:date9. EndDate:date9. Cost:dollar8. Sales_Price:dollar8.;
run;

proc print data=work.prices;
	format StartDate ddmmyy10. EndDate ddmmyy10.;
run;

* Challenger;
data work.salesmgmt;
	infile "&path\managers.dat" dlm='/' dsd missover;
	input ID First:$15. Last:$15. Gender:$1. Salary Title:$20. Country:$15. StartDate:date9. HireDate:mmddyy10.;
run;

proc print data=work.salesmgmt;
	var ID last Title HireDate Salary;
	format HireDate date9.;
run;

* data functions;
data work.increase;
	set orion.staff;
	Increase = Salary * 0.010;
	NewSalary = Salary + Increase;
	BdayQtr = qtr(Birth_Date);
	keep Employee_ID Salary Birth_Date Increase NewSalary BdayQtr;
	format Salary comma12. Increase comma12. NewSalary comma12.;
run;

proc print data=work.increase split='*';
	label Employee_ID='Employee*ID' Salary='Employee*Annual*Salary' Birth_Date='Employye*Birth Date' BdayQtr='Bday*Qtr';
run;

* 2;
data work.birthday;
	set orion.customer;
	Bday2012 = mdy(month(Birth_Date),day(Birth_Date),2012);
	BdayDOW2012 = weekday(Bday2012);
	Age2012 = (Bday2012-Birth_Date)/365.25;
	format Bday2012 date9.;
	keep Customer_Name Birth_Date Bday2012 BdayDOW2012 Age2012;
run;

proc print data=work.birthday;
	format Age2012 comma.;
run;

* Challenger;
data work.employees;
	set orion.sales;
	FullName=catx(' ',First_Name,Last_Name);
	Yrs2012=Hire_Date-'01Jan2012'd;
run;

* If Else;
data work.ordertype;
	set orion.orders;
	length Method $ 15;
	if Order_Type = 1 then Method = 'Retail';
	else if Order_Type = 2 then Method = 'Catalog';
	else if Order_Type = 3 then Method = 'Internet';
	else Method = 'Unknown';
run;

proc print data=work.ordertype;
	var Order_Id Order_Type Method;
run;

* 1.2;
data work.region;
	set orion.supplier;
	length DiscountType $ 20 Region $ 50;
	if upcase(Country) in ('CA','US') then do;
	Discount = 0.10;
	DiscountType = 'Required';
	Region = 'North America';
	end;
	else do;
	Discount = 0.05;
	DiscountType = 'Optional';
	Region = 'Not North America';
	end;
	keep Supplier_Name Country Discount DiscountType Region;
run;

proc print data=work.region;
	var Supplier_Name Country Region Discount DiscountType;
run;

* 2;
data work.season;
	set orion.customer_dim;
	length Promo $ 25 Promo2 $ 15;
	if qtr(Customer_BirthDate) = 1 then Promo = 'Winter';
	else if qtr(Customer_BirthDate) = 2 then Promo = 'Spring';
	else if qtr(Customer_BirthDate) = 3 then Promo = 'Summer';
	else if qtr(Customer_BirthDate) = 4 then Promo = 'Fall';

	if Customer_Age >= 18 and Customer_Age <= 25 then Promo2 = 'YA';
	else if Customer_Age >= 65 then Promo2 = 'Senior';
	
	keep Customer_FirstName Customer_LastName Customer_BirthDate Customer_Age Promo Promo2;
run;

proc print data=work.season;
run;

data work.ordertype;
	set orion.orders;
	length Type $ 20 SalesAds $ 20;
	DayOfWeek = weekday(Order_Date);
	if Order_Type = 1 then Type = 'Retail Sale';
	else if Order_Type = 2 then Type = 'Catalog Sale';
	else if Order_Type = 3 then Type = 'Internet Sales';

	if Order_Type = 2 then SalesAds = 'Mail';
	else if Order_Type = 3 then SalesAds = 'Email';

	drop Order_Type Employee_Id Customer_ID;
run;

* Unions;
data work.thirdqtr;
	set orion.mnth7_2011 orion.mnth8_2011 orion.mnth9_2011;
run;

proc print data=work.thirdqtr;
run;

* 1.2;
proc contents data=orion.sales;
run;

proc contents data=orion.nonsales;
run;

data work.allemployee;
	set orion.sales orion.nonsales(rename=(First=First_Name Last=Last_Name));
	keep Employee_ID First_Name Last_Name Job_Title Salary;
run;

proc print data= work.allemployee;
run;

* 2;
proc contents data=orion.charities;
run;

proc contents data=orion.us_suppliers;
run;

proc contents data=orion.consultants;
run;

data work.contacts;
	set orion.charities orion.us_suppliers orion.consultants;
run;

/*********************************************************
*  1. Complete the program to match-merge the sorted     *
*     SAS data sets referenced in the PROC SORT steps.   *                                  *
*  2. Submit the program. Correct and resubmit,          *
*     if necessary.                                      * 
*  4. What are the modified, completed statement?        *
*********************************************************/

proc sort data=orion.employee_payroll
          out=work.payroll; 
   by Employee_ID;
run;

proc sort data=orion.employee_addresses
          out=work.addresses;
   by Employee_ID;
run;

data work.payadd;
   merge work.payroll work.addresses;
   by Employee_ID;
run;

proc print data=work.payadd;
   var Employee_ID Employee_Name Birth_Date Salary;
   format Birth_Date weekdate.;
run;

* Merge;
proc contents data=orion.orders;
run;

proc contents data=orion.order_item;
run;

data work.allorders;
	merge orion.orders orion.order_item;
	by Order_ID;
	keep Order_ID Order_Item_Num Order_Type Order_Date Quantity Total_Retail_price;
run;

proc print data=work.allorders;
	where qtr(Order_Date)=4 and year(Order_Date)=2011;
run;

proc sort data=orion.product_list out=work.product_list;
	by Product_Level;
run;

data work.listlevel;
	merge orion.product_level work.product_list;
	by Product_Level;
	keep Product_ID Product_Name Product_Level Product_Level_Name;
run;

proc print data=work.listlevel;
	where Product_Level=3;
run;

proc sort data=orion.customer
          out=cust_by_type;
	by Customer_Type_ID;
run;

data customers;
   merge cust_by_type orion.customer_type;
	by Customer_Type_ID;
	if Country='US';
run;

* Ej 1;
proc sort data=orion.product_list
          out=work.product;
	by Supplier_ID;
run;

data work.prodsup;
	merge work.product(in=prd) orion.supplier(in=supp);
	by supplier_id;
	if prd and not supp;
run;

proc print data=work.prodsup;
	var Product_ID Product_Name Supplier_ID Supplier_Name;
run;

* 2;
proc sort data=orion.customer out=work.customer;
	by country;
run;

data work.allcustomer;
	merge orion.lookup_country(rename=(start=country label=country_name) in=count) work.customer(in=cust);
	by country;
	keep customer_id country customer_name country_name;
	if count and cust;
run;

proc print data=work.allcustomer;
run;

* challenger;
proc sort data=orion.orders out=work.orders;
	by employee_id;
run;

data work.allorders;
	merge orion.staff(in=staff) work.orders(in=ord);
	by employee_id;
	if ord;
	keep employee_id job_title gender order_id order_type order_date;
run;

data work.noorders;
	merge orion.staff(in=staff) work.orders(in=ord);
	by employee_id;
	if staff and not ord;
	keep employee_id job_title gender order_id order_type order_date;
run;

proc print data=work.allorders;
	title 'All Orders';
run;

proc print data=work.noorders;
	title 'No Orders';
run;
title clear;

* Tablas Proc Frec;
proc freq data=orion.sales;
	tables gender*country;
run;

proc freq data=orion.sales /nlevels;
	tables gender country /noprint;
run;

proc freq data=orion.nonsales2 order=freq;
   tables Job_Title / nocum nopercent;
run; 

proc freq data=orion.nonsales2 nlevels;
   tables Job_Title / nocum nopercent;
run; 

proc freq data=orion.sales;
	tables Hire_date / nocum;
	format hire_date year4.;
run;

* Ej 1;
proc freq data=orion.orders;
   tables Customer_ID Employee_ID;
run;

proc freq data=orion.orders nlevels;
	tables Customer_ID Employee_ID / nofreq nocum nopercent;
	where order_type = 1;
run;

* 1.2;
proc freq data=orion.orders nlevels order=freq;
	tables Customer_ID / nofreq nocum;
	where order_type ~= 1;
run;

proc freq data=orion.shoes_tracker nlevels;
	tables Supplier_Name supplier_id;
run;

* 2;
proc format;
   value ordertypes
      1='Retail'
      2='Catalog'
      3='Internet';
run;

proc freq data=orion.orders;
	tables order_date order_type order_date*order_type;
	format order_date year4. order_type ordertypes.;
run;

* 2.2;

proc freq data=orion.qtr2_2011 nlevels;
	tables order_id order_type;
run;

* challenger;
proc freq data=orion.order_fact noprint;
	tables product_id / out=product_orders;
run;

data work.product_names;
	merge product_orders orion.product_list;
	by product_id;
	keep product_id product_name count;
run;

* proc mean;
proc format;
   value ordertypes
      1='Retail'
      2='Catalog'
      3='Internet';
run;

title 'Revenue from All Orders';
proc means data=orion.order_fact sum;
	var Total_Retail_Price;
	class order_date order_type;
	format order_date year4. order_type ordertypes.;
run;
title;

proc univariate data=orion.price_current nextrobs=5;
	var Unit_Sales_Price factor;
run;

* ODS;
proc template;
	list styles;
run;

options nodate nonumber;
title; footnote;

ods listing; /* OUTPUT window in SAS Windowing Environment */
ods html file="&path\report\myreport.html";
ods pdf file="&path\report\myreport.pdf";
ods rtf file="&path\report\myreport.rtf";

title 'Report 1';
proc freq data=orion.sales;
   tables Country;
run;

title 'Report 2';
proc means data=orion.sales;
   var Salary;
run;

title 'Report 3';
proc print data=orion.sales;
   var First_Name Last_Name 
       Job_Title Country Salary;
   where Salary > 75000;
run;

ods _all_ close;
ods html; /* SAS Windowing Environment */

* Ej 1;
options nodate nonumber;
title; footnote;

ods listing; /* OUTPUT window in SAS Windowing Environment */
*ods html file="&path\report\myreport.html";
ods pdf file="&path\report\p111e12.pdf" style=curve;
ods rtf file="&path\report\p111e12.rtf" style=journal;

title 'July 2011 Orders';
proc print data=orion.mnth7_2011;
run;

ods _all_ close;
ods html; /* SAS Windowing Environment */

