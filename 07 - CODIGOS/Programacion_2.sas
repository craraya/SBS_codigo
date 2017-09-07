
* ;

%let path=D:/workshop/PG2;

%include "&path/setup.sas";

* Ejemplo;
data forecast;
   set orion.growth;
   Year=1;
   Total_Employees=Total_Employees*(1+Increase);
   *output;
   Year=2;
   Total_Employees=Total_Employees*(1+Increase);
   output;
run;
proc print data=forecast noobs;
   var Department Total_Employees Year;
run;

* Ej 2-16;
data work.price_increase;
   set orion.prices;
   year=1;
   unit_price = unit_price * (factor);
   output;
   year=2;
   unit_price = unit_price * (factor);
   output;
   year=3;
   unit_price = unit_price * (factor);
   output;
run;
proc print data=work.price_increase;
 var Product_ID Unit_Price Year;
run;

* Level 2;
title 'Promotions';
data work.extended;
	set orion.discount;
	where start_date = '01dec2011'd;
	promotion = 'Happy Holidays';
	season = 'Winter';
	output;
	start_date = '01Jul2012'd;
	end_date = '31Jul2012'd;
	season = 'Summer';
	output;
	drop unit_sales_price;
run;
proc print data=work.extended;
run;
title;

* Challenger;

data work.lookup;
	set orion.country;
	if country_formername ~= '' then
	do;
	country_name = country_formername;
	output;
	end;
run;

proc print data=work.lookup;
run;

* Ejemplo;
data usa australia other;
   set orion.employee_addresses;
   if Country='AU' then output australia;
   else if Country='US' then output usa;
   else output other;
run;

  /* alternate solution */
data usa australia other;
   set orion.employee_addresses;
   if Country='US' then output usa;
   else if Country='AU' then output australia;
   else output other;
run;

* Ejemplo select;
data usa australia;
   set orion.employee_addresses;
   select (upcase(Country));
	  when ('US') output usa;
	  when ('AU') output australia;
	  otherwise output other;
   end;
run;

* Ejercicio 2-28;
proc freq data=orion.employee_organization nlevels;
run;

* 1.-;
data admin stock purchasing;
	set orion.employee_organization;
	select (department);
		when ('Administration') output admin;
		when ('Stock & Shipping') output stock;
		when ('Purchasing') output purchasing;
		otherwise;
	end;
run;

* 2.-;
data fast slow veryslow;
	set orion.orders;
	where order_type in (2,3);
	shipdays = delivery_date - order_date;
	select;
		when (shipdays < 3) output fast;
		when (shipdays >= 5 and shipdays <=7) output slow;
		when (shipdays > 7) output veryslow;
		otherwise;
	end;
	drop employee_id;
run;

* Ejercicio 2-45;
* Level 1;
data sales(keep=Employee_id job_title manager_id) exec(keep=employee_id job_title);
	set orion.employee_organization;
	if department = 'Sales' then output sales;
	else if department = 'Executives' then output exec;
run;

proc print data=work.sales(obs=6);
run;

proc print data=work.exec(firstobs=2 obs=3);
run;

* Level 2;
data instore(keep=order_id customer_id order_date) delivery(keep=order_id customer_id order_date shipdays);
	set orion.orders;
	where order_type = 1;
	shipdays = delivery_date - order_date;
	if shipdays = 0 then output instore;
	else if shipdays > 0 then output delivery;
run;

title 'Orders Delivery';
proc print data=delivery;
run;
title;

proc freq data=instore;
	format order_date year4.;
	tables Order_Date;
run;

* Ejercicio 3-15;
* 1.-;
data work.mid_q4;
  set orion.order_fact;
  where '01nov2008'd <= Order_Date <= '14dec2008'd;
  retain num_orders 0 sales2dte 0;
	num_orders=sum(num_orders,1);
	sales2dte = sum(sales2dte,total_retail_price);
run;

title 'Orders from 01Nov2008 through 14Dec2008';
proc print data=work.mid_q4;
  var Order_ID Order_Date Total_Retail_Price sales2dte num_orders;
  format sales2dte dollar10.2;
run;
title;

* 2.-;

data typetotals;
	set orion.order_fact;
	retain Total_Retail 0 Total_Catalog 0 Total_Internet 0;
	if order_type = 1 then total_retail = sum(total_retail,quantity);
	else if order_type = 2 then total_catalog = sum(total_catalog,quantity);
	else if order_type = 3 then Total_Internet = sum(Total_Internet,quantity);
run;

* Summaryze by data step;
* Ejercicio 3-36;
proc sort data=orion.order_summary out=sumsort;
	by customer_id;
run;

data customers(keep=customer_id total_sales);
	set sumsort;
	by customer_id;
	if first.customer_id then total_sales = 0;
	total_sales + sale_amt;
	if last.customer_id; * if last.customer_id then output;
run;

title 'Total Sales to each Customer';
proc print data=customers(obs=10);
run;
title;

* 2.-;

proc sort data=orion.order_qtrsum out=order_strsum_sort;
	by customer_id order_qtr;
run;

data qtrcustomers;
	set order_strsum_sort;
	by customer_id order_qtr;
	if first.order_qtr then 
	do;
	total_sales = 0;
	num_month = 0;
	end;
	total_sales + sale_amt;
	num_month + 1;
	if last.order_qtr;
run;

title 'Total Sales to each Customer for each Quarter';
proc print data= qtrcustomers;
	var customer_id order_qtr total_sales num_month;
run;
title;

* 2.-;
proc sort data=orion.usorders04 out=usorders04_sort;
	by order_type customer_id;
run;

data 	discount1(keep=customer_id customer_name totSales) 
		discount2(keep=customer_id customer_name totSales)
		discount3(keep=customer_id customer_name totSales);
	set usorders04_sort;
	by order_type customer_id;
	if first.customer_id then totSales = 0;
	totSales + total_retail_price;
	if last.customer_id and order_type = 1 then output discount1;
	else if last.customer_id and order_type = 2 then output discount2;
	else if last.customer_id and order_type = 3 then output discount3;
	format totSales dollar11.2;
run;

* Ejercicio Formatted Input;
data sales_staff;
	infile "&path\sales1.dat";
	input @1 emploee_id 6.
			@8 first_name $12.
			@21 last_name $18.
			@40 gender $1.
			@43 job_title $20.
			@64 salary dollar8.
			@73 country $2.
			@76 birth_date mmddyy10.
			@87 hire_date mmddyy10.;
	format birth_date mmddyy10. hire_date mmddyy10. salary dollar8.;
run;

title 'Sales Staff';
proc print data= sales_staff;
run;
title;

data US_trainees(drop=Country) AU_trainees(drop=country);
	infile "&path\sales1.dat";
	input @1 emploee_id 6.
			@8 first_name $12.
			@21 last_name $18.
			@40 gender $1.
			@43 job_title $20.
			@64 salary dollar8.
			@73 country $2.
			@76 birth_date mmddyy10.
			@87 hire_date mmddyy10.;
	format birth_date mmddyy10. hire_date mmddyy10. salary dollar8.;
	if upcase(country) = 'US' and job_title = 'Sales Rep. I' then output US_trainees;
	else if upcase(country) = 'AU' and job_title = 'Sales Rep. I' then output AU_trainees;
run;

title 'Australian Trainees';
proc print data= AU_trainees;
run;
title;

title 'US Trainees';
proc print data= US_trainees;
run;
title;

* Ejercicio 4-43;
* Multi Input Records;
data sales_staff2;
	infile "&path\sales2.dat";
	input @1 employee_id 6.
			@8 first_name $12.
			@21 last_name $18.;
	input @1 job_title $20.
			@22 hire_date mmddyy10.
			@33 salary dollar8.;
	input @1 gender $1.
			@3 birth_date mmddyy10.
			@14 country $2.;
	format salary dollar8. hire_date mmddyy10. birth_date mmddyy10.;
run;

proc print data=sales_staff2;
run;

* 2;
data US_sales AU_sales;
	infile "&path\sales3.dat";
	input @1 employee_id 6.
			@8 first_name $12.
			@21 last_name $18.
			@40 gender $1.
			@43 job_title $20.;
	input @10 country $2. @;
	if upcase(country) = 'AU' then 
		do;
		input @1 salary dolla8.
				@10 country $2.
				@13 birth_date ddmmyy10.
				@24 hire_date ddmmyy10.;
				output AU_sales;
		end;
	else if upcase(country) = 'US' then 
		do;
		input @1 salary dollar8.
				@10 country $2.
				@13 birth_date mmddyy10.
				@24 hire_date mmddyy10.;
				output US_sales;
		end;
run;

* Ejercicio 5-21;
data codes(keep=first_name fcode1 fcode2 last_name lcode);
	set orion.au_salesforce;
	length fcode1 $1. fcode2 $1. lcode $4.;
	fcode1 = lowcase(substr(first_name,1,1));
	fcode2 = lowcase(substr(last_name,length(last_name),1));
	lcode = lowcase(substr(last_name,1,4));
run;

title 'Extracted Letters for User IDs';
proc print data=codes;
run;
title;

* Level 2;
data newcompetitors_code(keep=store_code country city postal_code);
	set orion.newcompetitors;
	country = substr(ID,1,2);
	store_code = left(substr(ID,3));
	city = propcase(city,' ');
	if substr(store_code,1,1)='1';
run;

title 'New Small-Store Competitors';
proc print data=newcompetitors_code;
	var store_code country city Postal_Code;
run;
title;

* Challenge - Use Zip code and location;
data states(keep= ID name location);
	set orion.contacts;
	zip=left(substr(address2,length(address2)-5));
	location = zipname(zip);
run;

title 'States';
proc print data=states;
run;
title;

* Ejercicio 5-41;
data names(keep=new_name name gender);
	set orion.customers_ex5;
	length new_name $25;
	if gender = 'M' then new_name = 'Mr.';
	else if gender = 'F' then new_name = 'Ms.';
	new_name = catx(' ',new_name,substr(name,find(name,', ')+2),substr(name,1,find(name,',')-1));
	new_name = propcase(new_name,' ');
run;

title 'New Names';
proc print data=names;
	var new_name name Gender;
run;
title;

* Optional;
data silver gold platinum;
	keep customer_id name country;
	set orion.customers_ex5;
	substr(customer_id,find(customer_id,'-00-'),4)='-15-';
	if lowcase(substr(customer_id,1,find(customer_id,'0')-1)) = 'silver' then output silver;
	else if lowcase(substr(customer_id,1,find(customer_id,'0')-1)) = 'gold' then output gold;
	else if lowcase(substr(customer_id,1,find(customer_id,'0')-1)) = 'platinum' then output platinum;
run;

title 'Silver-Level Custimers';
proc print data=silver;
run;

title 'Gold-Level Custimers';
proc print data=gold;
run;

title 'Platinum-Level Custimers';
proc print data=platinum;
run;
title;

* Level 2;
data split(keep=employee_id charity);
	set orion.employee_donations;
	charity=substr(recipients,1,find(recipients,'%'));
	if substr(charity,length(charity)) = '%' then output;
	charity=substr(recipients,find(recipients,'%')+3);
	if substr(charity,length(charity)) = '%' then output;
run;

title 'Charity Contributions for each Employee';
proc print data=split;
run;title;

* Ejercicion 5-54;
data sale_stats;
	set orion.orders_midyear;
	monthavg = round(mean(of month1-month6));
	monthmax = max(of month1-month6);
	monthsum = sum(of month1-month6);
	format monthmax comma8.2 monthsum comma8.2;
run;

title 'Statistics on Months in wich the Customer Placed an Order';
proc print data=sale_stats;
	var customer_id monthavg monthmax monthsum;
run;
title;

* Challenge;
data freqcustomers;
	set orion.orders_midyear;
	month_median=median(of month1-month6);
	month_miss=cmiss(of month1-month6);
	month_highest=max(of month1-month6);
	month_2nd_highest=max(of month1-month6);
run;

* Ejercicio 5-79;
proc contents data=orion.shipped;
run;

data shipping_notes;
  set orion.shipped;
  length Comment $ 21;
  Comment = cat('Shipped on ',put(Ship_Date,ddmmyy10.));
  Total = Quantity * input(Price,dollar8.);
run;

proc print data=shipping_notes noobs;
  format Total dollar7.2;
run;
* Level 2;
proc contents data=orion.us_newhire;
run;

data us_converted;
	set orion.us_newhire;
	id_num = input(compress(tranwrd(id,'-','')),8.);
	birth_num = input(birthday,date9.);
	tel = substr(left(put(telephone,8.)),1,3)!!"-"!!substr(left(put(telephone,8.)),4);
	format birth_num ddmmyy10.;
	label id_num='ID' tel='Telephone' birth_num='Birthday';
run;

title 'US New Hires';
proc print data=us_converted;
	*label id_num='ID' tel='Telephone' birth_num='Birthday';
	var id_num tel birth_num;
run;
title;

* Ejercicio 6-13;
  /* Program with a logic error */
data customers;
	set orion.order_summary;
	by Customer_ID;
	putlog _ALL_;
	if first.Customer_ID=1 then Total_Sales=0;
	Total_Sales+Sale_Amt;
	if last.Customer_ID=1;
	keep Customer_ID Total_Sales;
run;

proc print data=customers;
run;

* Ejercicio 7-17;
* DO LOOP;
data future_expenses;
   drop start stop; 
   Wages=12874000;
   Retire=1765000;
   Medical=649000;
   start=year(today())+1;
   stop=start+9;
  /* insert a DO loop here */
	do year = start to stop;
		wages = wages*1.06;
		retire = retire*1.14;
		medical = medical*1.095;
		total_cost = wages + retire + medical;
		output;
	end;
run;
proc print data=future_expenses;
   format wages retire medical total_cost comma14.2;
   var year wages retire medical total_cost;
run;

* Ejercicio 7-29;
data future_expenses;
   drop start stop;
   Wages=12874000;
   Retire=1765000;
   Medical=649000;
	Income=50000000;
   start=year(today())+1;
   stop=start+100;
   do Year=start to stop while(total_cost<=Income);
      wages = wages * 1.06;
      retire=retire*1.014;
      medical=medical *1.095;
		income=income*1.01;
      Total_Cost=sum(wages,retire,medical);
      output;
   end;
run;

proc print data=future_expenses;
	format Year 4. Income total_cost comma14.2;
	var Year Income total_cost;
run;

* Level 2;
data expenses;
	drop start stop;
   Wages=12874000;
   Retire=1765000;
   Medical=649000;
	Income=50000000;
	expenses=38750000;
   start=year(today())+1;
   stop=start+100;
   do Year=start to stop while(expenses<=Income and year-start+1<=30);
      wages = wages * 1.06;
      retire=retire*1.014;
      medical=medical *1.095;
		income=income*1.01;
		expenses=expenses*1.02;
      Total_Cost=sum(wages,retire,medical);
		year2=year-start+1;
   end;
run;

proc print data=expenses;
	format Year 4. Income expenses comma14.2;
	var Year Income expenses year2;
run;

* Array;
data discount_sales;
	drop i;
	set orion.orders_midyear;
	array mon{*} month1-month6;
	do i=1 to dim(mon);
	mon{i}=mon{i}*0.95;
	end;
run;

title 'Monthy Sales with 5% Discount';
proc print data=discount_sales;
	format month1-month6 dollar8.;
run;
title;

* Level 2;
data special_offer;
	set orion.orders_midyear;
	keep total_sales projected_sales difference;
	total_sales = sum(of month1-month6);
	array mon{*} month1-month6;
	do i=1 to 3;
		mon{i}=mon{i}*0.9;
	end;
	projected_sales = sum(of month1-month6);
	difference = total_sales - projected_sales;
	format total_sales projected_sales difference dollar8.2;
run;

title 'Total Sales with 10% Discount in First Three Month';
proc print data=special_offer;
	sum difference;
run;
title;

* Ejercicio 7-62;
data preferred_cust;
   set orion.orders_midyear;
   array Mon{6} Month1-Month6;
   keep Customer_ID Over1-Over6 Total_Over;
   array Target{6} _temporary_ (200,400,300,100,100,200);
   array Over{6};
	do i=1 to 6;
		if mon{i}-target{i} > 0 then over{i}=mon{i}-target{i};
	end;
	total_over = sum(of over{*});
	if total_over > 500 then output;
	format over1-over6 total_over comma8.2;
run;

proc print data=preferred_cust noobs;
run;

* Level 2;
data passed failed;
	set orion.test_answers;
	drop i;
	array resp{*} q1-q10;
	array answer{10} $ 1 _temporary_ ('A','C','C','B','E','E','D','B','B','A');
	score=0;
	do i=1 to dim(resp);
		if resp{i}=answer{i} then score=score+1;
	end;
	if score >= 7 then output passed;
	else output failed;
run;

* Ejercicio 8-23;
data sixmonths(keep=Customer_ID month sales);
	set orion.orders_midyear;
	array sal{6} month1-month6;
	do i=1 to dim(sal);
		if sal{i} ne . then do;
			month=i;
			sales=sal{i};
			output;
		end;
	end;
	format sales comma8.2;
run;

proc print data=sixmonths;
run;

* Level 2;
data expense(keep=trip_id employee_id expense_type amount);
	set orion.travel_expense;
	array ex{5} $20 ("Airfare","Hotel","Meals","Transportation","Miscellaneous");
	array trip{*} exp1-exp5;
	do i=1 to dim(trip);
		if trip{i} ne . then do;
			expense_type=ex{i};
			amount=trip{i};
			output;
		end;
	end;
	format amount dollar10.2;
run;

* Challenge;
proc sort data=orion.order_summary out=order_summary_sort;
	by customer_id order_month;
run;

data customer_orders(keep=Customer_ID month1-month12);
	set order_summary_sort;
	by customer_id;
	retain month1-month12;
	array month{12};
	if first.customer_id then call missing(of month{*});
	month{order_month}=sale_amt;
	if last.customer_id;
run;

* Match Merge;
data 	revenue(keep=product_id price quantity product_name customer revenue) 
		notsold(keep=product_id price product_name)
		invalidcode(keep=product_id quantity customer);
	merge orion.web_products(in=prod) orion.web_orders(in=order);
	by product_id;
	revenue = price*quantity;
	if prod and not order then output notsold;
	if prod and order then output revenue;
	if not prod and order then output invalidcode;
run;

title 'Revenue from Orders';
proc print data=revenue noobs;
run;
title;

title 'Products Not Ordered';
proc print data=notsold noobs;
run;
title;

title 'Invalid Orders';
proc print data=invalidcode noobs;
run;
title;

* Level 2;
data web_converted;
	set orion.web_products2;
	keep product_id_char Name Price;
	product_id_char = put(product_id,$12.);
run;

data 	revenue(keep=product_id price quantity name customer revenue) 
		notsold(keep=product_id price name)
		invalidcode(keep=product_id quantity customer);
	merge orion.web_orders2(in=order rename=(name=customer)) web_converted(in=prod rename=(product_id_char=product_id));
	by product_id;
	revenue = price*quantity;
	if prod and not order then output notsold;
	if prod and order then output revenue;
	if not prod and order then output invalidcode;
run;

title 'Revenue from Orders';
proc print data=revenue noobs;
run;
title;

title 'Products Not Ordered';
proc print data=notsold noobs;
run;
title;

title 'Invalid Orders';
proc print data=invalidcode noobs;
run;
title;

* Ejercicio 10-20;
* Creamos el formato;
data continent;
	keep start label FmtName;
	retain FmtName 'continent';
	set orion.continent(rename=(continent_id=start continent_name=label));
run;

proc format library=orion.MyFmts cntlin=continent;
run;

proc catalog cat=orion.MyFmts;
	contents;
run;

/*******************/
/* Part C          */
/* Use continent.  */
/*******************/

options fmtsearch=(orion.MyFmts);

data countries;
   set orion.country;
   Continent_Name=put(Continent_ID, continent.);
run;

proc print data=countries(obs=10);
   title 'Continent Names';
run;
title;

* Ejercicio 11-20;
proc sql;
  select d.product_id, product_name, start_date, end_date, discount
    from orion.discount as d, orion.product_dim as p
    where d.Product_ID=p.Product_ID;
quit;

* Macros;
%put &sysver;
%put _user_;

* Level 1;
%let country_in=ZA;
title "Customers in &country_in";
proc print data=orion.customer;
   var customer_id customer_name gender; 
   where country="&country_in";
run;
title;

* Level 2;
%let minSal=100000;
title "Employees Earning at Least $&minSal";
proc print data=orion.employee_payroll;
	var Employee_ID Employee_Gender Salary Birth_Date Employee_Hire_Date Employee_Term_Date Marital_Status Dependents;
	format Birth_Date Employee_Hire_Date employee_term_date date9.;
	where salary >= &minSal;
run;
title;

