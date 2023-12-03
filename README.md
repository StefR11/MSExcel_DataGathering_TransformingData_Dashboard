In this End-to-End Excel Project I will gather the data first, transform it then use Pivot Table and Pivot Chart to build a dynamic interactive Coffee Sales Dashboard.
DASHBOARD VISUALS: 
In this Dashboard I am going to have
•	one Line Chart with Total Sales broken down by four coffee types
•	a Bar Chart with Total Sales broken down by the 3 countries US, UK and Ireland where coffee is sold
•	top 5 customer Bar Chart
•	a Timeline that the user can use to manipulate these visuals
•	a Slicer showing roast type names
•	a Slicer for size of sold packages
•	a slicer showing whether or not the customers have a loyalty card
EXCEL WORKBOOK USED:
CoffeeOrdersData.xlsx (database on Coffee Bean Sales) made up of  Orders, Customers and Products worksheets
FIRST STEPS > GATHERING AND TRANSFORMING DATA
•	populating Orders sheet’s cells  from column F to column M  by means of  XLOOKUP and the dynamic INDEX MATCH function to be able to get the needed values  from  Customers and Products sheets
=XLOOKUP(orders!C2, customers!$A$1:$A$1001, customers!$B$1:$B$1001)
=IF(XLOOKUP(orders!C3,customers!$A$1:$A$1001,customers!$C$1:$C$1001)= 0, "", (XLOOKUP(orders!C3,customers!$A$1:$A$1001,customers!$C$1:$C$1001)))
=INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0))
•	populating the column Sales by multiplying Quantity and Unit Price
•	populating and categorising the newly created column called Coffee Type Name in the Orders sheet showing the Full Coffee type Names 
=SWITCH(TRUE(), I2 = "Rob", "Robusta", I2 = "Exc", "Excelsa", I2 = "Ara", "Arabica", I2 = "Lib", "Liberica")
•	populating and categorising the newly created column called Roast Type Name in the Orders sheet showing the Full Roast type Names 
=SWITCH(TRUE(), J2 = "M", "Medium", J2 = "L", "Light", J2 = "D", "Dark", "")

•	formatting cells and checking for duplicates values 
SECOND STEPS > PIVOT TABLE (ALT N V T) AND PIVOT CHART
•	First Pivot Table > Sales broken down by Year, Month and Coffee Type Name
•	Line Chart > Sales Over Time
•	Insert and formatting Timeline > Pivot Chart Analyse > Insert Timeline 
•	creating a new column called Loyalty Card  in the Orders sheet to see whether customers have a Loyalty Card
=XLOOKUP([@[Customer Name]], customers!$B$1:$B$1001,customers!$I$1:$I$1001)
•	clicking either into the Pivot Table or Pivot Chart and Refresh
•	Inserting and formatting 3 Slicers showcasing Size, Roast Coffee Name and Loyalty Card
•	Bar Chart showing  Sales broken down by Country
•	Bar Chart with top  5 customers generating most sales
•	Making slicers and timeline filter all of the visuals > Timeline / Slicers > Report Connections
•	Adding protected specific cells within the Orders, Customers, Products worksheets to prevent users from changing data > Review Tab > Protect Worksheet
•	Protecting the Structure of the entire workbook to prevent unwanted changes, such as moving, deleting, renaming or adding sheets > Review Tab > Protect Workbook



