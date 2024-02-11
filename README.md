# Excel-Fee-Management-System
![fee system dashboard](https://github.com/jakejosh6751/Excel-Fee-Management-System/blob/main/fee%20system%20dashboard.png)

### 1.	Project Overview

Fee collection is an essential task in schools or any educational institution. Managing all the fees received sometimes become a nuisance with all the manual work involved. Very importantly, the management will always look out for the revenue due and generated from their institution.

This “School Fee Management System” in Microsoft Excel will help educational institutions to work hassle free with regards to fees payment report. It is a comprehensive solution for schools to automate the fee payment reporting process. It ensures a hassle-free payment documentation and visualization, saving over 80 percent of manual work. The module calculates automatically the pending fees, payment history, paid percentage of the total expected fee, and any other metric (to be added) as requested by the management for each selected student or group of students. The fee management system is designed in excel to achieve the following:

* Collect and store students' information – registration number, surname, first name, class, division, boarding/day, new/returning, gender and expected fee;
* Keep records of fee transactions with information such as date of transaction, student’s data, and paid amount; and
* Visualize payment progress in real time using a dashboard that provides key performance indicators that includes but not restricted to expected fees, paid fees and balance (to be paid), as well as charts showing the paid percentage, payment trends (by months and weeks) and fees payment compliance base on year groups.

### 2.	Data Collection & Storage

* VBA enabled data entry UserForms are designed to collect required students and transactions data and store in excel sheets;
* Students' data is stored in “students” sheet.
* Payment Records data is stored in “transactions” sheet.

_students sheet data (below)_:

![students sheet data](https://github.com/jakejosh6751/Excel-Fee-Management-System/blob/main/students%20sheet%20data.png)

_transactions sheet data (below)_:

![transactions sheet data](https://github.com/jakejosh6751/Excel-Fee-Management-System/blob/main/transactions%20sheet%20data.png)

### 3.	VBA Programming

* UserForm module procedures are written to “Add”, “Update”, and “Delete” entries.
* “Export” is used to copy worksheet records to a new excel sheet for ad-hoc analysis.
* “Close” is used to close the UserForm.
* The “List box” displays worksheet data. When a row is double-clicked, its data is made available in the text and combo boxes for editing and resubmission using the “Add” button.
* “Reg. No” text box in the “Students List” UserForm is automated since the registration numbers follow serially.

_students list userform (below)_:

![students list userform](https://github.com/jakejosh6751/Excel-Fee-Management-System/blob/main/students%20list%20userform.png)

_payment records userform (below)_:

![payment records userform](https://github.com/jakejosh6751/Excel-Fee-Management-System/blob/main/payment%20records%20userform.png)

### 4.	Data Extraction, Transformation & Modeling

* Data from the “students” and “transactions” sheets are imported into power query and the data type for each column is transformed.
* For the “students” data, all columns are converted to text data type except “expected fees” which is changed to integer data type.
* For the “transactions” data, all columns are converted to text data type except the “Date” and “Paid fees” columns which are converted to date and integer data types respectively.

Power query serves as a pipeline to update tables and charts when new records are added to the “students” or “transactions” sheets as power pivot (in excel 2013) connected to sheets doesn’t update tables automatically even with a refresh command.

* The “Close & Load” tab in power query adds the “students” and “transactions” tables to new sheets within the same workbook which are then added to a data model in power pivot.
* A “calendar” table which spans the year covered by the academic term is created in Power BI. This table is copied and added in a new sheet in the fees management workbook. The calendar table sheet is then added to the already created data model which now has 3 tables - students, transactions, and calendar.
* Relationships are created in power pivot “Diagram View” between all 3 tables (students, transactions, and calendar) by connecting relevant columns.
* Multiple pivot tables are created to hold information to be displayed on charts and card visuals for the fee management dashboard.

_fee system model diagram (below)_:
![fee system model diagram](https://github.com/jakejosh6751/Excel-Fee-Management-System/blob/main/fee%20system%20model%20diagram.png)

### 5.	Data Visualization

* Card visuals are created for key performance indicators: students count, boarders count, day count, expected fees, paid fees, and balance.
* Donut charts showing payment rates in percent for overall, boarders, and day students.
* Line chart showing monthly and weekly payment trend.
* Clustered bar chart showing paid and expected fees for each year group.
* Filters (new/returning, boarding/day, class, reg. no, surname, and first name) are added to view dashboard for individual students or group of students.

_fee system dashboard (below)_:

![fee system dashboard](https://github.com/jakejosh6751/Excel-Fee-Management-System/blob/main/fee%20system%20dashboard.png)
[See Interactive Dashboard](https://app.powerbi.com/groups/579e1741-4356-4184-93fb-13e61310efdc/reports/d4017250-464e-4403-8326-8921f853f2ee/ReportSection?experience=power-bi)
