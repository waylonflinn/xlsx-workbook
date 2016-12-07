A simple, friendly interface for [xslx](https://github.com/SheetJS/js-xlsx), that allows easy creation of new Excel spreadsheets and simple editing of existing ones.

Creating a new spreadsheet and adding a little data is easy.

```javascript
// the Worksheet object gives you a simple interface to a single sheet
var Worksheet = require('xlsx-workbook').Worksheet;

var worksheet = new Worksheet("Hello Spreadsheet");
worksheet[0][0] = "Hello";
worksheet[0][1] = "Spreadsheet";

// saving automatically creates a new workbook with the same name
var workbook = worksheet.save();
```


Creating a workbook with multiple sheets is a snap!
```javascript
// the Workbook object gives you more control and stores multiple sheets
var Workbook = require('xlsx-workbook').Workbook;

var workbook = new Workbook();

var sales = workbook.add("Sales");
var costs = workbook.add("Costs");

sales[0][0] = 304.50;
sales[0][1] = 159.24;
sales[0][2] = 493.38;

costs[0][0] = 102.50;
costs[0][1] = 59.14;
costs[0][2] = 273.32;

// automatically appends the '.xlsx' extension
workbook.save("Revenue-Summary");

```

Editing existing workbooks is easy too!
```javascript
var Workbook = require('xlsx-workbook').Workbook;

// looks for a file with a '.xlsx' extension
var wb = new Workbook("Revenue-Summary");

var sales = wb["Sales"];
var costs = wb["Costs"];

var profits = wb.add("Profits");

for(i = 0; i < sales[0].length; i++){
	profits[0][i] = sales[0][i] - costs[0][i];
}

wb.save("Revenue-Summary");

```
