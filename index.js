var xlsx = require('xlsx');

/* initial code from http://sheetjs.com/demos/writexlsx.html */

// simple type function that works for arrays
function type(obj) { return Object.prototype.toString.call(obj).slice(8, -1);}

function Workbook(sheets){
	this.sheets = [];

	// what kind of argument did we get?
	if(type(sheets) === "String"){
		// string, treat as filename and try to open
		var name = sheets;

		if(!name.endsWith('.xlsx')) name += ".xlsx";

		var wb = xlsx.readFile(name);

		this.sheets = parse(wb);

	} else if(type(sheets) === "Array"){
		for(var i = 0; i < sheets.length; i++){
			this.sheets[i] = sheets[i];
		}
	} else if(type(sheets) === "Uint8Array"){
		var wb = xlsx.read(data, {type: 'binary'});

		this.sheets = parse(wb);
	} else {
		// treat it as a worksheet object
		this.sheets[0] = sheets;
	}

	var name;
	for(var i = 0; i < this.sheets.length; i++){
		name = this.sheets[i].name;
		addSheetProperty(this, name);
	}
}

function addSheetProperty(wb, S){
	Object.defineProperty(wb, S, {
		"enumerable" : true,
		"value" : wb.sheets[S]
	});
}

/* turn an xslx workbook into a Workbook object */
function parse(workbook){

	var ws, name, range;
	var sheets = [];

	for(var i = 0; i < workbook.SheetNames.length; i++){
		name = workbook.SheetNames[i];
		ws = workbook.Sheets[name];
		range = xlsx.utils.decode_range(ws['!ref']);

		// create new Worksheet object
		sheets[i] = new Worksheet(name, range.r);

		// copy data
		sheets[i]

	}

	return sheets;
}

/* add an existing sheet to the Workbook or create a new one with the given name
 */
Workbook.prototype.add = function(sheet){

	if(typeof sheet == "string"){
		var name = sheet;
		sheet = new Worksheet(name);
	}
	this.sheets.push(sheet);

	return sheet;
}

/* turn a Workbook object into something xlsx can understand */
Workbook.prototype.objectify = function(){

	var wb = {
		"SheetNames" : [],
		"Sheets" : {}
	};

	var sheet, name, object;

	for(var i = 0; i < this.sheets.length; i++){
		sheet = this.sheets[i];
		name = sheet.name;
		object = sheet.objectify();

		wb.SheetNames.push(name);
		wb.Sheets[name] = object;
	}

	return wb;
}

Workbook.prototype.save = function(name){

	if(this.sheets.length > 0){
		name = name || this.sheets[0].name;
	}

	wb = this.objectify();

	var filename = name  + ".xlsx";

	xlsx.writeFile(wb, filename, {bookType:'xlsx', bookSST:true, type: 'binary'});
}

Workbook.prototype.push = Workbook.prototype.add;

var DEFAULT_ROWS = 100000;

function Worksheet(name, rows){
	this.name = name;
	rows = rows || DEFAULT_ROWS;

	this.data = [];

	var self = this;
	for(var R = 0; R < rows; R++){
		this.data[R] = [];
		addRowProperty(this, R);
	}
}

function addRowProperty(ws, R){
	Object.defineProperty(ws, R, {
		"enumerable" : true,
		"value" : ws.data[R]
	});
}

/* turn a Worksheet object into something xlsx can understand */
Worksheet.prototype.objectify = function(){

	var ws = {};

	// create base range object
	var range = {s: {c:0, r:0}, e: {c:0, r:this.data.length }};

	// iterate through our dense array
	for(var R = 0; R != this.data.length; ++R) {
		for(var C = 0; C != this.data[R].length; ++C) {

			// update column range, if necessary
			if(range.e.c < C) range.e.c = C;

			// add data
			var cell = {v: this.data[R][C] };
			if(cell.v == null) continue;


			// set the type
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = xlsx.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';

			// generate encoded location
			var cell_ref = xlsx.utils.encode_cell({c:C,r:R});

			// add the cell to the worksheet
			ws[cell_ref] = cell;
		}
	}

	// encode and set range
	ws['!ref'] = xlsx.utils.encode_range(range);

	return ws;
}

/* create a new workbook containing only this sheet with the same name
 */
Worksheet.prototype.save = function(){
	var workbook = new Workbook(this);
	workbook.save();

	return workbook;
}

module.exports = {
	"Workbook" : Workbook,
	"Worksheet" : Worksheet
}
