var xlsx = require('xlsx'),
	path = require('path');

/* initial code from http://sheetjs.com/demos/writexlsx.html */

// simple type function that works for arrays
function type(obj) { return Object.prototype.toString.call(obj).slice(8, -1);}

function fixExtension(filename){
	var ext = path.extname(filename);

	if(ext === '.xlsx') return filename;

	return filename.slice(0, -ext.length) + '.xlsx';
}

function Workbook(sheets){

	Object.defineProperty(this, "sheets", {
		"enumerable" : false,
		"writable" : true,
		"value" : []
	});

	// what kind of argument did we get?
	if(sheets === void(0)){
		this.sheets = [];
	} else if(type(sheets) === "String"){
		// string, treat as filename and try to open
		var name = sheets;

		name = fixExtension(name);

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
		addSheetProperty(this, name, i);
	}
}

function addSheetProperty(wb, S, i){
	Object.defineProperty(wb, S, {
		"enumerable" : true,
		"writable" : false,
		"value" : wb.sheets[i]
	});
}

/* turn an xslx workbook into a Workbook object */
function parse(workbook){

	var ws, name, range;
	var sheets = [];

	var sheet, cell, c;

	for(var i = 0; i < workbook.SheetNames.length; i++){
		name = workbook.SheetNames[i];
		ws = workbook.Sheets[name];
		range = xlsx.utils.decode_range(ws['!ref']);

		// create new Worksheet object
		sheets[i] = new Worksheet(name, range.e.r);

		sheet = sheets[i];
		for (z in ws) {
			/* all keys that do not begin with "!" correspond to cell addresses */
			if(z[0] === '!') continue;

			cell = ws[z];
			c = xlsx.utils.decode_cell(z);

			sheet[c.r][c.c] = cell.v;

		}

		// copy data
		/*
		for(var R = range.s.r; R < range.e.r; R++){
			for(var C = range.s.c; C < range.e.c; C++){
				sheets[R][C] = cell.v;
			}
		}
		*/
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

function fixName(name){

	name = name.replace(/(^\W+)|(\W+$)/g, '');
	name = name.replace(/\W+/g, '-');

	return name;
}

Workbook.prototype.save = function(name){

	if(this.sheets.length > 0){
		name = name || fixName(this.sheets[0].name);
	}

	var filename = fixExtension(name);

	wb = this.objectify();

	// "xlsx" or "xlsm"
	xlsx.writeFile(wb, filename, {bookType:'xlsx', bookSST:true, type: 'binary'});
}

Workbook.prototype.push = Workbook.prototype.add;

var DEFAULT_ROWS = 100000;

function Worksheet(name, rows){

	rows = rows || DEFAULT_ROWS;

	Object.defineProperty(this, "name", {
		"enumerable" : false,
		"writable" : true,
		"value" : name
	});

	Object.defineProperty(this, "data", {
		"enumerable" : false,
		"writable" : true,
		"value" : []
	});

	Object.defineProperty(this, "length", {
		"enumerable" : false,
		"writable" : true,
		"value" : 0
	});


	var self = this;
	if(type(rows) === "Array"){
		var r = Math.max(DEFAULT_ROWS, rows.length)

		for(var R = 0; R < r; R++){
			if(R < rows.length && type(rows[R]).endsWith("Array")){
				this.data[R] = rows[R];

				this.length = (R + 1);
			} else {
				this.data[R] = [];
			}
			addRowProperty(this, R);
		}

	} else {
		for(var R = 0; R < rows; R++){
			this.data[R] = [];
			addRowProperty(this, R);
		}
	}
}

function addRowProperty(ws, R){
	Object.defineProperty(ws, R, {
		"enumerable" : true,
		"get" : function(){
			if(R >= ws.length) ws.length = (R + 1);
			return ws.data[R];
		},
		"set" : function(value){
			if(R >= ws.length) ws.length = (R + 1);

			if(type(value).endsWith("Array"))
				ws.data[R] = value.slice();
			else
				ws.data[R] = [value];
		}
	});
}

/* turn a Worksheet object into something xlsx can understand */
Worksheet.prototype.objectify = function(){

	var ws = {};

	// create base range object
	var range = {s: {c:0, r:0}, e: {c:0, r:this.length }};

	// iterate through our dense array
	for(var R = 0; R != this.length; ++R) {
		for(var C = 0; C != this.data[R].length; ++C) {

			// add data
			var cell = {v: this.data[R][C] };
			if(cell.v == null) continue;

			// update column range, if necessary
			if(range.e.c < C) range.e.c = C;
			if(range.e.r < R) range.e.r = R;

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
Worksheet.prototype.save = function(name){
	var workbook = new Workbook(this);
	workbook.save(name);

	return workbook;
}

module.exports = {
	"Workbook" : Workbook,
	"Worksheet" : Worksheet
}
