/*********************************************************************/
var prevcol
var asc
var col
var type
var qt = "\""

/*********************************************************************/
function create_xls() {

var cnt=0;

	table = document.all.mytable
	var s = ""
	for (var i = 0; i < table.rows.length; i++) {
		var row = table.rows(i)
		var line = ""
		for (var j = 0; j < row.cells.length; j++) {
			if (line != "") line += "\t"

			line += row.cells(j).innerText.replace(/\r\n/ig,'/')
		}
		line += "\n"
		s += line
	}

	document.all.xls.innerText = s
	document.all.xlsholder.style.display = "block"
	document.anchors("xlsanchor").scrollIntoView();

}

/*********************************************************************/
function onload() {
	prevcol = -1
	asc = true
}


/*********************************************************************/
function sort_by_col (c, t) {
	//document.all.wait.innerText = "Sorting.  Please wait..."
	col = c
	type = t
	setTimeout("sort_by_col_impl()",1)
}

/*********************************************************************/
function sort_by_col_impl () {

	table = document.all.mytable
	holder = document.all.myholder

	var myarray = new Array()

	// load an array with the cells
	var i, j
	var len = table.rows.length
	for (i = 1; i < len; i++) {
		j = i - 1
		myarray[j] = new Object
		myarray[j].cells = table.rows(i).cells
		myarray[j].val = table.rows(i).cells(col).innerText
	}


	// clicking on a column twice toggles the sort order
	if (col == prevcol) {
		if (asc)
			asc = false
		else
			asc = true
	}
	else {
		asc = true
	}

	/*
	document.all.sortedby.innerText = "Sorted by "
	+ table.rows(0).cells(col).innerText
	+ (asc ? " asc" : " desc")
	*/

	if (type == "num") {
		myarray.sort(function compare_number(a, b) {
				return flip(parseFloat(a.val) - parseFloat(b.val))
			}
		)
	}
	else if (type == "date") {
		myarray.sort(function compare_date(a, b) {
				if (Date.parse(b.val) < Date.parse(a.val)) {
					return flip(1)
				}
				else {
					return flip(-1)
				}

			}
		)
	}
	else {
		myarray.sort(function compare_string(a, b) {
				if (b.val < a.val)
					return flip(1)
				else
					return flip(-1)
			}
		)
	}
	
	var s = "<table id=" + table.id
	+ " border=0"
	+ " cellspacing=0"
	+ " cellpadding=0"
	+ " class=datagrid" 
	+ "><caption>" + table.createCaption().innerText + "</caption>"
	+ table.rows(0).outerHTML


	// append the sorted rows to the table
	for (i = 1; i < len; i++) {

		var cells = myarray[i - 1].cells
		s += "<tr class=datagrid_row><th>&nbsp;</th>"

		for (var j = 1; j < cells.length; j++) {
			s += cells(j).outerHTML
		}

		s += "</tr>"
	}

	s += "</table>"
	holder.innerHTML = s

	prevcol = col

	//document.all.wait.innerHTML = "&nbsp;"

}

function flip (num) {

	// toggle ascending and descending order
	if (!asc) {
		return -1 * num
	}
	else {
		return num
	}
}
