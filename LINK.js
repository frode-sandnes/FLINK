// By Frode Eika Sandnes, Oslo Metropolitan University, Oslo, Norway, May 2021

// global variables
var file1, file2;
var linked;

// update GUI when slider moves
function acceptChange()
	{
	var value = document.getElementById("acceptLimitId").value;
	document.getElementsByName("acceptLabel")[0].textContent = "Accept limit ("+value+")";	
	}
// string distance measure using DICE
function dice(l, r) 
	{
	if (l.length < 2 || r.length < 2) return 0;
	let lBigrams = new Map();
	for (let i = 0; i < l.length - 1; i++) 
		{
		const bigram = l.substr(i, 2);
		const count = lBigrams.has(bigram)
		? lBigrams.get(bigram) + 1
		: 1;
		lBigrams.set(bigram, count);
		};
	let intersectionSize = 0;
	for (let i = 0; i < r.length - 1; i++) 
		{
		const bigram = r.substr(i, 2);
		const count = lBigrams.has(bigram)
		? lBigrams.get(bigram)
		: 0;
		if (count > 0) 
			{
			lBigrams.set(bigram, count - 1);
			intersectionSize++;
			}
		}
	// 6/11/2022 changed to spcial purpose adaptation to adopt for name pairs with missing names
	// simply compare intersection to the shortest string only, not both
	// based on the assumption there that the longest string is most correct and complete
	minLength = (l.length > r.length)?r.length:l.length;
	return intersectionSize/minLength;
//	return (2.0 * intersectionSize) / (l.length + r.length - 2);
	}	
// retrieving file contents in excel format	
function loadBinaryFile(selector)
	{
	const fileSelector = document.getElementById("file-selector"+selector);
	fileSelector.addEventListener('change', (event) => 
		{
		const files = event.target.files;
	
		for (var i = 0, f; f = files[i]; i++) 
			{			
			var reader = new FileReader();
			
			reader.onload = (function(theFile) 
				{
				return function(e) 
					{
					var workbook = XLSX.read(e.target.result, {type: 'binary'});	
					for (var sheetName of workbook.SheetNames)
						{
						var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
						var json_object = JSON.stringify(XL_row_object);
						json_object = json_object.toLowerCase(); // unify comparisons by converting all to lowercase (simplifies matching and string based indexing)
/**/					XL_row_object = JSON.parse(json_object); // ensure the original binary object also in lowercase, assuming everything is in lowerCase from here on - less code.
						if (selector === "1")
							{
							file1 = json_object;	
							}
						else
							{
							file2 = json_object;								
							}	
						outputGrid(XL_row_object,"table"+selector,"",'gray',true,3);
						mergeOnFirstColumns();						
						}						
					};
				})(f);		
			reader.readAsBinaryString(f);
			}
		});
	}

// modifed to include - hyphen between name and capitalization of first character after hyphen	
const capitalize = (str, lower = false) =>
  (lower ? str.toLowerCase() : str).replace(/(?:^|\s|[-"'([{])+\S/g, match => match.toUpperCase());
//  (lower ? str.toLowerCase() : str).replace(/(?:^|\s|["'([{])+\S/g, match => match.toUpperCase());
	;


// output JSON in table format
//function outputGrid(jsn, id, limit = 1000)
function outputGrid(obj, id, heading, colour, checkboxes, limit = 1000)
	{
	if (obj.length === 0) // only proceed if we have elements
		{
		document.getElementById(id).innerHTML = ""; // if empty table, then we need to clear the html element also			
		return;
		}
	// make the text lighter as it is just for verification and not central for understanding the information.
	document.getElementById(id).style.color = colour;								
		
	// the json object is stored as array with one element per row, and column represented with column name as key, and cell as value.
	
	// first build a list of header names based on the first row
	var table = "<h2>"+heading+"</h2><small><table><tr>";
	var headers = [];
	for (var head in obj[0])
		{
		headers.push(head);
		table += "<th>";
		var checkboxid = id+head;
		if (checkboxes)	// if we are to associate checkbox with header?
			{
			table += "<input type=\"checkbox\" id=\""+checkboxid+"\" name=\""+checkboxid+"\" value=\"Bike\">";
			table += "<label for=\""+checkboxid+"\">"+head+"</label><br>";
			}	
		else
			{
			table += head;
			}
		table += "</th>";			
		}		
	table += "</tr><tr>";			
	// then traverse each row and recall the value by looping through the keys (accessed via bracket notation).
	var count = 0; // output line counter
	for (var line of obj)
		{
		for (var head of headers)
			{
			// get string from sheet
			var str = line[head];
			// capitalize names
			if (typeof(str) != "undefined")
				{
				str = capitalize(str);
//To get capitalized result in file also 		
				line[head] = str;
				}	
			// put in table			
			table += "<td>"+str+"</td>";					
			}			
		if (++count >= limit)
			{
			break;	// return if not more
			}
		table += "</tr><tr>";				
		}
	table += "</tr></table></small>";
	document.getElementById(id).innerHTML = table;	
	// if checkboxes - set the first column as default
	if (checkboxes)	
		{
		var checkboxid = id+headers[0];	
		document.getElementById(checkboxid).checked  = true;			
		}		
	}
// get table headers
function getHeaders(obj)
	{
	var headers = [];
	for (var head in obj[0])
		{
		headers.push(head);
		}	
	return headers;
	}

// get all the values in cells for a given set of columns with checkboxes set	
function getColumns(o,id,headers)
	{
	var keylist = [];
	var rowlist = [];
	for (var row of o)	
		{
		var value = "";			
		for (var i = 0;i<headers.length; i++)
			{
			var checkboxid = id+headers[i];		// refer back to the tpreview table in the form in the gridouput routine
			if (document.getElementById(checkboxid).checked) 
				{
				var head = headers[i];
				if (typeof row[head] !== 'undefined') // check that it actually exist
					{					
					if (value !== "")
						{
						value += " "; // add space separator
						}
					value += row[head];		// if the checkbox is set, the value in the cell is added to the key
					}					
				}
			}		
		keylist.push(value);
		rowlist.push(row);
		}
	return {keys: keylist, rows: rowlist};
	}		

// remove a set of items in onle list from another	
function subtractElements(list1,list2)
	{
	for (var v of list2)
		{
		if (list1.includes(v))
			{
			list1.splice(list1.indexOf(v), 1);		// removing the element without leaving a hole	
			}
		}					
	}		
// create a new json object for the three first inspection columns	
function createRecord(distance,key1,key2,value1,value2)
	{
    return {similarity: distance, key1: value1, key2: value2};
	}

// find a given row in the json spreadsheet	
function getRecord2(keyrow,key)
	{
	var index = keyrow.keys.indexOf(key);
	return keyrow.rows[index];
	}
	
// rename the key of a simple ojbect	
function renameKey(obj, oldKey, newKey) 
	{   
    Object.defineProperty(obj, newKey, Object.getOwnPropertyDescriptor(obj, oldKey));
    delete obj[oldKey];                              
    }	
// add suffix to all header elements to make these unique	
function alterRecordHeader(r,suffix)
	{
	for (var head in r)
		{
		renameKey(r,head,head+suffix);
		}
	}
// find set of elements in two lists that are above the threshold in similarity	
function matchingElements(o1,o2,key1,key2,keylist1,keylist2,acceptLimit,keyrow1,keyrow2)
	{
	var matching = [];
	var selected1 = [];
	for (var value1 of keylist1)	
		{
		var max = 0;
		var maxItem;
		// for each row compare with each row of file 2
		for (var value2 of keylist2)	
			{	
			var distance = dice(value1,value2);	
			// find the largest one
			if (distance > max)
				{
				max = distance;
				maxItem = value2;
				}	
			}
		// if above or equal to accept limit
		if (max >= acceptLimit)
			{	
			var r1 = getRecord2(keyrow1,value1);	
			alterRecordHeader(r1,"-1");
			var r2 = getRecord2(keyrow2,maxItem);			
			alterRecordHeader(r2,"-2");
			var r3 = createRecord(max.toFixed(2),key1,key2,value1,maxItem);
			var r = Object.assign(r3, r1, r2);		// merging objects
			matching.push(r);
			selected1.push(value1);
			//  remove items that have been selected
			keylist2.splice(keylist2.indexOf(maxItem), 1);
			}				
		}
	//  remove items that have been select
	subtractElements(keylist1,selected1);
	return matching;
	}
// return elements in the same list that are above the threshold of similarity	
function similarElements(o,key,keylist,acceptLimit)
	{
	var matching = [];
	for (var i = 0;i<keylist.length;i++)
		{			
		var value1 = keylist[i];
		// for each row compare with each row of file 2
		for (var j = i+1;j<keylist.length;j++)
			{	
			var value2 = keylist[j];
			var distance = dice(value1,value2);	
			// if above or equal to accept limit
			if (distance >= acceptLimit)
				{	
				var r = createRecord(distance.toFixed(2),key,key,value1,value2);
				matching.push(r);
				}			
			}
		}
	return matching;
	}	

// find records for keylist
function retrieveElements2(keylist,keyrow)
	{
	var matching = [];
	for (var key of keylist)	
		{
		var r = getRecord2(keyrow,key);			
		matching.push(r);
		}
	return matching;
	}	
// attach a distance undefined column so that users of spreadsheet sees the full range of records	
function labelAsUnmatched(unmatched)
	{
	var result = [];
	for (var row of unmatched)
		{
		var r = {similarity: "no-match"};			
		var newRow = Object.assign(r, row);		// merging objects
		result.push(newRow);
		}
	return result;
	}
// check for duplicates and output warning
function checkDuplicates(id,keylist)
	{
	var warning = "";
    // create a Set with array elements
    var s = new Set(keylist);
    // compare the size of array and Set
    if(keylist.length !== s.size)
		{
        warning = "<p>Warning - Duplicate entries in File.</p>";	
		}
	document.getElementById(id).innerHTML = warning;	
	document.getElementById(id).style.color = "red";	
	}
	
// merges the two files based on the two respective first columns 
// (later expand to arbitrarily selected columns)
function mergeOnFirstColumns()
	{
	linked = "";	// effectively clear the global object before it is used again

	// check if both files are set
	if (typeof file1 === 'undefined' || typeof file2 === 'undefined')	
		{
		return;
		}

	// setup
	var o1 = JSON.parse(file1);
	var o2 = JSON.parse(file2);
	var acceptLimit = document.getElementById("acceptLimitId").value/100;
		
	// get header for first rows
	var headers1 = getHeaders(o1);
	var headers2 = getHeaders(o2);
	var key1 = headers1[0];	// use first header item for comparison
	var key2 = headers2[0];	// use first header item for comparison	
	// create keylist for files 1 and 2 vased on respective columns in the sheet
	var keyrow1 = getColumns(o1,"table1",headers1);
	var keyrow2 = getColumns(o2,"table2",headers2);
	var keylist1 = [...keyrow1.keys];	// shallow true copy of arrays
	var keylist2 = [...keyrow2.keys];	
	var rowlist1 = keyrow1.rows;
	var rowlist2 = keyrow2.rows;	
	// check for duplicated entries, if duplicates - output warning
	checkDuplicates("messageFile1",keylist1);
	checkDuplicates("messageFile2",keylist2);
	// First, find cases in each file that are very similar
	var similarList1 = similarElements(o1,key1,keylist1,acceptLimit);	
	var similarList2 = similarElements(o2,key2,keylist2,acceptLimit);		

	// Then, assign matches that are above limit
	var match = matchingElements(o1,o2,key1,key2,keylist1,keylist2,acceptLimit,keyrow1,keyrow2);	
	// Make a deep true copy of match as some here make alterations to the headers in the structure
	match = JSON.parse(JSON.stringify(match));
	// Finally, add assign the rest as no matches.		
	var unmatched1 = retrieveElements2(keylist1,keyrow1);
	var unmatched2 = retrieveElements2(keylist2,keyrow2);
	
	// output the result
	outputGrid(match, "matchingId","Matching items","green");
	var name1 = document.getElementById("file-selector1").value.split(/(\\|\/)/g).pop();
	var name2 = document.getElementById("file-selector2").value.split(/(\\|\/)/g).pop();	
	outputGrid(unmatched1,"nonmatching1Id","Non-matching items in "+name1, "red");	
	outputGrid(unmatched2,"nonmatching2Id","Non-matching items in "+name2, "red");	
	outputGrid(similarList1, "duplicates1Id","Possible dubplicates in "+name1,"red");	
	outputGrid(similarList2, "duplicates2Id","Possible dubplicates in "+name2,"red");	
	
	// alther the heading so that columns are associated correctly
	alterAllRecordHeaders(unmatched1,"-1");
	alterAllRecordHeaders(unmatched2,"-2");
	
	// add the distance column so that it is clear that the rows are there in the resulting spreadsheet
	unmatched1 = labelAsUnmatched(unmatched1);
	unmatched2 = labelAsUnmatched(unmatched2);

	// need to merge with dummy records so that sheet is balanced.
	linked = match.concat(unmatched1).concat(unmatched2);
	// return false so that form is not cleared.
	return false;
	}
// alter all headers in the JSON structure	- using?	
function alterAllRecordHeaders(o,postfix)
	{
	for (var line of o)
		{	
		alterRecordHeader(line,postfix);
		}
	}		
// output spreadsheet
function outputSpreadsheet()
	{
	if (typeof linked === 'undefined')	
		{
		return false;
		}		
    var filename='linked-output.xlsx';
	var ws = XLSX.utils.json_to_sheet(linked);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Linked-output");
    XLSX.writeFile(wb,filename);
	return false;
	}
