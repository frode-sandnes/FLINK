// By Frode Eika Sandnes, Oslo Metropolitan University, Oslo, Norway, May 2021

"use strict"

// global variables
let file1, file2;
let linked;
let combinedFilename = ""; // prefix of te new filename based on a concatenation of the two sources

// update GUI when slider moves
function acceptChange()
	{
	const value = document.getElementById("acceptLimitId").value;
	document.getElementsByName("acceptLabel")[0].textContent = "Accept limit ("+value+")";	
	}
// string distance measure using DICE
function dice(str1, str2)
	{
	if (str1.length < 2 || str2.length < 2) return 0;
	const charArr1 = [...str1.toLowerCase()], charArr2 = [...str2.toLowerCase()];
	const bigrams1 = charArr1.filter((e, i) => i < charArr1.length - 1)
						 .map((e, i) => e + charArr1[i + 1]);						
	const bigrams2 = charArr2.filter((e, i) => i < charArr2.length - 1)
						 .map((e, i) => e + charArr2[i + 1]);
	const intersection = new Set(bigrams1.filter(e => bigrams2.includes(e)));
	// count number of intersecting bigrams
	const intersectionCounts = [...intersection].map(bigram => Math.min(bigrams1.filter(e => e == bigram).length,
							 								            bigrams2.filter(e => e == bigram).length));
	const intersectionSize = intersectionCounts.reduce((accumulator, e) => accumulator + e, 0);
	// simply compare intersection to the shortest string only, not both
	// based on the assumption there that the longest string is most correct and complete
	let minLength = Math.min(str1.length, str2.length);
	// short names give very few bigrams, therefore need to adjust minLength in such cases- subdcract ibrams involving space to compensate for this	
	const [smallest, largest] = (str1.length > str2.length)? [str2, str1]: [str1, str2];
	if (!smallest.includes(" "))	// smallest does not include space
		{
		minLength -= largest.split(" ").length -1; // subtract false bigram counts due to likely missing space
		}	
	return intersectionSize/minLength;
	}
// Bootstrapping
window.addEventListener('DOMContentLoaded', (event) => setup());
function setup()
    {
    // Add the two file load handlers
	const selector1 = "1", selector2 = "2";
	const fileSelector1 = document.getElementById("file-selector" + selector1);
	fileSelector1.addEventListener('change', (event) => loadSpreadSheet(event, selector1));
	const fileSelector2 = document.getElementById("file-selector" + selector2);
	fileSelector2.addEventListener('change', (event) => loadSpreadSheet(event, selector2));
    }
// retrieving file contents in excel format	
function loadSpreadSheet(event,selector)
	{
	const files = event.target.files;

	for (var i = 0, f; f = files[i]; i++) 
		{			
		let {name} = f;		// extract current filename
		combinedFilename += name.substring(0, name.lastIndexOf(".")) + "-"; // adding the filename to the new filename

		let reader = new FileReader();
		
		reader.onload = (function(theFile) 
			{
			return function(e) 
				{
				let workbook = XLSX.read(e.target.result, {type: 'binary'});	
				if (workbook.SheetNames.length > 1)
					{
					report("WARNING: More than one sheets in the workbook(" + workbook.SheetNames+ ") for "+name+", selecting the first one ("+workbook.SheetNames[0]+")");
					}
				const sheetName = workbook.SheetNames[0];	// selecting the first one 
				let XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
				let json_object = JSON.stringify(XL_row_object);
	//			json_object = json_object.normalize();
				if (selector === "1")
					{
					file1 = json_object;	
					}
				else
					{
					file2 = json_object;								
					}	
				outputGrid(XL_row_object, "table" + selector , "", 'gray', true,3);
				mergeOnFirstColumns();						
				};
			})(f);		
		reader.readAsBinaryString(f);
		}
	}

// create html element and attach to parent
function createAndAttachHTML(parent, htmlTag)
	{
	let htmlElement = document.createElement(htmlTag);
	parent.appendChild(htmlElement);
	return htmlElement;	
	}
// output JSON in table format
function outputGrid(sheet, id, heading, colour, checkboxes = false, limit = 1000)
	{
	let root = document.getElementById(id);  // the root element of this table
	root.innerHTML = ""; //  clear the html element in case it is a subsequent function call			
	if (sheet.length === 0) // only proceed if we have elements
		{
		return;
		}
	// make the text lighter as it is just for verification and not central for understanding the information.
	root.style.color = colour;										
	// descriptive table heading for the table
	let h2 = createAndAttachHTML(root, "h2");
	h2.innerText = heading;
	// create the table and make it small
	let small = createAndAttachHTML(root, "small");
	let table = createAndAttachHTML(small, "table");
	// first build a list of header names based on the first row
	let tr = createAndAttachHTML(table, "tr");
	let headers = getHeaders(sheet);
	headers.forEach((head, i) =>
		{
		let td = createAndAttachHTML(tr, "td");
		let checkboxid = id + head;
		if (checkboxes)	// if we are to associate checkbox with header?
			{
			let input = createAndAttachHTML(td, "input");
			input.type = "checkbox";
			input.id = checkboxid;
			input.name = checkboxid;
			input.checked = (i == 0); // set the first checkbox by default as checked
			let label = createAndAttachHTML(td, "label");
			label.for = checkboxid;
			label.innerText = head;
			}	
		else
			{
			td.innerText = head;		// insert ordinary text header
			}
		});
	// then traverse each row and recall the value by looping through the keys (accessed via bracket notation).
	sheet.filter((line, count) => count < limit)	// just include the first ones
		.forEach((line, count) => 
			{
			let tr = createAndAttachHTML(table, "tr");
			// if the similarity is close to limit chance the color¨to catch the user's attention, if this is table without the similarity attribute, the code will not trigger as "undefined"
			let {similarity} = line;
			const acceptLimit = document.getElementById("acceptLimitId").value/100;
			if (similarity - acceptLimit < 1.0 - similarity)
				{
				tr.classList.add("lowSimilarityWarning"); // set style for the entire table row.
				}
			headers.forEach(head => 
				{
				let td = createAndAttachHTML(tr, "td");
				// get string from sheet
				let str = line[head];
				if (typeof(str) != "undefined")
					{
					td.innerText = str;
					}				
				})
			});	
	}
// get table headers -- assumming the first row contains all the elements
function getHeaders(sheet)
	{
	return [...Object.keys(sheet[0])];
	}
// get all the values in cells for a given set of columns with checkboxes set	
function getColumns(sheet, id, headers)
	{
	// identify checked headers
	const checkedIds = headers.map((head, i) => ({checkboxID: id + head, header:head}))
							  .filter(({checkboxID}) => document.getElementById(checkboxID).checked);
	// combine key parts for row to form the concatednated key
	const combinedKeylist = sheet.map(row => checkedIds.map(({header}) => 
							((typeof row[header] !== 'undefined')  // check that it actually exist
								? row[header]		// if the checkbox is set, the value in the cell is added to the key
								: "")).join(" "));	// add space separator
	// create the rowlist 
	const rowlist = sheet.map(row => row);
	return {keys: combinedKeylist, rows: rowlist};
	}		
// remove a set of items in one list from another	
function subtractElements(list1, list2)
	{
	list2.filter(v => list1.includes(v)) 	// intersection of two lists
		 .map(v => list1.indexOf(v))	 	// indexes of elements
		 .reverse()							// remove high indices first
		 .forEach(i => list1.splice(i, 1)); // removing the element without leaving a hole
	}		
// find a given row in the json spreadsheet	
function getRecord2(keyrow, key)
	{
	const index = keyrow.keys.indexOf(key);
	return keyrow.rows[index];
	}
// rename the key of a simple object	
function renameKey(sheet, oldKey, newKey) 
	{   
    Object.defineProperty(sheet, newKey, Object.getOwnPropertyDescriptor(sheet, oldKey));
    delete sheet[oldKey];                              
    }	
// add suffix to all header elements to make these unique	
function alterRecordHeader(r, suffix)
	{
	[...Object.keys(r)].forEach(head => renameKey(r, head, head+suffix));
	}
// find set of elements in two lists that are above the threshold in similarity	
// more robust approach
function matchingElements(key1, key2, keylist1, keylist2, acceptLimit, keyrow1, keyrow2)
	{
	// create list of distances above acceptLimit
	let topDistances = keylist1.flatMap((value1,i) =>
		{
		return keylist2.reduce((accumulator,value2,j) => 
			{
			let distance = dice(value1, value2);
			if (isNaN(distance) || !isFinite(distance))
				{
				return accumulator;	
				}
			if (distance < acceptLimit)
				{
				return accumulator;	
				}
			return [...accumulator, {distance: distance, i: i, j:j}];
			}, [])
		});
		// sort list of distances in decending order
	topDistances.sort((a, b) => b.distance - a.distance);
	// go through list, pick pairs that are not already selected
	// if selected mark as selected
	let selectedList1 = {};
	let selectedList2 = {};
	let matches = [];
	topDistances.forEach(entry => 
		{
		let {i, j} = entry;
		if (i in selectedList1)
			{
			return;	
			}
		if (j in selectedList2)
			{
			return;	
			}
		matches.push(entry);
		selectedList1 = {...selectedList1, [i]:i};
		selectedList2 = {...selectedList2, [j]:j};
		});
	let selstr1 = keylist1.filter((value, i) => (i in selectedList1));
	let selstr2 = keylist2.filter((value, j) => (j in selectedList2));
let matstr = matches.map(({distance,i,j}) => ({distance:distance, i:keylist1[i], j:keylist2[j]}));

		// find unselected based on those not selected
	let unselectedList1 = keylist1.filter((value, i) => !(i in selectedList1));
	let unselectedList2 = keylist2.filter((value, j) => !(j in selectedList2));
	// can return compund object with all the results	
//	console.log("acceptLimit",acceptLimit);
//	console.log("keylist1",keylist1)	
//	console.log("matches",matches);
//	console.log("matches",matstr);
//	console.log("selectedList1",selectedList1);
//	console.log("selectedList2",selectedList2);
//	console.log("selectedList1",selstr1);
//	console.log("selectedList2",selstr2);
//	console.log("unselectedList1",unselectedList1);
//	console.log("unselectedList2",unselectedList2);
	// prepare the right format
	let match = matches.map(({distance, i, j}) => 
		{
		let r1 = getRecord2(keyrow1, keylist1[i]);	
		alterRecordHeader(r1, "-1");
		let r2 = getRecord2(keyrow2, keylist2[j]);			
		alterRecordHeader(r2, "-2");
		let r3 = {similarity: distance.toFixed(2), [key1]: keylist1[i], [key2]: keylist2[j]};
		let r = {...r3, ...r1, ...r2};		
		return r;			
		});
//	console.log("unselectedList1",unselectedList1);
	return {match: match, unmatched1: unselectedList1, unmatched2: unselectedList2};
	}
// find set of elements in two lists that are above the threshold in similarity	
function 
matchingElements2(key1, key2, keylist1, keylist2, acceptLimit, keyrow1, keyrow2)
	{
/*console.log(key1)		
console.log(key2)
console.log(keylist1)		
console.log(keylist2)		
console.log(keyrow1)		
console.log(keyrow2)		
console.log(acceptLimit)	*/	

	// fore each item in keylist1 - find the item with highest match in keylist2
	const matching = keylist1.map(value1 => 
								{		// calc stuff we need
								let distances = keylist2.map(value2 => dice(value1, value2));
								let max = Math.max(...distances);			
								return ({value: value1, distances: distances, max: max});
								})
							.filter(({max}) => max >= acceptLimit)	// only continue if sufficiently high match
							.map(value1 => // calc some more stuff we need
								{ 							// check backwards from keylist2 to keylist1 if there are better alternatives
								let maxItem = keylist2[value1.distances.indexOf(value1.max)];
								let backDistances = keylist1.map(checkValue =>  dice(checkValue, maxItem))
															.filter(distance => distance > value1.max);
							    return { ...value1, maxItem: maxItem, backDistances: backDistances};
								})
							.filter(({backDistances}) => backDistances.length == 0) 	// only continue if there are no better alternatives
							.map(value1 => 	// prepare element
								{
								let r1 = getRecord2(keyrow1, value1.value);	
								alterRecordHeader(r1, "-1");
								let r2 = getRecord2(keyrow2, value1.maxItem);			
								alterRecordHeader(r2, "-2");
								let r3 = {similarity: value1.max.toFixed(2), [key1]: value1.value, [key2]: value1.maxItem};
								let r = {...r3, ...r1, ...r2};
								//  remove items that have been selected
								keylist2.splice(keylist2.indexOf(value1.maxItem), 1);
								return r;					
								});
	//  remove items that have been selected
	const selected = matching.map(row => row[key1+"-1"]);
//console.log("selected",selected);	
//console.log("before",keylist1)	
// here does not work - need rework
	subtractElements(keylist1, selected);
//console.log("after",keylist1)	
	return matching;
	}
// return elements in the same list that are above the threshold of similarity	
function similarElements(key, keylist, acceptLimit)
	{
	return keylist.flatMap((value1, i) =>  
		keylist.filter((v, j) => j > i)
			      .filter(value2 => dice(value1, value2) >= acceptLimit)
			      .map(value2 => ({similarity: dice(value1, value2).toFixed(2), [key+"1"]: value1, [key+"2"]: value2})));
	}	
// find records for keylist
function retrieveElements2(keylist, keyrow)
	{
//console.log("variables: ", keylist,keyrow);
//console.log("sjekk: ",keylist.map(key => getRecord2(keyrow, key)));
return keylist.map(key => getRecord2(keyrow, key));
	}	
// attach a distance undefined column so that users of spreadsheet sees the full range of records	
function labelAsUnmatched(unmatched)
	{
	return unmatched.map(row => ({similarity: "no-match", ...row}));
	}
// general status messages for the GUI
function report(message, id = "statusMessageId")
	{
	document.getElementById(id).innerText += "\n\n" + message;
	}
// check for duplicates and output warning
function checkDuplicates(keylist, id)
	{
	let warning = "";
    // create a Set with array elements
    let s = new Set(keylist);
    // compare the size of array and Set
    if(keylist.length !== s.size)
		{
        report("Warning: Duplicate entries in File.", id);	
		}
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
	const sheet1 = JSON.parse(file1);
	const sheet2 = JSON.parse(file2);
	const acceptLimit = document.getElementById("acceptLimitId").value/100;
	// get header for first rows
	const headers1 = getHeaders(sheet1);
	const headers2 = getHeaders(sheet2);
	const key1 = headers1[0];	// use first header item for comparison
	const key2 = headers2[0];	// use first header item for comparison	
	// create keylist for files 1 and 2 vased on respective columns in the sheet
	const keyrow1 = getColumns(sheet1, "table1", headers1);
	const keyrow2 = getColumns(sheet2, "table2", headers2);
	const keylist1 = [...keyrow1.keys];	// shallow true copy of arrays
	const keylist2 = [...keyrow2.keys];
	// check for duplicated entries, if duplicates - output warning
	checkDuplicates(keylist1, "messageFile1");
	checkDuplicates(keylist2, "messageFile2");
	// First, find cases in each file that are very similar
	const similarList1 = similarElements(key1, keylist1, acceptLimit);	
	const similarList2 = similarElements(key2, keylist2, acceptLimit);		
	// Then, assign matches that are above limit
//	let match = [], unmatched1 = [], unmathced2 = [];	
	let {match, unmatched1, unmatched2} = matchingElements(key1, key2, keylist1, keylist2, acceptLimit, keyrow1, keyrow2);	
	// create table format for output
//console.log("sheet1",sheet1);	
//console.log("unmatched1",unmatched1);	
//	unmatched1 = sheet1.filter(row => unmatched1.includes(row[keyrow1]));
//	unmatched2 = sheet2.filter(row => unmatched2.includes(row[keyrow2]));
	unmatched1 = retrieveElements2(unmatched1, keyrow1);
	unmatched2 = retrieveElements2(unmatched2, keyrow2);	
//	let match = matchingElements(key1, key2, keylist1, keylist2, acceptLimit, keyrow1, keyrow2);	
	// Make a deep true copy of match as some here make alterations to the headers in the structure
	match = JSON.parse(JSON.stringify(match));
	// Finally, add assign the rest as no matches.		
//	let unmatched1 = retrieveElements2(keylist1, keyrow1);
//	let unmatched2 = retrieveElements2(keylist2, keyrow2);
	// output the result
	outputGrid(match, "matchingId", "Matching items", "lime");
	const name1 = document.getElementById("file-selector1").value.split(/(\\|\/)/g).pop();
	const name2 = document.getElementById("file-selector2").value.split(/(\\|\/)/g).pop();	
	outputGrid(unmatched1, "nonmatching1Id", "Non-matching items in " + name1, "orangered");	
	outputGrid(unmatched2, "nonmatching2Id", "Non-matching items in " + name2, "orangered");	
	outputGrid(similarList1, "duplicates1Id", "Possible dubplicates in " + name1, "orangered");	
	outputGrid(similarList2, "duplicates2Id", "Possible dubplicates in " + name2, "orangered");	
//console.log("matching",match);
//console.log("unmatched1",unmatched1);
//console.log("unmatched2",unmatched2);
//console.log("similarList1",similarList1);
//console.log("similarList2",similarList2);
	// alther the heading so that columns are associated correctly
	alterAllRecordHeaders(unmatched1, "-1");
	alterAllRecordHeaders(unmatched2, "-2");
	// add the distance column so that it is clear that the rows are there in the resulting spreadsheet
	unmatched1 = labelAsUnmatched(unmatched1);
	unmatched2 = labelAsUnmatched(unmatched2);
	// need to merge with dummy records so that sheet is balanced.
	linked = match.concat(unmatched1).concat(unmatched2);
	// return false so that form is not cleared.
	return false;
	}
// alter all headers in the JSON structure	- using?	
function alterAllRecordHeaders(sheet, postfix)
	{
	sheet.forEach(line => alterRecordHeader(line, postfix));
	}		
// output spreadsheet
function outputSpreadsheet()
	{
	if (typeof linked === 'undefined')	
		{
		return false;
		}		
    const filename = combinedFilename + 'linked.xlsx';
	const ws = XLSX.utils.json_to_sheet(linked);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Linked-output");
    XLSX.writeFile(wb,filename);
	report("Linked sheet saved to " + filename);	
	return false;
	}