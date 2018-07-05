
require(["xlsx.full.min.js"], function () {

	 isSend = function(){

	   		/* original data */
	   		var XLSX = {};
	   
	   		make_xlsx_lib(XLSX);
			//var ws_name = "SheetJS"; //Имя листа задаётся в другом файле
			var tableHTML = editable.outerHTML;


			var myTableArray = [];

			$("#editable tr").each(function() {
			    var arrayOfThisRow = [];
			    var tableDataHeader = $(this).find('th');

			    if (tableDataHeader.length > 0) {
			        tableDataHeader.each(function() {
			        	if ($(this).text() != 'undefined') {
			        		arrayOfThisRow.push($(this).text().trim()); 
			        	} else {
				    	tableDataHeader.each(function() {arrayOfThisRow.push('').trim(); });
				    	}
			        });
			        myTableArray.push(arrayOfThisRow);
			    }

			    var tableData = $(this).find('td');
			    if (tableData.length > 0) {		        
			        tableData.each(function() {
			        	if ($(this).text() != 'undefined') {
			        		arrayOfThisRow.push($(this).text().trim()); 
			        	} else {
				    	tableData.each(function() {arrayOfThisRow.push(''); });
				    	}
			        });
			        if (JSON.stringify(arrayOfThisRow) != JSON.stringify(['','','',''])) {
			      		myTableArray.push(arrayOfThisRow);
			    	} 
			    };
			});


			var ws = XLSX.utils.aoa_to_sheet(myTableArray);
			//var wb = XLSX.utils.book_new();
			/* add worksheet to workbook */
			//XLSX.utils.book_append_sheet(wb, ws, ws_name);

			jQuery.ajax({
            type: 'POST',
            data: 'arrObjects=' + JSON.stringify(ws)
        	});

	}

});


