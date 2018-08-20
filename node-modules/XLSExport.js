
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
			        console.log(arrayOfThisRow);
			        var amount = 0;
			        var length = arrayOfThisRow.length;
			        arrayOfThisRow.forEach(function(item, i, arrayOfThisRow) {
		                if (item == '') {                
		                    amount++;
		                }
		            });
			        if (amount != length) {
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
        	location.reload();

	}

	isNew = function(){
	   		 var table = document.querySelector("tbody");
	   		 var countTh = document.getElementsByTagName ('th').length;
	   		 tr = table.insertRow(0);
	   		 for (var i = 0; i < countTh; i++) {
	   		 	td = tr.insertCell(i);
			  	td.appendChild(document.createTextNode(""));
	   		 };
			  
			  console.log(table);

			  $("#editable").editableTableWidget();
             
	}



	list0 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list0')[0].value
	    });
	    location.reload();
	}

	list1 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list1')[0].value
	    });
	    location.reload();
	}

	list2 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list2')[0].value
	    });
	    location.reload();
	}

	list3 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list3')[0].value
	    });
	    location.reload();
	}

	list4 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list4')[0].value
	    });
	    location.reload();
	}

	list5 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list5')[0].value
	    });
	    location.reload();
	}

	list6 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list6')[0].value
	    });
	    location.reload();
	}

	list7 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list7')[0].value
	    });
	    location.reload();
	}

	list8 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list8')[0].value
	    });
	    location.reload();
	}

	list9 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'list_name=' + document.getElementsByName('list9')[0].value
	    });
	    location.reload();
	}

	file0 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'file_name=' + document.getElementsByName('file0')[0].value
	    });
	    location.reload();
	}

	file1 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'file_name=' + document.getElementsByName('file1')[0].value
	    });
	    location.reload();
	}

	file2 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'file_name=' + document.getElementsByName('file2')[0].value
	    });
	    location.reload();
	}

	file3 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'file_name=' + document.getElementsByName('file3')[0].value
	    });
	    location.reload();
	}

	file4 = function(){
		jQuery.ajax({
	        type: 'POST',
	        data: 'file_name=' + document.getElementsByName('file4')[0].value
	    });
	    location.reload();
	}
	

	

});