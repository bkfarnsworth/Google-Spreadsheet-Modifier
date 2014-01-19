/*

This code is to iterate through many google spreadsheets and make the same changes to all of them


*/

function myFunction(){

	var countriesFolder = DocsList.getFolder('Countries');
	var subFolders = countriesFolder.getFolders();//this is a collection of all the countries folders
	var counter = 0;
	var countriesToRun = [
		//"American Indians",
		"Brazil",
		//"Mexico",
		//"Philippines",
		//"Zimbabwe",
		"Peru"
	];


    for (var i in subFolders) {

	    //add this continue statement to skip over certain countries that are done.  
	    if(countriesToRun.indexOf(subFolders[i].getName())==-1){
			continue;
	    }

	    var files = subFolders[i].getFiles();
	    for(var j in files){//for all the spreadsheets in a country folder
			if(files[j].getName().indexOf("Team") == -1){//skip over team lead files.  
				Logger.log(files[j].getName());
				counter++;
				var spreadsheet = SpreadsheetApp.open(files[j]);
				var sheets = spreadsheet.getSheets();
				for(sheet in sheets){

					if(sheets[sheet].getName().indexOf("Week") != -1){//only the week files

					    //these are the cells where the net profit cell could possibly be (the actual number value for net profit, not the word)
					    var startCells = [15,51,42,49,52,41,47];
					    var foundCounter = 0;
					    for (cell in startCells){
							if(foundCounter==2){
						  		continue;
							}

							//this is all stuff for the formulas we wanted to modify this time.  the is net profit, the unknown counter, and a check cell
							//I SHOULD ALWAYS HAVE A CHECK CELL SO THAT i KNOW i AM GETTING THE WRITE ONE.
							var cellStartNum = startCells[cell];
							var npRange = sheets[sheet].getRange("D" + cellStartNum);
							var checkRange = npRange.offset(0,-1);
							var unknownRange = npRange.offset(1,0);
							var npFormula = "=if(AND(isnumber(D" + (cellStartNum-6) + "),isnumber(D" + (cellStartNum-5) + "),isnumber(D" + (cellStartNum-4) + "),isnumber(D" + (cellStartNum-3) + "),isnumber(D" + (cellStartNum-2) + "),isnumber(D" + (cellStartNum-1) + ")),D" + (cellStartNum-4) + "-D" + (cellStartNum-3) + "-D" + (cellStartNum-2) + "-D" + (cellStartNum-1) + ", unknown)";
							var unknownFormula = "=countif(D" + (cellStartNum-6) + ":D" + cellStartNum + ",unknown)";
							var unknownRange2=unknownRange.offset(0,1);
							var unknownFormula2 = "=countif(E" + (cellStartNum-6) + ":E" + cellStartNum + ",unknown)";

							if(checkRange.getValue().toUpperCase() == "Net Profit".toUpperCase() || checkRange.getValue().toUpperCase() == "Lucro Líquido".toUpperCase() || checkRange.getValue().toUpperCase() == "Ganancia Neta".toUpperCase()){
								npRange.setFormula(npFormula);
								unknownRange.setFormula(unknownFormula);
								foundCounter++;//there are two places on each sheet that need these formulas so foundCounter will equal two before the end
								
								//this is a bad way to do this because if they ever change that cell, I will never know if I missed some.  But I will let Matzen tell me about any I missed.
								if(unknownRange2.offset(-8,0).getValue()=="Answers from last week".toUpperCase() || unknownRange2.offset(-9,0).getValue()=="Answers from last week".toUpperCase()){
									unknownRange2.setFormula(unknownFormula2);
								}
							}else if(cell==startCells.length-1 && foundCounter!=2){
								Logger.log("COULD NOT FIND CORRECT CELL: " + files[j].getName() + ", " + sheets[sheet].getName() + ", " + checkRange.getValue());

							}//if Net Profit
				  		}//for cell in possible start cells
					}//if a sheet with week in the name
			  	}//for sheets
			}//if a file that does not have team in it
	  	}//for each file in subfolder
		//break;
		Logger.log("FINISHED COUNTRY: " + subFolders[i].getName());
  	}//for each country folder
	Logger.log("counter: " + counter);
}//function
