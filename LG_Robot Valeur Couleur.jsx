#target Illustrator  
#targetengine main 

var thisFile = new File ($.fileName);
var basePath = thisFile.path;
var csvFile = new File (basePath+"/LG_Couleurs_Recette_Caldera et 2CM7.csv");

var recetteCouleurs = readInCSV(csvFile);
var recetteCouleursLength = recetteCouleurs.length;

function readInCSV(fileObj) {  
	var recetteCouleurs = [];  
	fileObj.open('r');  
	fileObj.seek(0, 0);  
	while(!fileObj.eof) {  
				var thisLine = fileObj.readln();  
				var csvArray = thisLine.split(',');  
				recetteCouleurs.push(csvArray);  
	}  
	fileObj.close();  
	return recetteCouleurs;  
} 

var openDocuments = app.documents;

if (openDocuments.length == 0){
	alert ("Ouvrez le document que vous vouler convertir avant d'activer le script");
} else {
	var openDoc = app.activeDocument;
	var spots = openDoc.spots;
	var spotsLength = spots.length;
	
	var printEnv = myInput();
	var found;
	var results = "";
	
	if (printEnv == "Caldera"){
		for (var i = 0; i < spotsLength; i++){
			found = false;
			var currentSwatch = spots[i];
			if (currentSwatch.name != "[Registration]"){
				for (var j = 0; j<recetteCouleursLength; j++){
//~ 					if (currentSwatch.colorType == "ColorModel.SPOT"){
//~ 						results += "Avertissement: couleurs manquantes - "+currentSwatch.name+"\n";
//~ 						found = true;
//~ 						break;
//~ 					}else{
						if (currentSwatch.name == recetteCouleurs[j][0]){
							if (recetteCouleurs[j][1] == "couleur non trouvée" || recetteCouleurs[j][1] == "" || recetteCouleurs[j][1] == null){
								results += "Avertissement: "+currentSwatch.name+" - aucune valeur Caldera\n";
								found = true;
								break;
							} else {
								var newCMYK = new CMYKColor();
								alert(currentSwatch.name);
								currentSwatch.colorType = ColorModel.PROCESS;
								newCMYK.cyan = parseFloat(recetteCouleurs[j][1]);
								newCMYK.magenta = parseFloat(recetteCouleurs[j][2]);
								newCMYK.yellow = parseFloat(recetteCouleurs[j][3]);
								newCMYK.black = parseFloat(recetteCouleurs[j][4]);
								currentSwatch.color = newCMYK;
								results += "Converti to Caldera: "+currentSwatch.name+"\n";
								found = true;
								break;
							}			
						}
					//}
				}
				if (!found){
					results += "Avertissement: "+currentSwatch.name+" - pas trouvé dans les chartes\n"
				}
			}
		}
		alert (results);
	} else if (printEnv == "2CM7"){
		for (var i = 0; i < spotsLength; i++){
			found = false;
			var currentSwatch = spots[i];
			if (currentSwatch.name != "[Registration]"){
				for (var j = 0; j<recetteCouleursLength; j++){
//~ 					if (currentSwatch.colorType == "ColorModel.SPOT"){
//~ 						results += "Skipped SPOT: "+currentSwatch.name+"\n";
//~ 						found = true;
//~ 						break;
//~ 					}else{
						if (currentSwatch.name == recetteCouleurs[j][0]){
							if (recetteCouleurs[j][5] == "couleur non trouvée" || recetteCouleurs[j][5] == "" || recetteCouleurs[j][5] == null){
								results += "Avertissement: "+currentSwatch.name+" - aucune valeur 2CM7\n";
								found = true;
								break;
							} else {
								currentSwatch.colorType = ColorModel.PROCESS;
								var newCMYK = new CMYKColor();
								newCMYK.cyan = parseFloat(recetteCouleurs[j][5]);
								newCMYK.magenta = parseFloat(recetteCouleurs[j][6]);
								newCMYK.yellow = parseFloat(recetteCouleurs[j][7]);
								newCMYK.black = parseFloat(recetteCouleurs[j][8]);
								currentSwatch.color = newCMYK;
								results += "Converti to 2CM7: "+currentSwatch.name+"\n";
								found = true;
								break;
							}			
						}
//~ 					}
				}
				if (!found){
					results += "Avertissement: "+currentSwatch.name+" - pas trouvé dans les chartes\n"
				}
			}
		}
		alert (results);
	} 
	var userDesktop = Folder.desktop;
	var logFile = new File (userDesktop+"/couleaur_robot.log");
	logFile.open ('w');
	logFile.writeln (results);
	logFile.close();
	logFile.execute();
}




// FUNCTIONS

function myInput(){
		var myWindow = new Window("dialog");
		var myInputGroup = myWindow.add("group");
			myInputGroup.orientation = "column";
			myInputGroup.add ("statictext",undefined, "Impression:");
			var myDropdown = myInputGroup.add("dropdownlist", undefined, ["Caldera","2CM7"]);
			myDropdown.selection = 0;
		var myButtonGroup = myWindow.add("group");
			myButtonGroup.alignment = "center";
			myButtonGroup.add('button', undefined, "Continue", {name: "ok"});
			myButtonGroup.add('button', undefined, "Cancel");
		if (myWindow.show() == 1){
			return myDropdown.selection.text;
		}
}
