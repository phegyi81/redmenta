import parse from 'html-dom-parser';
import { readFileSync } from 'fs';
import { readdirSync } from 'fs';
import { readdir } from 'fs';
import { writeFileSync } from 'fs';
import * as path from 'path';
import {fileURLToPath} from 'url';
import * as XLSX from 'xlsx';

// Global values and variables
let exerciese=""
let actualQuestion=[]
let sheetOfQuestions=[]
let pointsToGet=0
let csvNeeded=0
let consoleNeeded=0
let q=1;
let o=1;
let classOfAdditionalSection="float-left clear-both mb-1.5 uppercase text-gray-dark text-xs"
let classOfOptions="text-dark bg-gray-lightest border-gray-lightest hover:bg-gray-light hover:border-gray-light py-1.25 px-4 mb-4 mr-4 text-left font-bold leading-snug border-4 rounded-5 max-w-full break-words print:border-none print:flex print:items-center print:mb-1";
let classOfTrueOrFalse="w-full flex justify-center text-center break-words"
let classOfMatching="p-3.5 flex justify-center flex-1 w-full break-words"
let classOfPoints="text-cream text-2xl font-bold w-16 h-8 text-center"

//Finding out working directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
//console.log(__dirname)

// For all files in directory
var filesSync = readdirSync(__dirname, []);
filesSync.forEach(file => {
	//console.log(path.parse(file))

	// Filtering for HTML extension
	let fileExtension=path.parse(file).ext
	if(fileExtension===".html"){

		let fileToOpen=file;
		let outputBaseName="ntrmdt-"+path.parse(fileToOpen).name; //intermediate

		console.log(fileToOpen+" -> "+outputBaseName);

		// Re-creating output
		if(csvNeeded==1){
			writeFileSync(outputBaseName+".csv", "", err => {
				if (err) {
					console.error(err);
				}
			});
		}

		// Reading input
		const htmlFile = readFileSync(`${process.cwd()}/`+fileToOpen, 'utf8');

		// Parsing input to java object
		const javaObject = parse(htmlFile);

		//Processing input
		searchOf(javaObject[4].children[2], "", outputBaseName);
		
	}
});

function output(textToEmit, outputBaseName){
	if(csvNeeded==1){
		writeFileSync(outputBaseName+".csv", textToEmit+"\n", { flag: 'a+' }, err => {
			if (err) {
				console.error(err);
			}
		});
	}
	if(consoleNeeded==1){
		console.log(textToEmit);
	}
}

function saveActualQuestion(){
	sheetOfQuestions.push(actualQuestion);
	actualQuestion=[];
}

function searchOf(json_object, path, outputBaseName){
	// Iterate through every children
	for (let i in json_object.children) {
		// Does it have attribs?
		if (typeof json_object.children[i].attribs !== 'undefined'){
			// Is the class equal to break-words?
			if(json_object.children[i].attribs.class==="break-words"){
				// Then it is the question!
				output(exerciese, outputBaseName);
				if(actualQuestion.length>0){
					saveActualQuestion();
				}
				exerciese=json_object.children[i].children[0].data+";"+pointsToGet+";;";
				//console.log(q+". "+json_object.children[i].children[0].data);
				actualQuestion.push(json_object.children[i].children[0].data);
				actualQuestion.push(pointsToGet);
				actualQuestion.push("");
				q++
			// Is the class the one belonging to the options?
			} else if(json_object.children[i].attribs.class===classOfOptions){
				// Then it is an option!
				exerciese=exerciese+json_object.children[i].children[1].data+";"

				// saving exerciese
				actualQuestion.push(json_object.children[i].children[1].data);
				//console.log(" -"+json_object.children[i].children[1].data);
			// Is the class the one belonging to the subtitle of adding more questions?
			} else if(json_object.children[i].attribs.class===classOfTrueOrFalse){
				// Then it is a true or false question!
				exerciese=exerciese+json_object.children[i].children[0].data+";"

				// saving exerciese
				actualQuestion.push(json_object.children[i].children[0].data);

				//console.log(" -"+json_object.children[i].children[0].data);

			} else if(json_object.children[i].attribs.class===classOfMatching){
				// Then it is a matching question!
				exerciese=exerciese+json_object.children[i].children[0].data+";"

				// saving exerciese
				actualQuestion.push(json_object.children[i].children[0].data);

				//console.log(" -"+json_object.children[i].children[0].data);
			} else if(json_object.children[i].attribs.class===classOfAdditionalSection){
				// Then close this file
				output(exerciese, outputBaseName);

				// saving last exerciese
				saveActualQuestion();
				
				// name of Excel
				let outputBaseNameXLS=outputBaseName+".xlsx"
				
				//console.log(sheetOfQuestions);
				var workbook = XLSX.utils.book_new();
				var worksheet = XLSX.utils.aoa_to_sheet(sheetOfQuestions);
				XLSX.utils.book_append_sheet(workbook, worksheet, "Feladatlap");
				XLSX.writeFile(workbook, outputBaseNameXLS);

				sheetOfQuestions = [];
				return 1;
			} else if(json_object.children[i].attribs.class===classOfPoints){
				pointsToGet=json_object.children[i].children[0].data;
				//console.log(pointsToGet)
			}
		}
		if (typeof json_object.children[i].children !== 'undefined'){
				//console.log(path+json_object.children[i].name+": "+json_object.children[i].children.length);
			if(json_object.children[i].children.length>0){
				let completed=searchOf(json_object.children[i], path+json_object.children[i].name+"/", outputBaseName)
				if(completed==1){
					return completed;
				}
			}
		} else {
			//console.log(json_object.children[i].type)
			//console.log(json_object.children[i].data)
		}
	}
	return 0;
}
