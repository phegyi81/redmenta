import { readFileSync } from 'fs';
import { readdirSync } from 'fs';
import { readdir } from 'fs';
import { writeFileSync } from 'fs';
import * as path from 'path';
import {fileURLToPath} from 'url';
//import * as XLSX from 'xlsx';
import * as fs from "fs";
import { writeFile, readFile, set_fs, utils } from "xlsx/xlsx.mjs";
set_fs(fs);

// Global values and variables

//Finding out working directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
console.log(__dirname)

// For all files in directory
var filesSync = readdirSync(__dirname, []);
filesSync.forEach(file => {
	//console.log(path.parse(file))

	// Filtering for Excel files containing the solutions
	let fileExtension=path.parse(file).ext
	let fileName=path.parse(file).name
	if(fileExtension===".xlsx" && fileName.startsWith("ntrmdt-")){

		var workbookIn = readFile(file);
		var worksheetIn = workbookIn.Sheets[workbookIn.SheetNames[0]];

		console.log(file+" "+workbookIn.SheetNames[0]);

		let orderNumber=1;

		var arrIn = utils.sheet_to_json(worksheetIn, {header: 1});

		var arrOut = [];

		var header = ["", "KÉRDÉS", "VÁLASZ", "TÍPUS", "TÉMAKÖR", "KATEGÓRIA", "FŐKATEGÓRIA", "NYELV", "OPCIÓK", "MEGJEGYZÉS", "PRIVÁT_MEGJEGYZÉS", "MAGYARÁZAT", "PONT", "RÉSZPONTOZÁS", "RÉSZPONTOK", "KÉZI_ÉRTÉKELÉS", "BÜNTETŐPONT", "BÜNTETŐPONTOZÁS", "SEGÍTSÉG", "MEGOLDÁS", "MEGOLDÁS_KÉP", "FORRÁS", "CSOPORTOSÍTÁS", "CÍMKE", "VÁLASZ_ELVÁRÁS", "KÉRDÉS_FORMÁTUM", "KÉP", "MÉDIA_VIDEÓ", "MÉDIA_AUDIÓ", "NEHÉZSÉG", "PARAMÉTEREK", "MEGKÖTÉSEK", "FIX_OPCIÓK", "VÁLASZ_SORREND", "VÁLASZ_REJTETT", "VÁLASZ_CÍMKE", "VÁLASZ_INDEFINIT", "SZINKRON_PARAMÉTEREK", "PONTOSSÁG", "HIBATŰRÉS", "NUMERIKUS_TARTOMÁNY", "IGAZHAMIS_HARMADIK_OPCIÓK", "IGAZHAMIS_HARMADIK_OPCIÓK_FELIRAT", "DÁTUMIDŐ_PONTOSSÁG", "DÁTUMIDŐ_TARTOMÁNY", "SZABADSZÖVEG_KARAKTEREK", "KIFEJEZÉS_BŐVÍTETT", "KIFEJEZÉS_FÜGGVÉNYEK", "KIFEJEZÉS_VÁLTOZÓ", "SEGÍTSÉG_PONTLEVONÁS", "MEGOLDÁS_PONTLEVONÁS"
];

		arrOut.push(header);

		try{
			arrIn.forEach(function (exerciese) {
				let question=exerciese[0];
				let point=exerciese[1];
				let qType=exerciese[2];

				let reserved=3;

				let newExerciese=[];
				newExerciese.push(orderNumber);
				orderNumber++;
				newExerciese.push(question);
				
				if(qType==='r'){
					// RÖVID VÁLASZ
					let answers=" "
					for(let i=reserved;i<exerciese.length;i++){
						if(answers.length>1){
							answers+=" &&& ";
						}
						answers+=exerciese[i];
					}
					newExerciese.push(answers);

					newExerciese.push("SZÖVEGES")

					for(let i=1;i<=4;i++){
						newExerciese.push("");
					}
					newExerciese.push(""); // OPCIOK

				} else if(qType==='k'){
					// KIFEJTŐS
					newExerciese.push("");
					newExerciese.push("SZABAD-SZÖVEGES")
					for(let i=1;i<=4;i++){
						newExerciese.push("");
					}
					newExerciese.push(""); // OPCIOK
				} else if(/^[0-9]+$/.test(qType)){
					// EGY OPCIÓS
					let properAnswerPos = reserved-1 + parseInt(qType);
					let properAnswer=exerciese[properAnswerPos]

					newExerciese.push(properAnswer);
					newExerciese.push("VÁLASZTÁS");
					for(let i=1;i<=4;i++){
						newExerciese.push("");
					}
					
					let wrongAnswers=""
					for(let i=reserved;i<exerciese.length;i++){
						if(i!=properAnswerPos){
							if(wrongAnswers.length>0){
								wrongAnswers+=" &&& ";
							}
							wrongAnswers+=exerciese[i];
						}
					}
					newExerciese.push(wrongAnswers);
					
				} else if(/^[0-9]+\ [0-9]+/.test(qType)){
					// TÖBB OPCIÓS
					let properAnswerPositions=new Set(String(qType).split(" "))
					let properAnswers="";
					properAnswerPositions.forEach(function(pos){
						if(properAnswers.length>0){
							properAnswers+=" &&& ";
						}
						let properAnswerPos = reserved-1 + parseInt(pos);
						properAnswers+=exerciese[properAnswerPos];
					});
					
					newExerciese.push(properAnswers);
					newExerciese.push("TÖBBSZÖRÖS-VÁLASZTÁS");
					for(let i=1;i<=4;i++){
						newExerciese.push("");
					}

					let wrongAnswers=""
					for(let i=reserved;i<exerciese.length;i++){
						let propPosFormat=i-(reserved-1)
						if(!properAnswerPositions.has(propPosFormat.toString())){
							if(wrongAnswers.length>0){
								wrongAnswers+=" &&& ";
							}
							wrongAnswers+=exerciese[i];
						}
					}
					newExerciese.push(wrongAnswers);
				} else if(/(I|H)+/.test(qType)){
					// IGAZ-HAMIS
					let truePositions=new Set();
					let falsePositions=new Set();
					let qTypeUpper=qType.toUpperCase();
					for(let i=0;i<=qTypeUpper.length-1;i++){
						switch(qTypeUpper.charAt(i)){
							case 'I':
								truePositions.add(i+reserved);
								break;
							case 'H':
								falsePositions.add(i+reserved);
								break;
							default:
								throw "Not only I or H in true-false question type ->"+qType+"<-"
						}
					}
					
					let properAnswers="";
					truePositions.forEach(function(pos){
						if(properAnswers.length>0){
							properAnswers+=" &&& ";
						}
						let properAnswerPos = parseInt(pos);
						properAnswers+=exerciese[properAnswerPos];
					});

					newExerciese.push(properAnswers);
					newExerciese.push("IGAZ/HAMIS");

					for(let i=1;i<=4;i++){
						newExerciese.push("");
					}

					let wrongAnswers="";
					falsePositions.forEach(function(pos){
						if(wrongAnswers.length>0){
							wrongAnswers+=" &&& ";
						}
						let wrongAnswerPos = parseInt(pos);
						wrongAnswers+=exerciese[wrongAnswerPos];
					});

					newExerciese.push(wrongAnswers);
					
				} else if(qType==='m'){
					// PÁROSÍTÁS

					newExerciese.push(""); // VALASZ
					newExerciese.push("PÁROSÍTÓS");
					for(let i=1;i<=4;i++){
						newExerciese.push("");
					}
					newExerciese.push(""); // OPCIOK

					console.log("Matching - NOT SUPPORTED by EDUBASE!!!");
				} else {
					// UNKNOWN
					throw "Unknown or unfilled question type! ->"+qType+"<-"
					//console.log(qType)
				}

				for(let i=1;i<=3;i++){
					newExerciese.push("");
				}
				//POINTS
				newExerciese.push(point);

				arrOut.push(newExerciese);
			});

			var workbookOut = utils.book_new();
			var worksheetOut = utils.aoa_to_sheet(arrOut);
			utils.book_append_sheet(workbookOut, worksheetOut, "Feladatlap");
			writeFile(workbookOut, "final-"+fileName.substring(7)+".xlsx");

		} catch(e) {
			console.log(e)
		}
		arrOut = [];
	}
});
