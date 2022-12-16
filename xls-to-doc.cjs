// NEM MUKODIK: https://dev.to/iainfreestone/how-to-create-a-word-document-with-javascript-24oi
// MUKODIK:     https://www.npmjs.com/package/docx
// Példák is onnan:		https://runkit.com/dolanmiu/docx-demo2

const docx = require("docx")
const { HeadingLevel, Document, Packer, Paragraph, TextRun, Table, TableCell, TableRow, WidthType, AlignmentType, BorderStyle } = docx;

const xlsx = require("xlsx");
const { writeFile, readFile, set_fs, utils } = xlsx

const fs = require("fs");

filename="ntrmdt-2021-egysejtűek - 2020.11.19. 10.b"
filename="ntrmdt-1920-biokémia - 2020.04.15-17"
filename="ntrmdt-1920-embertan 2020.06.02. 11.b"

marginToSet=400

// Global values and variables
const CellId = {
    QUESTION:0,
    POINT:1,
    QTYPE:2
}

const noBorderObj= {
	top: {
		style: BorderStyle.NONE,
		size: 3,
		color: "FF0000",
	},
	bottom: {
		style: BorderStyle.NONE,
		size: 3,
		color: "FF0000",
	},
	left: {
		style: BorderStyle.NONE,
		size: 3,
		color: "FF0000",
	},
	right: {
		style: BorderStyle.NONE,
		size: 3,
		color: "FF0000",
	}
}

var workbookIn = readFile(filename+".xlsx");
var worksheetIn = workbookIn.Sheets[workbookIn.SheetNames[0]];

var arrIn = utils.sheet_to_json(worksheetIn, {header: 1});

let docxObject={}

// https://docx.js.org/#/usage/styling-with-xml
const styles = fs.readFileSync("./styles.xml", "utf-8");
docxObject.externalStyles=styles

docxObject.sections=[];

let docxSection={}
docxSection.properties={
	page: {
		margin: {
			top: 1000,
			right: marginToSet,
			bottom: 1000,
			left: marginToSet,
		}
	}
}

docxSection.children=[]

let orderNumber=1;

arrIn.forEach(function (exerciese) {


	// QUESTION
	let questionObj={}
	//questionObj.bold=true;
	questionObj.text=(orderNumber++)+"\) "+exerciese[CellId.QUESTION];

	const qParagraphObj=new Paragraph({children: [new TextRun(questionObj)], style: "Question"})

	//https://docx.js.org/api/enums/BorderStyle.html
	const qTableCell = new TableCell({
		children: [qParagraphObj],
		width: {
			size: 80,
			type: WidthType.PERCENTAGE,
		},
		borders: noBorderObj,
	});

	//https://docx.js.org/#/usage/paragraph
	const pParagraphObj=new Paragraph({
		children: [
			new TextRun("_____/"+exerciese[CellId.POINT]+" pont")
		],
		alignment: AlignmentType.RIGHT
	})

	const pTableCell = new TableCell({
		children: [pParagraphObj],
		width: {
			size: 20,
			type: WidthType.PERCENTAGE,
		},
		borders: noBorderObj,
    });

	const tableRow = new TableRow({
		children: [	qTableCell, pTableCell]
	});

	//https://docx.js.org/#/usage/tables?id=table-cells
	const table = new Table({
		rows: [ tableRow ],
		width: {
			size: 100,
			type: WidthType.PERCENTAGE,
		},
		columnWidths: [80, 20]
	});

	docxSection.children.push(table)

	let reserved=Object.keys(CellId).length;

	let qType=exerciese[CellId.QTYPE];

	let tickableSquare={}
	tickableSquare.text='\u25a2';
	tickableSquare.style="tickableStyle";
	tickableSquare.size=60;

	if(qType==='r'){
		// RÖVID VÁLASZ
	} else if(qType==='k'){
		// KIFEJTŐS
	} else if((/^[0-9]+$/.test(qType))||(/^[0-9]+\ [0-9]+/.test(qType))){
		// EGY, vagy TÖBB OPCIÓS
		let paragraphObj={}
		paragraphObj.children=[];
		
		for(let i=reserved;i<exerciese.length;i++){
			paragraphObj.children.push(new TextRun(tickableSquare));
			//https://stackoverflow.com/questions/5237989/how-is-a-non-breaking-space-represented-in-a-javascript-string
			let nonbreakableText = exerciese[i].replace(/ /g, String.fromCharCode(160));
			paragraphObj.children.push(new TextRun(String.fromCharCode(160)+nonbreakableText+" "));
		}

		docxSection.children.push(new Paragraph(paragraphObj))

	} else if(/(I|H)+/.test(qType)){
		// IGAZ-HAMIS
	} else if(qType==='m'){
		// PÁROSÍTÁS
	} else {
		// UNKNOWN
		throw "Unknown or unfilled question type! ->"+qType+"<-"
		//console.log(qType)
	}

});

docxObject.sections.push(docxSection)

let doc = new Document(docxObject);

const paragraph = new docx.Paragraph("Amazing Heading");

Packer.toBuffer(doc).then((buffer) => {
	fs.writeFileSync(filename+".docx", buffer);
});
