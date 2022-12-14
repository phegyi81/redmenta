// NEM MUKODIK: https://dev.to/iainfreestone/how-to-create-a-word-document-with-javascript-24oi
// MUKODIK:     https://www.npmjs.com/package/docx
// Példák is onnan:		https://runkit.com/dolanmiu/docx-demo2

const docx = require("docx")
const { HeadingLevel, Document, Packer, Paragraph, TextRun } = docx;

const xlsx = require("xlsx");
const { writeFile, readFile, set_fs, utils } = xlsx

const fs = require("fs");

filename="ntrmdt-2021-egysejtűek - 2020.11.19. 10.b"

var workbookIn = readFile(filename+".xlsx");
var worksheetIn = workbookIn.Sheets[workbookIn.SheetNames[0]];

let orderNumber=1;

var arrIn = utils.sheet_to_json(worksheetIn, {header: 1});


let heading2style={}
heading2style.id="Heading1"
heading2style.name="Heading 1"
heading2style.basedOn= "Normal"
heading2style.next= "Normal"
heading2style.quickFormat= true;
heading2style.run={}
heading2style.run.size=28;
heading2style.run.bold=true;
heading2style.run.italics=true;
heading2style.run.color="ff0000";
heading2style.paragraph={}
heading2style.paragraph.spacing={}
heading2style.paragraph.spacing.before=120;

let docxObject={}
docxObject.styles={}
docxObject.styles.paragraphStyles=[heading2style]

docxObject.sections=[];

let docxSection={}
docxSection.properties={}
docxSection.children=[]

arrIn.forEach(function (exerciese) {

	let paragraphObj={}

	let questionObj={}
	//questionObj.bold=true;
	questionObj.text="Kérdés: "+exerciese[0]

	paragraphObj.children=[new TextRun(questionObj)]
	paragraphObj.heading=HeadingLevel.HEADING_2
	docxSection.children.push(new Paragraph(paragraphObj))

	paragraphObj.heading=HeadingLevel.NORMAL

	paragraphObj.children=[new TextRun("Elérhető pontok száma:\t"+exerciese[1])]
	docxSection.children.push(new Paragraph(paragraphObj))

	paragraphObj.children=[new TextRun("Kérdés típusa:\t"+exerciese[2])]
	docxSection.children.push(new Paragraph(paragraphObj))

});

docxObject.sections.push(docxSection)

let doc = new Document(docxObject);

//////////////

//let doc = new Document({
//        sections: [{
//            properties: {},
//            children: [
//                new Paragraph({
//                    children: [
//                        new TextRun("Hello World"),
//                        new TextRun({
//                            text: "Foo Bar",
//                            bold: true,
//                        }),
//                        new TextRun({
//                            text: "\tGithub is the best",
//                            bold: true,
//                        }),
//                    ],
//                }),
//            ],
//        }],
//    })

Packer.toBuffer(doc).then((buffer) => {
	fs.writeFileSync(filename+".docx", buffer);
});
