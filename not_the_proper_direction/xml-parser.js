import { XMLValidator } from 'fast-xml-parser';
import { XMLParser } from 'fast-xml-parser';
import { readFileSync } from 'fs';

const xmlFileInPlace = `<?xml version="1.0"?>
<catalog>
   <book id="bk101">
      <author>Gambardella, Matthew</author>
      <title>XML Developer's Guide</title>
      <genre>Computer</genre>
      <price>44.95</price>
      <publish_date>2000-10-01</publish_date>
      <description>An in-depth look at creating applications 
      with XML.</description>
   </book>   
</catalog>`;

const xmlFile = readFileSync(`${process.cwd()}/Redmenta - Feladatlap szerkesztése-Biokemia1.html`, 'utf8');
const parser = new XMLParser();
const json = parser.parse(xmlFile);

const result = XMLValidator.validate(xmlFile);
if (result === true) {
  console.log(`XML file is valid`, result);
}

if (result.err) {
  console.log(`XML is invalid becuause of - ${result.err.msg}`, result);
}
