import JSZip from "jszip";
import { Paragraph, HeadingLevel, TextRun } from "@file/paragraph";
import { patchDocument, PatchType } from "./from-docx";
import { js2xml, xml2js } from "xml-js";

// Create a mock template with a styles.xml and document.xml
async function createMockTemplate(): Promise<JSZip> {
    const zip = new JSZip();
    
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>{{heading_paragraph}}</w:t>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:style w:type="paragraph" w:styleId="Normal">
        <w:name w:val="Normal"/>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:rPr>
            <w:b/>
            <w:sz w:val="32"/>
        </w:rPr>
    </w:style>
</w:styles>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
</Types>`;

    zip.file("word/document.xml", docXml);
    zip.file("word/styles.xml", stylesXml);
    zip.file("[Content_Types].xml", contentTypesXml);
    
    return zip;
}

async function main() {
    const zip = await createMockTemplate();
    
    const result = await patchDocument({
        outputType: "uint8array",
        data: zip,
        patches: {
            // eslint-disable-next-line @typescript-eslint/naming-convention
            heading_paragraph: {
                type: PatchType.DOCUMENT,
                children: [
                    new Paragraph({
                        text: "Blah Blah Blah",
                        heading: HeadingLevel.HEADING_1
                    })
                ]
            }
        }
    });
    
    // Read the output zip
    const outputZip = await JSZip.loadAsync(result);
    const docContent = await outputZip.file("word/document.xml")!.async("text");
    
    console.log("Output document.xml:");
    console.log(docContent);
    console.log("\nContains pStyle Heading1?", docContent.includes("Heading1"));
    console.log("Contains pStyle?", docContent.includes("pStyle"));
    
    const stylesContent = await outputZip.file("word/styles.xml")!.async("text");
    console.log("\nOutput styles.xml:");
    console.log(stylesContent);
    console.log("\nStyles contains Heading1?", stylesContent.includes("Heading1"));
}

main().catch(console.error);
