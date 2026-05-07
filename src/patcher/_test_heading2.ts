import xml from "xml";
import { Formatter } from "@export/formatter";
import { Paragraph } from "@file/paragraph/paragraph";
import { HeadingLevel } from "@file/paragraph/formatting/style";
import { Media } from "@file/media";
import { toJson } from "./util";
import type { XmlComponent } from "@file/xml-components";
import { replacer } from "./replacer";
import { PatchType } from "./from-docx";
import type { File } from "@file/file";
import type { IViewWrapper } from "@file/document-wrapper";

const formatter = new Formatter();
const file = { Media: new Media() } as unknown as File;
const context = {
    file,
    viewWrapper: {
        Relationships: {
            addRelationship: () => {}
        }
    } as unknown as IViewWrapper,
    stack: [] as never[]
};

// Simulate the replacer behavior for a DOCUMENT type patch with heading
const para = new Paragraph({
    text: "Hello World",
    heading: HeadingLevel.HEADING_1
});

const formatted = formatter.format(para as unknown as XmlComponent, context);
const xmlStr = xml(formatted);
console.log("Raw XML from formatter:", xmlStr);

const parsed = toJson(xmlStr);
console.log("Parsed elements count:", parsed.elements?.length);
console.log("First element name:", parsed.elements?.[0]?.name);
console.log("First element:", JSON.stringify(parsed.elements?.[0], null, 2));

// Now simulate what replacer.ts does
const textJson = [para].map((c) => toJson(xml(formatter.format(c as unknown as XmlComponent, context)))).map((c) => c.elements![0]);
console.log("\ntextJson[0]:", JSON.stringify(textJson[0], null, 2));
