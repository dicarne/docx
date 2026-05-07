import xml from "xml";
import { Formatter } from "@export/formatter";
import { Paragraph } from "@file/paragraph/paragraph";
import { HeadingLevel } from "@file/paragraph/formatting/style";
import { Media } from "@file/media";
import { toJson } from "./util";
import type { XmlComponent } from "@file/xml-components";

const formatter = new Formatter();
const file = { Media: new Media() } as any;
const context = {
    file,
    viewWrapper: {
        Relationships: {
            addRelationship: () => {}
        }
    } as any,
    stack: [] as never[]
};

const para = new Paragraph({
    text: "Hello World",
    heading: HeadingLevel.HEADING_1
});

const formatted = formatter.format(para as unknown as XmlComponent, context);
const xmlStr = xml(formatted);
console.log("XML:", xmlStr);

// Also check the paragraph-only case (without heading)
const para2 = new Paragraph({
    text: "Hello World",
    style: "MyStyle"
});
const formatted2 = formatter.format(para2 as unknown as XmlComponent, context);
const xmlStr2 = xml(formatted2);
console.log("XML2:", xmlStr2);
