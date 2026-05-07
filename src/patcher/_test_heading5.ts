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
import { js2xml } from "xml-js";

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

// What happens when PatchType.PARAGRAPH but children include a Paragraph (wrong usage)?
// This creates textJson with a w:p element.
// Then it's inserted inside the paragraphElement (another w:p)

const mockJson = {
    elements: [{
        type: "element",
        name: "w:document",
        elements: [{
            type: "element",
            name: "w:body",
            elements: [{
                type: "element",
                name: "w:p",
                elements: [{
                    type: "element",
                    name: "w:r",
                    elements: [{
                        type: "element",
                        name: "w:t",
                        elements: [{ type: "text", text: "{{heading_paragraph}}" }]
                    }]
                }]
            }]
        }]
    }]
};

// Let's look at what textJson is for a Paragraph with heading
const formatter = new Formatter();
const para = new Paragraph({
    text: "Blah Blah Blah",
    heading: HeadingLevel.HEADING_1
});
const textJson = [para].map((c) => toJson(xml(formatter.format(c as unknown as XmlComponent, context)))).map((c) => c.elements![0]);
console.log("textJson[0].name:", textJson[0].name);
console.log("textJson[0] elements count:", textJson[0].elements?.length);
console.log("textJson[0] first element:", JSON.stringify(textJson[0].elements?.[0], null, 2));

// Now run the PARAGRAPH replacer with this
const { element: result } = replacer({
    json: mockJson as any,
    patch: {
        type: PatchType.PARAGRAPH,
        children: [para as any]
    },
    patchText: "{{heading_paragraph}}",
    context,
    keepOriginalStyles: false
});

console.log("\nResult JSON:", JSON.stringify(result, null, 2));
