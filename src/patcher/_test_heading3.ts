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

// Create a mock document JSON with a placeholder
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

const { element: result } = replacer({
    json: mockJson as any,
    patch: {
        type: PatchType.DOCUMENT,
        children: [
            new Paragraph({
                text: "Blah Blah Blah",
                heading: HeadingLevel.HEADING_1
            })
        ]
    },
    patchText: "{{heading_paragraph}}",
    context,
    keepOriginalStyles: true
});

const outputXml = js2xml(result as any, {
    attributeValueFn: (str) =>
        String(str)
            .replace(/&(?!amp;|lt;|gt;|quot;|apos;)/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;"),
});

console.log("Output XML:", outputXml);
console.log("\nContains pStyle?", outputXml.includes("pStyle"));
console.log("Contains Heading1?", outputXml.includes("Heading1"));
