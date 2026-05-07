import xml from "xml";
import { Formatter } from "@export/formatter";
import { Heading1Style, Heading2Style, Heading3Style, Heading4Style, Heading5Style, Heading6Style, TitleStyle } from "@file/styles/style/default-styles";
import { Media } from "@file/media";
import { toJson } from "./util";
import type { XmlComponent } from "@file/xml-components";
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

const h1 = new Heading1Style({ run: { color: "2E74B5", size: 32 } });
const xmlStr = xml(formatter.format(h1 as unknown as XmlComponent, context));
console.log("Heading1Style XML:", xmlStr);
const json = toJson(xmlStr);
console.log("JSON elements[0]:", JSON.stringify(json.elements?.[0], null, 2));
