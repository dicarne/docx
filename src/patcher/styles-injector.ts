/**
 * Styles injector module for document patching.
 *
 * When patching a document, the template may not have all necessary styles
 * (e.g., heading styles) defined in its styles.xml. This module detects
 * which styles are used in the patched content and injects missing standard
 * styles (headings, title, etc.) into the styles.xml of the output document.
 *
 * @module
 */
import xml from "xml";
import type { Element } from "xml-js";

import { Formatter } from "@export/formatter";
import {
    Heading1Style,
    Heading2Style,
    Heading3Style,
    Heading4Style,
    Heading5Style,
    Heading6Style,
    TitleStyle,
} from "@file/styles/style/default-styles";
import type { XmlComponent } from "@file/xml-components";

import { toJson } from "./util";

const formatter = new Formatter();

/**
 * Map of known style IDs to their default style factory functions.
 *
 * These are the standard heading and document styles that may be needed
 * in a patched document. When a patch references a style that is not
 * in the template's styles.xml, it will be generated from these defaults.
 */
const DEFAULT_STYLE_FACTORIES: Readonly<Record<string, () => XmlComponent>> = {
    Heading1: () =>
        new Heading1Style({
            run: {
                color: "2E74B5",
                size: 32,
            },
        }) as unknown as XmlComponent,
    Heading2: () =>
        new Heading2Style({
            run: {
                color: "2E74B5",
                size: 26,
            },
        }) as unknown as XmlComponent,
    Heading3: () =>
        new Heading3Style({
            run: {
                color: "1F4D78",
                size: 24,
            },
        }) as unknown as XmlComponent,
    Heading4: () =>
        new Heading4Style({
            run: {
                color: "2E74B5",
                italics: true,
            },
        }) as unknown as XmlComponent,
    Heading5: () =>
        new Heading5Style({
            run: {
                color: "2E74B5",
            },
        }) as unknown as XmlComponent,
    Heading6: () =>
        new Heading6Style({
            run: {
                color: "1F4D78",
            },
        }) as unknown as XmlComponent,
    Title: () =>
        new TitleStyle({
            run: {
                size: 56,
            },
        }) as unknown as XmlComponent,
} as const;

/**
 * Collects all style IDs referenced by w:pStyle elements within an XML element tree.
 *
 * Recursively traverses the element structure looking for w:pStyle elements
 * and collects their w:val attribute values.
 *
 * @param element - The XML element to search
 * @returns Set of style IDs found in w:pStyle elements
 */
export const collectUsedStyleIds = (element: Element): ReadonlySet<string> => {
    const styleIds = new Set<string>();
    collectUsedStyleIdsRecursive(element, styleIds);
    return styleIds;
};

// eslint-disable-next-line functional/prefer-readonly-type
const collectUsedStyleIdsRecursive = (element: Element, styleIds: Set<string>): void => {
    if (element.name === "w:pStyle" && element.attributes?.["w:val"]) {
        // eslint-disable-next-line functional/immutable-data
        styleIds.add(String(element.attributes["w:val"]));
    }

    for (const child of element.elements ?? []) {
        collectUsedStyleIdsRecursive(child, styleIds);
    }
};

/**
 * Collects all style IDs that are already defined in a styles.xml element.
 *
 * Looks for w:style elements with w:styleId attributes and returns their values.
 *
 * @param stylesElement - The root styles XML element
 * @returns Set of style IDs already defined in the styles.xml
 */
export const collectDefinedStyleIds = (stylesElement: Element): ReadonlySet<string> => {
    const styleIds = new Set<string>();

    const stylesRoot = stylesElement.elements?.find((e) => e.name === "w:styles") ?? stylesElement;

    for (const child of stylesRoot.elements ?? []) {
        if (child.name === "w:style" && child.attributes?.["w:styleId"]) {
            // eslint-disable-next-line functional/immutable-data
            styleIds.add(String(child.attributes["w:styleId"]));
        }
    }

    return styleIds;
};

/**
 * Injects missing style definitions into the styles.xml element.
 *
 * For each style ID used in the patched content that is a known default style
 * (e.g., Heading1, Heading2) and is not already defined in styles.xml,
 * this function generates and appends the style definition.
 *
 * @param stylesElement - The styles.xml Element to modify
 * @param usedStyleIds - Set of style IDs referenced in the patched content
 */
export const injectMissingStyles = (stylesElement: Element, usedStyleIds: ReadonlySet<string>): void => {
    const definedStyleIds = collectDefinedStyleIds(stylesElement);

    const stylesRoot = stylesElement.elements?.find((e) => e.name === "w:styles");
    if (!stylesRoot) {
        return;
    }

    for (const styleId of usedStyleIds) {
        if (definedStyleIds.has(styleId)) {
            continue;
        }

        const factory = DEFAULT_STYLE_FACTORIES[styleId];
        if (!factory) {
            continue;
        }

        const styleXml = xml(formatter.format(factory()));
        const styleJson = toJson(styleXml);
        const styleElement = styleJson.elements?.[0];

        if (styleElement) {
            // eslint-disable-next-line functional/immutable-data
            stylesRoot.elements = [...(stylesRoot.elements ?? []), styleElement];
        }
    }
};
