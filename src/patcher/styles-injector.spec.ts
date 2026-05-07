import { describe, expect, it } from "vitest";

import type { Element } from "xml-js";

import { collectDefinedStyleIds, collectUsedStyleIds, injectMissingStyles } from "./styles-injector";

const createStylesJson = (styleIds: readonly string[]): Element => ({
    elements: [
        {
            type: "element",
            name: "w:styles",
            elements: styleIds.map((id) => ({
                type: "element",
                name: "w:style",
                attributes: {
                    "w:type": "paragraph",
                    "w:styleId": id,
                },
                elements: [
                    {
                        type: "element",
                        name: "w:name",
                        attributes: { "w:val": id.toLowerCase() },
                    },
                ],
            })),
        },
    ],
});

const createDocumentWithPStyles = (styleIds: readonly string[]): Element => ({
    elements: [
        {
            type: "element",
            name: "w:document",
            elements: [
                {
                    type: "element",
                    name: "w:body",
                    elements: styleIds.map((styleId) => ({
                        type: "element",
                        name: "w:p",
                        elements: [
                            {
                                type: "element",
                                name: "w:pPr",
                                elements: [
                                    {
                                        type: "element",
                                        name: "w:pStyle",
                                        attributes: { "w:val": styleId },
                                    },
                                ],
                            },
                        ],
                    })),
                },
            ],
        },
    ],
});

describe("styles-injector", () => {
    describe("collectUsedStyleIds", () => {
        it("should collect w:pStyle attribute values from a paragraph", () => {
            const document = createDocumentWithPStyles(["Heading1", "Heading2"]);
            const result = collectUsedStyleIds(document);

            expect(result.has("Heading1")).toBe(true);
            expect(result.has("Heading2")).toBe(true);
        });

        it("should return an empty set when no w:pStyle elements exist", () => {
            const document: Element = {
                elements: [
                    {
                        type: "element",
                        name: "w:document",
                        elements: [
                            {
                                type: "element",
                                name: "w:body",
                                elements: [
                                    {
                                        type: "element",
                                        name: "w:p",
                                        elements: [
                                            {
                                                type: "element",
                                                name: "w:r",
                                                elements: [{ type: "element", name: "w:t", elements: [{ type: "text", text: "Hello" }] }],
                                            },
                                        ],
                                    },
                                ],
                            },
                        ],
                    },
                ],
            };
            const result = collectUsedStyleIds(document);
            expect(result.size).toBe(0);
        });

        it("should collect unique style IDs (no duplicates)", () => {
            const document = createDocumentWithPStyles(["Heading1", "Heading1", "Heading2"]);
            const result = collectUsedStyleIds(document);

            expect(result.size).toBe(2);
            expect(result.has("Heading1")).toBe(true);
            expect(result.has("Heading2")).toBe(true);
        });
    });

    describe("collectDefinedStyleIds", () => {
        it("should collect styleId values from w:style elements", () => {
            const stylesJson = createStylesJson(["Normal", "Heading1", "Title"]);
            const result = collectDefinedStyleIds(stylesJson);

            expect(result.has("Normal")).toBe(true);
            expect(result.has("Heading1")).toBe(true);
            expect(result.has("Title")).toBe(true);
        });

        it("should return an empty set when no w:style elements exist", () => {
            const stylesJson: Element = {
                elements: [
                    {
                        type: "element",
                        name: "w:styles",
                        elements: [],
                    },
                ],
            };
            const result = collectDefinedStyleIds(stylesJson);
            expect(result.size).toBe(0);
        });
    });

    describe("injectMissingStyles", () => {
        it("should inject Heading1 style when it is missing from styles.xml", () => {
            const stylesJson = createStylesJson(["Normal"]);
            const usedStyleIds = new Set(["Heading1"]);

            injectMissingStyles(stylesJson, usedStyleIds);

            const stylesRoot = stylesJson.elements![0];
            const styleIds = stylesRoot.elements!.map((e) => e.attributes?.["w:styleId"]);
            expect(styleIds).toContain("Heading1");
        });

        it("should inject all missing heading styles", () => {
            const stylesJson = createStylesJson(["Normal"]);
            const usedStyleIds = new Set(["Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Title"]);

            injectMissingStyles(stylesJson, usedStyleIds);

            const definedIds = collectDefinedStyleIds(stylesJson);
            expect(definedIds.has("Heading1")).toBe(true);
            expect(definedIds.has("Heading2")).toBe(true);
            expect(definedIds.has("Heading3")).toBe(true);
            expect(definedIds.has("Heading4")).toBe(true);
            expect(definedIds.has("Heading5")).toBe(true);
            expect(definedIds.has("Heading6")).toBe(true);
            expect(definedIds.has("Title")).toBe(true);
        });

        it("should not inject a style that is already defined", () => {
            const stylesJson = createStylesJson(["Normal", "Heading1"]);
            const originalElementCount = stylesJson.elements![0].elements!.length;
            const usedStyleIds = new Set(["Heading1"]);

            injectMissingStyles(stylesJson, usedStyleIds);

            const stylesRoot = stylesJson.elements![0];
            expect(stylesRoot.elements!.length).toBe(originalElementCount);
        });

        it("should not inject an unknown/custom style", () => {
            const stylesJson = createStylesJson(["Normal"]);
            const usedStyleIds = new Set(["MyCustomStyle"]);
            const originalElementCount = stylesJson.elements![0].elements!.length;

            injectMissingStyles(stylesJson, usedStyleIds);

            const stylesRoot = stylesJson.elements![0];
            expect(stylesRoot.elements!.length).toBe(originalElementCount);
        });

        it("should leave styles.xml unchanged if no missing styles are needed", () => {
            const stylesJson = createStylesJson(["Normal", "Heading1", "Title"]);
            const usedStyleIds = new Set(["Normal", "Heading1"]);
            const originalElementCount = stylesJson.elements![0].elements!.length;

            injectMissingStyles(stylesJson, usedStyleIds);

            const stylesRoot = stylesJson.elements![0];
            expect(stylesRoot.elements!.length).toBe(originalElementCount);
        });

        it("should produce valid Heading1 style XML structure", () => {
            const stylesJson = createStylesJson(["Normal"]);
            const usedStyleIds = new Set(["Heading1"]);

            injectMissingStyles(stylesJson, usedStyleIds);

            const stylesRoot = stylesJson.elements![0];
            const heading1 = stylesRoot.elements!.find(
                (e) => e.name === "w:style" && e.attributes?.["w:styleId"] === "Heading1",
            );

            expect(heading1).toBeDefined();
            expect(heading1?.attributes?.["w:type"]).toBe("paragraph");
            expect(heading1?.elements?.some((e) => e.name === "w:name")).toBe(true);
        });
    });
});
