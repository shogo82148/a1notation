import { assertEquals } from "jsr:@std/assert";
import { A1Notation } from "./mod.ts";

interface TestCase {
    input: string;
    output: {
        sheetName?: string;
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
    };
}

const testCases: TestCase[] = [
    // test cases from https://developers.google.com/sheets/api/guides/concepts
    {
        input: "Sheet1!A1:B2",
        output: {
            sheetName: "Sheet1",
            top: 1,
            bottom: 2,
            left: 1,
            right: 2,
        },
    },
    {
        input: "Sheet1!A:A",
        output: {
            sheetName: "Sheet1",
            left: 1,
            right: 1,
        },
    },
    {
        input: "Sheet1!1:2",
        output: {
            sheetName: "Sheet1",
            top: 1,
            bottom: 2,
        },
    },
    {
        input: "Sheet1!A5:A",
        output: {
            sheetName: "Sheet1",
            top: 5,
            left: 1,
            right: 1,
        },
    },
    {
        input: "A1:B2",
        output: {
            top: 1,
            bottom: 2,
            left: 1,
            right: 2,
        },
    },
    {
        input: "Sheet1",
        output: {
            sheetName: "Sheet1",
        },
    },
    {
        input: "'My Custom Sheet'!A:A",
        output: {
            sheetName: "My Custom Sheet",
            left: 1,
            right: 1,
        },
    },
    {
        input: "'My Custom Sheet'",
        output: {
            sheetName: "My Custom Sheet",
        },
    },
    {
        input: "A1",
        output: {
            top: 1,
            left: 1,
            bottom: 1,
            right: 1,
        },
    },
    {
        input: "'A1'",
        output: {
            sheetName: "A1",
        },
    },

    // additional test cases
    {
        input: "ZZZ1", // the rightmost cell
        output: {
            top: 1,
            left: 18278,
            bottom: 1,
            right: 18278,
        },
    },
    {
        input: "AAAA1",
        output: {
            sheetName: "AAAA1",
        },
    },
    {
        input: "A",
        output: {
            sheetName: "A",
        },
    },
];

for (const { input, output } of testCases) {
    Deno.test(`A1Notation.parse(${input})`, () => {
        const a1 = A1Notation.parse(input);
        assertEquals(a1.sheetName, output.sheetName, "sheet name");
        assertEquals(a1.top, output.top, "top");
        assertEquals(a1.bottom, output.bottom, "bottom");
        assertEquals(a1.left, output.left, "left");
        assertEquals(a1.right, output.right, "right");
    });
}
