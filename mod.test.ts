import { assertEquals, assertThrows } from "jsr:@std/assert";
import { A1Notation } from "./mod.ts";

interface ParseTestCase {
    input: string;
    output: {
        sheetName?: string;
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
    };
}

const parseTestCases: ParseTestCase[] = [
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
    {
        input: "a1",
        output: {
            top: 1,
            left: 1,
            bottom: 1,
            right: 1,
        },
    },
    {
        input: "Sheet1!A1",
        output: {
            sheetName: "Sheet1",
            top: 1,
            left: 1,
            bottom: 1,
            right: 1,
        },
    },
    {
        input: "'Sheet1'!A1",
        output: {
            sheetName: "Sheet1",
            top: 1,
            left: 1,
            bottom: 1,
            right: 1,
        },
    },
];

for (const { input, output } of parseTestCases) {
    Deno.test(`A1Notation.parse(${input})`, () => {
        const a1 = A1Notation.parse(input);
        assertEquals(a1.sheetName, output.sheetName, "sheet name");
        assertEquals(a1.top, output.top, "top");
        assertEquals(a1.bottom, output.bottom, "bottom");
        assertEquals(a1.left, output.left, "left");
        assertEquals(a1.right, output.right, "right");
    });
}

const invalidParseTestCases: string[] = [
    // missing cell range
    "",
    "Sheet1!",
    "Sheet1!A1:",

    // missing sheet name
    "!A1:B2",
    "''!A1:B2",

    // invalid cell range
    "A1:B2:C3",
    "Sheet1!A1:B2:C3",
    "'Sheet1'!A1:B2:C3",
    "AAAA1:ZZZ1",

    // invalid delimiter
    "Sheet1?",
    "'Sheet1'?",
    "Sheet1!A1?B2",
    "'Sheet1'!A1?B2",
];

for (const input of invalidParseTestCases) {
    Deno.test(`A1Notation.parse(${input})`, () => {
        assertThrows(() => {
            A1Notation.parse(input);
        });
    });
}

interface ToStringTestCase {
    input: {
        sheetName?: string;
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
    };
    output: string;
}

const toStringTestCases: ToStringTestCase[] = [
    {
        input: {
            sheetName: "Sheet1",
            top: 1,
            bottom: 2,
            left: 1,
            right: 2,
        },
        output: "Sheet1!A1:B2",
    },
    {
        input: {
            sheetName: "Sheet1",
            left: 1,
            right: 1,
        },
        output: "Sheet1!A:A",
    },
    {
        input: {
            sheetName: "Sheet1",
            top: 1,
            bottom: 2,
        },
        output: "Sheet1!1:2",
    },
    {
        input: {
            sheetName: "Sheet1",
            top: 5,
            left: 1,
            right: 1,
        },
        output: "Sheet1!A5:A",
    },
    {
        input: {
            top: 1,
            bottom: 2,
            left: 1,
            right: 2,
        },
        output: "A1:B2",
    },
    {
        input: {
            sheetName: "Sheet1",
        },
        output: "Sheet1",
    },
    {
        input: {
            sheetName: "My Custom Sheet",
            left: 1,
            right: 1,
        },
        output: "'My Custom Sheet'!A:A",
    },
    {
        input: {
            sheetName: "My Custom Sheet",
        },
        output: "'My Custom Sheet'",
    },
    {
        input: {
            top: 1,
            left: 1,
            bottom: 1,
            right: 1,
        },
        output: "A1",
    },
    {
        input: {
            top: 1,
            left: 1,
        },
        output: "A1",
    },
    {
        input: {
            sheetName: "A1",
        },
        output: "'A1'",
    },
    {
        input: {
            top: 1,
            left: 18278,
            bottom: 1,
            right: 18278,
        },
        output: "ZZZ1",
    },
    {
        input: {
            sheetName: "AAAA1",
        },
        output: "AAAA1",
    },
    {
        input: {
            sheetName: "A",
        },
        output: "A",
    },
];

for (const { input, output } of toStringTestCases) {
    Deno.test(`A1Notation.toString(${JSON.stringify(input)})`, () => {
        const a1 = new A1Notation();
        a1.sheetName = input.sheetName;
        a1.top = input.top;
        a1.bottom = input.bottom;
        a1.left = input.left;
        a1.right = input.right;
        assertEquals(`${a1}`, output);
    });
}
