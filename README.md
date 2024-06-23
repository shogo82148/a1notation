# a1notation

This library is designed to parse A1 Notation, commonly used in spreadsheets
like Excel and Google Sheets, and convert it to row and column indices. It also
allows for the reverse operation, converting row and column indices back to A1
Notation.

## Usage

## Parsing A1 Notation

```typescript
import { A1Notation } from "@shogo82148/a1notation";

const a1 = A1Notation.parse("Sheet1!A1:B2");
console.log(a1.sheetName); // "Sheet1"
console.log(a1.left); // 1
console.log(a1.top); // 1
console.log(a1.right); // 2
console.log(a1.bottom); // 2
```

## Generating A1 Notation

```typescript
import { A1Notation } from "@shogo82148/a1notation";

const a1 = new A1Notation("Sheet1", 3, 2);
console.log(`${a1}`); // Sheet1!C2
```
