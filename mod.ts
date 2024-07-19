/**
 * Represents a range in A1 notation in a spreadsheet.
 * A1 notation is a way to refer to a specific cell or range of cells in spreadsheets.
 *
 * @export
 * @class A1Notation
 */
export class A1Notation {
  /**
   * The name of the sheet this range belongs to. Optional.
   *
   * @type {?string}
   * @memberof A1Notation
   */
  readonly sheetName?: string;

  /**
   * The left column number of the range. Optional.
   *
   * @type {?number}
   * @memberof A1Notation
   */
  readonly left?: number;

  /**
   * The top row number of the range. Optional.
   *
   * @type {?number}
   * @memberof A1Notation
   */
  readonly top?: number;

  /**
   * The right column number of the range. Optional.
   *
   * @type {?number}
   * @memberof A1Notation
   */
  readonly right?: number;

  /**
   * The bottom row number of the range. Optional.
   *
   * @type {?number}
   * @memberof A1Notation
   */
  readonly bottom?: number;

  constructor(
    sheetName?: string,
    left?: number,
    top?: number,
    right?: number,
    bottom?: number,
  ) {
    this.sheetName = sheetName;
    this.left = left;
    this.top = top;
    this.right = right;
    this.bottom = bottom;
    if (right === undefined && bottom === undefined) {
      this.right = left;
      this.bottom = top;
    }
  }

  /**
   * Parses a string in A1 notation and returns an A1Notation object.
   * A1 notation is a way to refer to a specific cell or range of cells in spreadsheets.
   * This method uses a Parser to interpret the string and extract components like sheet name,
   * and cell positions. It then constructs an A1Notation object with these components.
   *
   * @example basic usage
   * ```ts
   * import { A1Notation } from "@shogo82148/a1notation";
   *
   * const a1 = A1Notation.parse("Sheet1!A1:B2");
   * console.log(a1.sheetName); // "Sheet1"
   * console.log(a1.left); // 1
   * console.log(a1.top); // 1
   * console.log(a1.right); // 2
   * console.log(a1.bottom); // 2
   * ```
   *
   * @static
   * @param {string} s - The string in A1 notation to be parsed.
   * @returns {A1Notation} An A1Notation object representing the parsed range.
   */
  static parse(s: string): A1Notation {
    const parser = new Parser(s);
    const { sheetName, cell1, cell2 } = parser.parse();
    let left: number | undefined;
    let top: number | undefined;
    let right: number | undefined;
    let bottom: number | undefined;

    if (cell1) {
      const { row, col } = parseCell(cell1);
      top = row;
      left = col;
    }
    if (cell2) {
      const { row, col } = parseCell(cell2);
      bottom = row;
      right = col;
    }
    return new A1Notation(sheetName, left, top, right, bottom);
  }

  /**
   * Converts the A1Notation object to a string representation in A1 notation.
   * This method constructs the A1 notation string for the range represented by the A1Notation object.
   * It handles different scenarios such as:
   * - Single cell reference (e.g., "A1")
   * - Range reference (e.g., "A1:B2")
   * - Sheet name inclusion (e.g., "Sheet1!A1:B2")
   *
   * @example basic usage
   * import { A1Notation } from "@shogo82148/a1notation";
   *
   * const a1 = new A1Notation("Sheet1", 3, 2);
   * console.log(`${a1}`); // Sheet1!C2
   *
   * @returns {string} The string representation of the A1Notation object in A1 notation.
   */
  toString(): string {
    let cell1 = "";
    let cell2 = "";
    if (this.left) {
      cell1 += toBase26(this.left);
    }
    if (this.top) {
      cell1 += this.top.toString();
    }
    if (this.right) {
      cell2 += toBase26(this.right);
    }
    if (this.bottom) {
      cell2 += this.bottom.toString();
    }

    let cell = "";
    if (this.left && this.right && this.top && this.bottom) {
      if (this.left === this.right && this.top === this.bottom) {
        cell = cell1;
      } else {
        cell = cell1 + ":" + cell2;
      }
    } else if (cell1 && cell2) {
      cell = cell1 + ":" + cell2;
    } else {
      cell = cell1;
    }
    if (this.sheetName) {
      if (cell) {
        return escapeSheetName(this.sheetName) + "!" + cell;
      } else {
        return escapeSheetName(this.sheetName);
      }
    }
    return cell;
  }
}

const EOF = "EOF";
const reCell = /^([a-z]{0,3})([0-9]+)$/i;

interface InterimResult {
  sheetName?: string;
  cell1?: string;
  cell2?: string;
}

class Parser {
  buf: string[];
  index = 0;

  constructor(s: string) {
    this.buf = [...s];
  }

  peek(): string {
    if (this.index < this.buf.length) {
      return this.buf[this.index];
    }
    return EOF;
  }

  next() {
    this.index++;
  }

  parse(): InterimResult {
    const ch = this.peek();
    if (ch === "'") {
      return this.parseSheetNameAndCell();
    }
    if (/[a-z0-9]/i.test(ch)) {
      return this.parseCell();
    }
    throw new Error(`unexpected character: ${ch}`);
  }

  parseSheetNameAndCell(): InterimResult {
    const sheetName = this.parseQuotedName();
    if (this.peek() === EOF) {
      // it is `'My Custom Sheet'`
      return { sheetName };
    }
    if (this.peek() !== "!") {
      throw new Error(`expected "!", but got ${this.peek()}`);
    }
    this.next(); // skip "!"

    const cell1 = this.parseName();
    if (this.peek() === EOF) {
      // it is `'My Custom Sheet'!A1`
      const cell2 = cell1;
      return { sheetName, cell1, cell2 };
    }

    if (this.peek() !== ":") {
      throw new Error(`expected ":", but got ${this.peek()}`);
    }
    this.next(); // skip ":"

    const cell2 = this.parseName();
    if (this.peek() !== EOF) {
      throw new Error(`expected EOF, but got ${this.peek()}`);
    }
    return { sheetName, cell1, cell2 };
  }

  parseCell(): InterimResult {
    const name = this.parseName(); // name is a sheet name or cell name
    const ch = this.peek();

    if (ch === EOF) {
      if (reCell.test(name)) {
        // it is `A1`
        return { cell1: name, cell2: name };
      }
      // it is `Sheet1`
      return { sheetName: name };
    }

    if (ch === "!") {
      this.next(); // skip "!"

      // name is a sheet name
      const sheetName = name;
      const cell1 = this.parseName();
      const ch = this.peek();
      if (ch === EOF) {
        // it is `Sheet1!A1`
        const cell2 = cell1;
        return { sheetName, cell1, cell2 };
      }
      if (ch === ":") {
        this.next(); // skip ":"

        const cell2 = this.parseName();
        if (this.peek() !== EOF) {
          throw new Error(`expected EOF, but got ${this.peek()}`);
        }
        // it is `Sheet1!A1:B2`
        return { sheetName, cell1, cell2 };
      }
      throw new Error(`unexpected character: ${ch}`);
    }

    if (ch === ":") {
      this.next(); // skip ":"

      const cell1 = name;
      const cell2 = this.parseName();
      if (this.peek() !== EOF) {
        throw new Error(`expected EOF, but got ${this.peek()}`);
      }
      // it is `A1:B2`
      return { cell1, cell2 };
    }

    throw new Error(`expected "!" or ":", but got ${ch}`);
  }

  parseQuotedName(): string {
    if (this.peek() !== "'") {
      throw new Error(`unexpected character: ${this.peek()}`);
    }
    this.next(); // skip "'"

    // search corresponding "'"
    let name = "";
    for (;;) {
      const ch = this.peek();
      if (ch === "'") {
        this.next();
        if (this.peek() !== "'") {
          break;
        }
      }
      name += ch;
      this.next();
    }
    if (name === "") {
      throw new Error("invalid sheet name");
    }
    return name;
  }

  parseName(): string {
    let name = "";
    for (;;) {
      const ch = this.peek();
      if (ch === EOF || !/[a-zA-Z0-9]/.test(ch)) {
        break;
      }
      name += ch;
      this.next();
    }
    if (name === "") {
      throw new Error("invalid cell name");
    }
    return name;
  }
}

interface Cell {
  col?: number;
  row?: number;
}

function parseCell(s: string): Cell {
  const cell: Cell = {};
  const m = s.match(/^([a-z]{0,3})([0-9]*)$/i);
  if (!m) {
    throw new Error(`invalid cell name: ${s}`);
  }
  if (m[1]) {
    cell.col = fromBase26(m[1]);
  }
  if (m[2]) {
    cell.row = parseInt(m[2], 10);
  }
  return cell;
}

const charCodeA = "A".charCodeAt(0);
const charCodeZ = "Z".charCodeAt(0);
const charCode_a = "a".charCodeAt(0);
const charCode_z = "z".charCodeAt(0);

function fromBase26(s: string): number {
  let n = 0;
  for (const ch of s) {
    const code = ch.codePointAt(0);
    if (code === undefined) {
      throw new Error("invalid character");
    }
    if (charCodeA <= code && code <= charCodeZ) {
      n = n * 26 + code - charCodeA + 1;
    } else if (charCode_a <= code && code <= charCode_z) {
      n = n * 26 + code - charCode_a + 1;
    } else {
      throw new Error(`invalid character: ${ch}`);
    }
  }
  return n;
}

function toBase26(num: number): string {
  let s = "";
  let n = num;
  while (n > 0) {
    n--;
    s = String.fromCharCode(charCodeA + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function escapeSheetName(s: string): string {
  if (reCell.test(s)) {
    return `'${s}'`;
  }
  if (/^[a-z0-9]+$/i.test(s)) {
    return s;
  }
  return "'" + s.replace(/'/g, "''") + "'";
}
