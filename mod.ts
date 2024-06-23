export class A1Notation {
  // the name of the sheet
  sheetName?: string;

  top?: number;
  bottom?: number;
  left?: number;
  right?: number;

  constructor(
    sheetName?: string,
    top?: number,
    bottom?: number,
    left?: number,
    right?: number,
  ) {
    this.sheetName = sheetName;
    this.top = top;
    this.bottom = bottom;
    this.left = left || top;
    this.right = right || left;
  }

  static parse(s: string): A1Notation {
    const parser = new Parser(s);
    const { sheetName, cell1, cell2 } = parser.parse();

    const a1 = new A1Notation();
    if (sheetName) {
      a1.sheetName = sheetName;
    }
    if (cell1) {
      const { row, col } = parseCell(cell1);
      a1.top = row;
      a1.left = col;
    }
    if (cell2) {
      const { row, col } = parseCell(cell2);
      a1.bottom = row;
      a1.right = col;
    }
    return a1;
  }

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
        return { sheetName, cell1 };
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
    } else if (/[a-z0-9]/i.test(ch)) {
      const name = this.parseName(); // name is a sheet name or cell name
      const ch = this.peek();
      if (ch === EOF) {
        if (reCell.test(name)) {
          // it is `A1`
          return { cell1: name, cell2: name };
        } else {
          // it is `Sheet1`
          return { sheetName: name };
        }
      } else if (ch === "!") {
        this.next(); // skip "!"

        // name is a sheet name
        const sheetName = name;
        const cell1 = this.parseName();
        const ch = this.peek();
        if (ch === EOF) {
          // it is `Sheet1!A1`
          const cell2 = cell1;
          return { sheetName, cell1, cell2 };
        } else if (ch === ":") {
          this.next(); // skip ":"

          const cell2 = this.parseName();
          if (this.peek() !== EOF) {
            throw new Error(`expected EOF, but got ${this.peek()}`);
          }
          // it is `Sheet1!A1:B2`
          return { sheetName, cell1, cell2 };
        } else {
          throw new Error(`unexpected character: ${ch}`);
        }
      } else if (ch === ":") {
        this.next(); // skip ":"

        const cell1 = name;
        const cell2 = this.parseName();
        if (this.peek() !== EOF) {
          throw new Error(`expected EOF, but got ${this.peek()}`);
        }
        // it is `A1:B2`
        return { cell1, cell2 };
      } else {
        throw new Error(`expected "!" or ":", but got ${ch}`);
      }
    } else {
      throw new Error(`unexpected character: ${ch}`);
    }
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
  row?: number;
  col?: number;
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
    } else if (charCodeA <= code && code <= charCodeZ) {
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
