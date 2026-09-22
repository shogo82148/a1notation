// ex. scripts/build_npm.ts
import { build, emptyDir } from "@deno/dnt";

await emptyDir("./npm");

await build({
  entryPoints: ["./mod.ts"],
  outDir: "./npm",
  shims: {
    deno: true,
  },

  compilerOptions: {
    lib: ["ESNext", "DOM"],
  },

  package: {
    // package.json properties
    name: "@shogo82148/a1notation",
    version: Deno.args[0],
    description:
      "A TypeScript library for parsing and generating A1 notation (e.g. Sheet1!A1:B2) used in spreadsheets like Excel and Google Sheets.",
    license: "MIT",
    repository: {
      type: "git",
      url: "git+https://github.com/shogo82148/a1notation.git",
    },
    bugs: {
      url: "https://github.com/shogo82148/a1notation/issues",
    },
  },

  postBuild() {
    // steps to run after building and before running the tests
    Deno.copyFileSync("LICENSE", "npm/LICENSE");
    Deno.copyFileSync("README.md", "npm/README.md");
  },
});
