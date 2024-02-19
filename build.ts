Bun.build({
    entrypoints: ['./index.ts'],
    outdir: './out',
    minify: {
      whitespace: false,
      identifiers: false,
      syntax: true,
    },
    external: ["exceljs"],
    splitting: true
  })