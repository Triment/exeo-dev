Bun.build({
    entrypoints: ['./index.ts', 'reword.ts'],
    outdir: './out',
    minify: {
      whitespace: false,
      identifiers: false,
      syntax: true,
    },
    external: ["exceljs"],
    splitting: true
  })