import { defineConfig } from 'vite';

// Library build: one ESM file per public entry. The three runtime
// dependencies stay external, so an application bundles them once.
export default defineConfig({
  build: {
    target: 'es2022',
    sourcemap: false,
    minify: false,
    lib: {
      formats: ['es'],
      entry: {
        index: 'src/index.ts',
        docx: 'src/docx/index.ts',
        layout: 'src/layout/index.ts',
        template: 'src/template/index.ts',
        recipe: 'src/recipe/index.ts',
        diagrams: 'src/diagrams/index.ts',
        'diagrams-d3': 'src/diagrams/d3.ts',
        ream: 'src/ream/index.ts',
      },
    },
    rollupOptions: { external: ['docx', 'fflate', 'fast-xml-parser', 'd3-scale', 'd3-hierarchy', 'd3-shape'] },
  },
});
