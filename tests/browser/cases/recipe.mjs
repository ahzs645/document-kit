// The recipe API in a real browser: the made-up laboratory recipes of tests/recipe/fixtures (JSON) compiled against
// the kit's shared components, rendered from made-up data with a code component written here, written as a Word file
// and laid out with the bundled fonts. Everything from the built package, nothing from the network.

import { writeDocx } from 'document-kit/docx';
import { createFontProvider, layoutDocx } from 'document-kit/layout';
import {
  checkRecipe,
  compileRecipe,
  defineCodeComponents,
  isRecipeError,
  renderRecipe,
  sharedLibrary,
} from 'document-kit/recipe';

import reportTitle from '../../recipe/fixtures/components/report-title.json';
import resultsTable from '../../recipe/fixtures/components/results-table.json';
import sampleNotes from '../../recipe/fixtures/components/sample-notes.json';
import data from '../../recipe/fixtures/data/laboratory.json';
import interim from '../../recipe/fixtures/recipes/laboratory-interim.json';
import laboratory from '../../recipe/fixtures/recipes/laboratory.json';

function check(condition, message) {
  if (!condition) throw new Error(message);
}

// A part computed in code, used by an own recipe that extends the shipped one.
const code = defineCodeComponents([
  {
    id: 'sample-count',
    title: 'Sample count',
    needs: ['samples[]'],
    build: ({ value }) => [{ type: 'paragraph', content: `${(value('samples') ?? []).length} samples in this report.`, source: 'computed' }],
  },
]);

export default async function recipeCase(ctx) {
  const own = {
    kind: 'recipe',
    id: 'own-lab-with-count',
    name: 'Laboratory report with a sample count',
    extends: 'laboratory',
    changes: [{ add: { use: 'sample-count' }, after: 'notes' }],
  };
  const library = sharedLibrary({ components: [reportTitle, resultsTable, sampleNotes], recipes: [laboratory, interim, own] });

  const started = performance.now();
  const compiled = compileRecipe('own-lab-with-count', library, { code });
  const again = compileRecipe('own-lab-with-count', library, { code });
  check(JSON.stringify(compiled) === JSON.stringify(again), 'compiling twice gave different JSON');
  check(compiled.code.length === 1 && compiled.code[0].id === 'sample-count', 'the code component is listed');

  let refused = '';
  try {
    renderRecipe(compiled, data);
  } catch (error) {
    refused = isRecipeError(error, 'missing-code') ? error.code : String(error);
  }
  check(refused === 'missing-code', `rendering without the code component: ${refused}`);

  const problems = checkRecipe(compiled, data, { code });
  const { document, parts, toComplete, edits } = renderRecipe(compiled, data, {
    code,
    sources: { 'tables[].rows': 'instrument' },
    edits: { 'addressee||field|Reference': 'REF-1' },
  });
  const renderMs = Math.round(performance.now() - started);
  check(edits.targets.some((t) => t.edited === 'REF-1'), 'the hand edit was applied');

  const docx = await writeDocx(document, { sourceStyles: { edited: { highlight: 'cyan' } } });
  check(docx[0] === 0x50 && docx[1] === 0x4b, 'the Word file is a zip');
  const fonts = createFontProvider((file) => ctx.loadFont(file.file));
  const laid = await layoutDocx(docx, { fonts });
  const text = laid.pageText.join(' ').replace(/\s+/g, ' ');
  check(text.includes('3 samples in this report.'), 'the code component printed');
  check(text.includes('Reference: REF-1'), 'the edited value printed');
  check(text.includes('Table 2: Metals'), 'the second results table printed');

  const interimText = renderRecipe(compileRecipe('laboratory-interim', library), data).document.sections[0].body
    .map((block) => JSON.stringify(block))
    .join(' ');
  check(interimText.includes('Interim report.'), 'the interim recipe adds its notice');

  return {
    renderMs,
    bytes: docx.length,
    pages: laid.pageCount,
    parts: parts.map((p) => `${p.id}:${p.blocks}`),
    toComplete: toComplete.map((item) => item.what),
    problems: problems.map((p) => p.path),
    hash: compiled.from.hash.slice(0, 15),
  };
}
