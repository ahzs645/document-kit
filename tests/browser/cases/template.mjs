// The template API in a real browser: a Word template written with the `docx`
// library in the page, its tags listed, checked and filled from made-up data:
// a field split across runs, a condition, a table row repeated per item, a
// number format and a header. The template goes in as a Blob, as a page holds
// a file a person picked.

import { Document, Header, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from 'docx';
import { strFromU8, unzipSync } from 'fflate';
import { checkTemplate, fillTemplate, templateTags } from 'document-kit/template';

function check(condition, message) {
  if (!condition) throw new Error(message);
}

function template() {
  const cell = (text) => new TableCell({ children: [new Paragraph(text)] });
  return new Document({
    sections: [
      {
        headers: { default: new Header({ children: [new Paragraph('{{ site }}')] }) },
        children: [
          new Paragraph({
            children: [new TextRun({ text: 'Prepared for {{ con', bold: true }), new TextRun({ text: 'tact }}.', italics: true })],
          }),
          new Paragraph('Status: {{ if approved }}approved{{ else }}pending{{ endif }}.'),
          new Table({
            rows: [
              new TableRow({ tableHeader: true, children: [cell('Sample'), cell('Value')] }),
              new TableRow({ children: [cell('{{ for row in rows }}{{ row.id }}'), cell('{{ row.value | format "#,##0.00" }}{{ endfor }}')] }),
            ],
          }),
        ],
      },
    ],
  });
}

const textOf = (xml) => [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(([p]) => p.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&'));

export default async function templateCase() {
  const bytes = await Packer.toArrayBuffer(template());
  const blob = new Blob([bytes]);

  const tags = await templateTags(blob);
  check(tags.length === 7, `seven tags listed, got ${tags.length}: ${tags.map((t) => t.tag).join(' ')}`);
  check(tags.some((t) => t.kind === 'for' && t.item === 'row'), 'the loop is listed with its item');

  const rows = Array.from({ length: 25 }, (_, i) => ({ id: `S-${String(i + 1).padStart(3, '0')}`, value: 1000 + i * 2.5 }));
  const data = { site: 'Example Field Station', contact: 'A. Sample', approved: true, rows };

  const missing = await checkTemplate(blob, { data: { ...data, contact: undefined, site: 'x' } });
  check(missing.length === 0, 'a key present with no value is not missing');
  const { contact: _, ...withoutContact } = data;
  const problems = await checkTemplate(blob, { data: withoutContact });
  check(problems.length === 1 && problems[0].field === 'contact', `the missing field is reported, got ${JSON.stringify(problems.map((p) => p.field))}`);

  const out = await fillTemplate(blob, { data });
  const again = await fillTemplate(new Uint8Array(bytes), { data });
  check(out.length === again.length && out.every((b, i) => b === again[i]), 'a Blob and its bytes fill to the same document');

  const parts = unzipSync(out);
  const body = strFromU8(parts['word/document.xml']);
  const lines = textOf(body);
  check(lines[0] === 'Prepared for A. Sample.', `split tag filled: ${lines[0]}`);
  check(lines[1] === 'Status: approved.', `condition: ${lines[1]}`);
  check(lines.includes('S-025') && lines.includes('1,060.00'), 'every row filled and formatted');
  const tableRows = (body.match(/<w:tr[ >]/g) ?? []).length;
  check(tableRows === 26, `a heading row and 25 item rows, got ${tableRows}`);
  check(!body.includes('{{'), 'no tag left');
  const headerName = Object.keys(parts).find((name) => /^word\/header\d*\.xml$/.test(name));
  check(headerName && textOf(strFromU8(parts[headerName]))[0] === 'Example Field Station', 'header filled');

  return { tags: tags.length, problems: problems.length, bytes: out.length, tableRows, firstLine: lines[0] };
}
