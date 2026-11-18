// document-kit: page-number fields in the BODY of a document.
//
// §17.16.5.35 PAGE, §17.16.5.33 NUMPAGES and §17.16.5.45 PAGEREF print numbers
// only the finished pagination knows: the page a field stands on, how many
// pages there are, the page a bookmark lands on. Headers and footers are drawn
// after pagination and fill PAGE and NUMPAGES there (see `substitutePageFields`
// in ./styled-layout). In the body the numbers are part of the text being laid
// out, so they are filled from one layout and the document laid out again,
// until the numbers it prints are the numbers it gets (see `layoutStyledDocument`).
//
// A PAGE field is placed by its paragraph: the paragraph is given a bookmark
// of its own (MARK_PREFIX), whose page the pagination already records.
// PAGEREF prints the bookmark's page; a bookmark the layout did not place (one
// in a table cell, a header, a text box, or none of that name) keeps the
// field's cached text, as does a PAGE field in a table cell.

import type { BodyElement, Paragraph, Run } from '../core/document-model';
import type { LaidOutPage } from './page-doc';

/** The bookmark name given a paragraph that holds a PAGE field. Not a name Word can write. */
export const MARK_PREFIX = '\u0000dk-page-field-';

/** What the fields print, read from one layout. */
export interface FieldValues {
  /** The number of pages. */
  readonly totalPages: number;
  /** Each placed bookmark's page, as that page prints its number. */
  readonly bookmarkPages: ReadonlyMap<string, string>;
}

/** Whether a run is a page-number field the body fill handles. */
const isPageField = (run: Run): boolean => run.field !== undefined;

/**
 * Whether any paragraph of the body (table cells included), or a PAGEREF in a
 * header or footer, has a field only pagination can fill.
 */
export function hasLayoutFields(
  body: ReadonlyArray<BodyElement>,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined,
): boolean {
  if (anyParagraph(body, (p) => p.runs.some(isPageField))) return true;
  for (const content of headersFooters?.values() ?? []) {
    if (anyParagraph(content, (p) => p.runs.some((r) => r.field === 'PAGEREF'))) return true;
  }
  return false;
}

/**
 * Give every top-level body paragraph that holds a PAGE field a bookmark of
 * its own, so the pagination records its page.
 *
 * @returns The body with the marks, and how many were given.
 */
export function markPageFields(body: ReadonlyArray<BodyElement>): {
  body: Array<BodyElement>;
  marks: number;
} {
  let marks = 0;
  const out = body.map((el): BodyElement => {
    if (el.kind !== 'paragraph' || !el.paragraph.runs.some((r) => r.field === 'PAGE')) return el;
    const mark = `${MARK_PREFIX}${String(marks++)}`;
    return {
      kind: 'paragraph',
      paragraph: { ...el.paragraph, bookmarks: [...(el.paragraph.bookmarks ?? []), mark] },
    };
  });
  return { body: out, marks };
}

/**
 * What the fields print, read from a layout: the page count, and each placed
 * bookmark's page in that page's own numbering format.
 */
export function readFieldValues(
  pages: ReadonlyArray<LaidOutPage>,
  bookmarks: ReadonlyMap<string, { readonly pageIdx: number }>,
): FieldValues {
  const bookmarkPages = new Map<string, string>();
  for (const [name, at] of bookmarks) {
    const page = pages[at.pageIdx];
    if (page) bookmarkPages.set(name, page.pageLabel ?? String(at.pageIdx + 1));
  }
  return { totalPages: pages.length, bookmarkPages };
}

/**
 * The body with every page-number field's text replaced by what `values` says
 * it prints. Elements with no field are returned as they are.
 *
 * @param body   The body, marked by {@link markPageFields}.
 * @param values What the fields print.
 */
export function fillFields(
  body: ReadonlyArray<BodyElement>,
  values: FieldValues,
): Array<BodyElement> {
  return body.map((el) => fillElement(el, values, true));
}

/** The headers and footers with each PAGEREF filled (PAGE and NUMPAGES are filled per page). */
export function fillHeaderFooterRefs(
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  values: FieldValues,
): Map<string, ReadonlyArray<BodyElement>> {
  const out = new Map<string, ReadonlyArray<BodyElement>>();
  for (const [key, content] of headersFooters) {
    out.set(key, content.map((el) => fillElement(el, values, false)));
  }
  return out;
}

/**
 * What every page-number field prints, as one string: two layouts with the
 * same key printed the same numbers everywhere.
 */
export function printedKey(
  body: ReadonlyArray<BodyElement>,
  headersFooters: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined,
): string {
  const parts: Array<string> = [];
  const visit = (elements: ReadonlyArray<BodyElement>): void => {
    for (const el of elements) {
      if (el.kind === 'paragraph') {
        for (const run of el.paragraph.runs) if (run.field !== undefined) parts.push(run.text);
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
      }
    }
  };
  visit(body);
  for (const content of headersFooters?.values() ?? []) visit(content);
  return parts.join('\u0001');
}

/** Whether a bookmark name is one of the marks {@link markPageFields} gives. */
export const isPageFieldMark = (name: string): boolean => name.startsWith(MARK_PREFIX);

function fillElement(el: BodyElement, values: FieldValues, body: boolean): BodyElement {
  if (el.kind === 'paragraph') {
    const paragraph = el.paragraph;
    const wanted = body ? isPageField : (r: Run) => r.field === 'PAGEREF';
    if (!paragraph.runs.some(wanted)) return el;
    return {
      kind: 'paragraph',
      paragraph: {
        ...paragraph,
        runs: paragraph.runs.map((r) =>
          wanted(r) ? { ...r, text: fieldText(r, paragraph, values, body) } : r,
        ),
      },
    };
  }
  if (el.kind === 'table') {
    let changed = false;
    const rows = el.table.rows.map((row) => {
      let rowChanged = false;
      const cells = row.cells.map((cell) => {
        const content = cell.content.map((c) => fillElement(c, values, body));
        if (content.every((c, i) => c === cell.content[i])) return cell;
        rowChanged = true;
        return { ...cell, content };
      });
      if (!rowChanged) return row;
      changed = true;
      return { ...row, cells };
    });
    return changed ? { kind: 'table', table: { ...el.table, rows } } : el;
  }
  return el;
}

/** What one field prints: from `values` where they know, its cached text where not. */
function fieldText(run: Run, paragraph: Paragraph, values: FieldValues, body: boolean): string {
  switch (run.field) {
    case 'NUMPAGES':
      return body ? String(values.totalPages) : run.text;
    case 'PAGEREF':
      return (run.fieldTarget !== undefined && values.bookmarkPages.get(run.fieldTarget)) || run.text;
    case 'PAGE': {
      if (!body) return run.text;
      const mark = paragraph.bookmarks?.find(isPageFieldMark);
      return (mark !== undefined && values.bookmarkPages.get(mark)) || run.text;
    }
    default:
      return run.text;
  }
}

function anyParagraph(
  elements: ReadonlyArray<BodyElement>,
  test: (p: Paragraph) => boolean,
): boolean {
  for (const el of elements) {
    if (el.kind === 'paragraph' && test(el.paragraph)) return true;
    if (el.kind === 'table') {
      for (const row of el.table.rows) {
        for (const cell of row.cells) if (anyParagraph(cell.content, test)) return true;
      }
    }
  }
  return false;
}
