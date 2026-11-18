/**
 * A filled document's zip timestamps are pinned: the same bytes in every time zone, each entry stamped
 * 2001-01-01 00:00. fflate reads an entry's time from a Date's local fields, so a date built from UTC
 * fields (`Date.UTC(1980, 0, 1)`) is 1979 west of Greenwich and fflate throws; the engine builds its
 * date from local fields (`new Date(2001, 0, 1)`). Node applies a TZ set while it runs, so each zone is
 * tried in this one process.
 */
import { afterAll, describe, expect, it } from 'vitest';
import { Paragraph } from 'docx';
import { fillTemplate } from '../../src/template/index.js';
import { wordTemplate } from './helpers.js';

const ZONES = ['America/Vancouver', 'Pacific/Pago_Pago', 'UTC', 'Asia/Kolkata', 'Pacific/Kiritimati'];
const initialTz = process.env.TZ;
afterAll(() => {
    if (initialTz === undefined) delete process.env.TZ;
    else process.env.TZ = initialTz;
});

/** Each central-directory entry's name and DOS date and time, from the zip's own records. */
function zipStamps(zip: Uint8Array): { name: string; date: number; time: number }[] {
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    let eocd = zip.length - 22;
    while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd--;
    const count = view.getUint16(eocd + 10, true);
    let at = view.getUint32(eocd + 16, true);
    const out: { name: string; date: number; time: number }[] = [];
    for (let i = 0; i < count; i++) {
        expect(view.getUint32(at, true)).toBe(0x02014b50);
        const nameLength = view.getUint16(at + 28, true), extra = view.getUint16(at + 30, true), comment = view.getUint16(at + 32, true);
        out.push({ name: new TextDecoder().decode(zip.subarray(at + 46, at + 46 + nameLength)), time: view.getUint16(at + 12, true), date: view.getUint16(at + 14, true) });
        at += 46 + nameLength + extra + comment;
    }
    return out;
}

describe('pinned zip timestamps', () => {
    it('writes the same bytes, stamped 2001-01-01 00:00, in every time zone', async () => {
        const template = await wordTemplate({ children: [new Paragraph('{{ site }}')] });
        const outputs: Uint8Array[] = [];
        for (const zone of ZONES) {
            process.env.TZ = zone;
            // The zone is in effect: Vancouver is 8 hours behind UTC in January.
            if (zone === 'America/Vancouver') expect(new Date(2026, 0, 15).getTimezoneOffset()).toBe(480);
            if (zone === 'Pacific/Kiritimati') expect(new Date(2026, 0, 15).getTimezoneOffset()).toBe(-840);
            const out = await fillTemplate(template, { data: { site: 'Example Field Station' } });
            const stamps = zipStamps(out);
            expect(stamps.length).toBeGreaterThan(3);
            for (const stamp of stamps) {
                // DOS date: (year - 1980) << 9 | month << 5 | day; time 00:00:00 is 0.
                expect({ name: stamp.name, date: stamp.date, time: stamp.time }, zone).toEqual({ name: stamp.name, date: ((2001 - 1980) << 9) | (1 << 5) | 1, time: 0 });
            }
            outputs.push(out);
        }
        for (const out of outputs.slice(1)) expect(out).toEqual(outputs[0]);
    });
});
