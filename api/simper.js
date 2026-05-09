// api/simper.js — Vercel Serverless Function
// Generate SIMPER .docx dari template tanpa perlu GAS/Google Auth

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, WidthType, BorderStyle, ShadingType,
  VerticalAlign, HeightRule
} = require('docx');
const fs = require('fs');
const path = require('path');

// ── Konstanta ukuran ──────────────────────────────────────────
// Kartu SIMPER: 5.38cm × 8.56cm dalam DXA (1cm = 567 DXA)
const CARD_W = Math.round(5.38 * 567);   // 3050
const CARD_H = Math.round(8.56 * 567);   // 4853
const NAVY   = '1a3a6b';
const YELLOW = 'e8b800';
const WHITE  = 'ffffff';
const GREY   = 'aaaaaa';

// A4: margin 1.5cm = 850 DXA
const PAGE_W      = 11906;
const PAGE_H      = 16838;
const PAGE_MARGIN = 850;
const CONTENT_W   = PAGE_W - PAGE_MARGIN * 2; // 10206

// ── Helper borders ─────────────────────────────────────────────
const nb = () => ({ style: BorderStyle.NONE, size: 0, color: WHITE });
const sb = (color = NAVY, size = 6) => ({ style: BorderStyle.SINGLE, size, color });
const allNone  = { top: nb(), bottom: nb(), left: nb(), right: nb() };
const allNavy  = { top: sb(), bottom: sb(), left: sb(), right: sb() };

function cell(children, opts = {}) {
  return new TableCell({
    children,
    borders: opts.borders || allNone,
    shading: { fill: opts.fill || WHITE, type: ShadingType.CLEAR },
    width: opts.w ? { size: opts.w, type: WidthType.DXA } : undefined,
    verticalAlign: opts.va || VerticalAlign.TOP,
    margins: opts.m || { top: 40, bottom: 40, left: 60, right: 60 },
    columnSpan: opts.span,
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.sb || 0, after: opts.sa || 0 },
    children: [new TextRun({
      text: text || '',
      bold: opts.bold || false,
      size: opts.size || 16,
      color: opts.color || '000000',
      font: 'Arial',
    })],
  });
}

// ── Kartu Depan ────────────────────────────────────────────────
function buildFront(data, logoBytes, fotoBytes) {
  const logoImg = new ImageRun({ data: logoBytes, transformation: { width: 80, height: 26 }, type: 'png' });

  const fotoChildren = fotoBytes
    ? [new Paragraph({ children: [new ImageRun({ data: fotoBytes, transformation: { width: 53, height: 70 }, type: data.foto_type || 'jpeg' })], spacing: { before: 0, after: 0 } })]
    : [p('[ Foto ]', { size: 9, color: '999999', align: AlignmentType.CENTER })];

  const FOTO_W = Math.round(1.6 * 567);
  const INFO_W = CARD_W - FOTO_W;

  return new Table({
    width: { size: CARD_W, type: WidthType.DXA },
    borders: allNavy,
    rows: [
      // Logo
      new TableRow({ children: [cell(
        [new Paragraph({ alignment: AlignmentType.CENTER, children: [logoImg], spacing: { before: 20, after: 20 } })],
        { borders: { top: nb(), bottom: sb(NAVY, 8), left: nb(), right: nb() }, w: CARD_W }
      )] }),
      // Banner
      new TableRow({ children: [cell(
        [p('S I M P E R', { bold: true, size: 18, color: WHITE, align: AlignmentType.CENTER })],
        { borders: allNone, fill: NAVY, w: CARD_W, m: { top: 60, bottom: 60, left: 40, right: 40 } }
      )] }),
      // Foto + Info
      new TableRow({ children: [
        cell(fotoChildren, { w: FOTO_W, borders: allNone, m: { top: 40, bottom: 40, left: 50, right: 20 }, va: VerticalAlign.TOP }),
        cell([
          p(data.nama || '—', { bold: true, size: 13, sa: 15 }),
          p(data.no_badge || '—', { size: 11, color: '333333', sa: 15 }),
          p(data.company || '—', { size: 10, color: '555555', sa: 20 }),
          p(data.kategori || 'F', { bold: true, size: 26, color: YELLOW }),
        ], { w: INFO_W, borders: allNone, m: { top: 40, bottom: 40, left: 20, right: 40 }, va: VerticalAlign.TOP }),
      ] }),
      // Expired bar
      new TableRow({ children: [cell(
        [p('EXPIRED DATE', { size: 9, bold: true, color: NAVY, align: AlignmentType.CENTER }),
         p(data.exp_date || '—', { size: 12, bold: true, align: AlignmentType.CENTER })],
        { w: CARD_W, borders: { top: sb(NAVY, 6), bottom: sb(NAVY, 6), left: nb(), right: nb() }, m: { top: 50, bottom: 50, left: 40, right: 40 } }
      )] }),
      // Violation / Accident
      new TableRow({ children: [cell(
        [buildViolBox()],
        { w: CARD_W, borders: allNone, m: { top: 30, bottom: 30, left: 40, right: 40 } }
      )] }),
      // Footer nomor
      new TableRow({ children: [cell(
        [p(data.simper_num || '—', { size: 9, color: '999999', align: AlignmentType.CENTER })],
        { w: CARD_W, borders: allNone, m: { top: 20, bottom: 20, left: 0, right: 0 } }
      )] }),
    ],
  });
}

function buildViolBox() {
  const HW = Math.floor(CARD_W / 2);
  const mkBox = (label) => new Table({
    width: { size: HW, type: WidthType.DXA },
    borders: { top: sb(GREY, 4), bottom: sb(GREY, 4), left: sb(GREY, 4), right: sb(GREY, 4) },
    rows: [
      new TableRow({ children: [cell([p(label, { size: 8, bold: true, color: WHITE, align: AlignmentType.CENTER })],
        { fill: NAVY, borders: allNone, w: HW, m: { top: 25, bottom: 25, left: 0, right: 0 } })] }),
      new TableRow({ height: { value: 400, rule: HeightRule.EXACT }, children: [cell([p('')], { w: HW, borders: allNone })] }),
    ],
  });
  return new Table({
    width: { size: CARD_W, type: WidthType.DXA },
    borders: allNone,
    rows: [new TableRow({ children: [
      cell([mkBox('VIOLATION')], { w: HW, borders: allNone, m: { top: 0, bottom: 0, left: 0, right: 20 } }),
      cell([mkBox('ACCIDENT')],  { w: HW, borders: allNone, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
    ] })],
  });
}

// ── Kartu Belakang ─────────────────────────────────────────────
function buildBack(data, logoBytes) {
  const logoImg = new ImageRun({ data: logoBytes, transformation: { width: 80, height: 26 }, type: 'png' });
  const vehicleItems = (data.vehicle_list || []).map(v => p('• ' + v, { size: 11, sa: 20 }));

  return new Table({
    width: { size: CARD_W, type: WidthType.DXA },
    borders: allNavy,
    rows: [
      new TableRow({ children: [cell(
        [new Paragraph({ alignment: AlignmentType.CENTER, children: [logoImg], spacing: { before: 20, after: 20 } })],
        { borders: { top: nb(), bottom: sb(NAVY, 8), left: nb(), right: nb() }, w: CARD_W }
      )] }),
      new TableRow({ children: [cell(
        [p(data.simper_num || '—', { bold: true, size: 12, color: WHITE, align: AlignmentType.CENTER })],
        { fill: NAVY, borders: allNone, w: CARD_W, m: { top: 50, bottom: 50, left: 0, right: 0 } }
      )] }),
      new TableRow({ children: [cell(
        [p('KENDARAAN / ALAT DIOTORISASI', { bold: true, size: 9, color: NAVY, sa: 30 }),
         ...vehicleItems],
        { w: CARD_W, borders: allNone, m: { top: 60, bottom: 20, left: 80, right: 80 } }
      )] }),
      new TableRow({ children: [cell(
        [p('DITERBITKAN', { bold: true, size: 9, color: NAVY, sb: 40 }),
         p(data.issued_date || '—', { size: 11 })],
        { w: CARD_W, borders: { top: sb('dddddd', 2), bottom: nb(), left: nb(), right: nb() }, m: { top: 60, bottom: 40, left: 80, right: 80 } }
      )] }),
    ],
  });
}

// ── Tanda Tangan ───────────────────────────────────────────────
function buildSignature() {
  const HALF = Math.floor(CONTENT_W / 2) - 300;
  const sigBox = (title, sub) => cell([
    p(title, { bold: true, size: 18, align: AlignmentType.CENTER, sa: 40 }),
    p(sub,   { size: 13, color: '666666', align: AlignmentType.CENTER, sa: 80 }),
    new Paragraph({
      spacing: { before: 0, after: 100 },
      border: { top: sb(GREY), bottom: sb(GREY), left: sb(GREY), right: sb(GREY) },
      children: [new TextRun({ text: '\u00a0', size: 80 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: sb('333333', 4) },
      spacing: { before: 20, after: 0 },
      children: [new TextRun({ text: 'Nama & Tanda Tangan', size: 13, color: '888888', font: 'Arial' })],
    }),
  ], { w: HALF, borders: allNone, m: { top: 80, bottom: 80, left: 120, right: 120 } });

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    borders: { top: sb(NAVY, 8), bottom: nb(), left: nb(), right: nb() },
    rows: [new TableRow({ children: [
      sigBox('Kepala Teknik Tambang (KTT)', 'Menyetujui penerbitan SIMPER ini'),
      cell([p('')], { w: 600, borders: allNone }),
      sigBox('Health, Safety & Risk (HSR)', 'Memverifikasi kelengkapan dokumen'),
    ] })],
  });
}

// ── Main Handler ───────────────────────────────────────────────
module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ status: 'error', message: 'Method not allowed' });

  try {
    const data = req.body;
    if (!data || !data.nama) return res.status(400).json({ status: 'error', message: 'Data tidak lengkap' });

    // Logo STM — embedded di file ini sebagai path relatif
    const logoPath = path.join(__dirname, '../logo_stm.png');
    const logoBytes = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : null;

    // Foto karyawan dari base64 (optional)
    let fotoBytes = null;
    if (data.foto_b64 && data.foto_b64.startsWith('data:')) {
      const b64 = data.foto_b64.split(',')[1];
      fotoBytes = Buffer.from(b64, 'base64');
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H },
            margin: { top: PAGE_MARGIN, right: PAGE_MARGIN, bottom: PAGE_MARGIN, left: PAGE_MARGIN },
          },
        },
        children: [
          // ── Judul ──
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 120 },
            border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 1 } },
            children: [new TextRun({ text: 'SURAT IZIN MENGEMUDI PERUSAHAAN (SIMPER)', bold: true, size: 22, color: NAVY, font: 'Arial' })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 200 },
            children: [new TextRun({ text: `No. ${data.simper_num || '—'}`, size: 17, color: '444444', font: 'Arial' })],
          }),

          // ── Info pemegang ──
          new Table({
            width: { size: CONTENT_W, type: WidthType.DXA },
            columnWidths: [Math.floor(CONTENT_W/2), Math.floor(CONTENT_W/2)],
            borders: { top: sb('cccccc', 3), bottom: sb('cccccc', 3), left: sb('cccccc', 3), right: sb('cccccc', 3) },
            rows: [
              new TableRow({ children: [
                cell([p('Nama Lengkap', { size: 13, color: '666666' }), p(data.nama, { size: 16, bold: true })],
                  { w: Math.floor(CONTENT_W/2), borders: { top: nb(), bottom: sb('eeeeee', 2), left: nb(), right: sb('eeeeee', 2) }, m: { top: 80, bottom: 80, left: 120, right: 120 } }),
                cell([p('No. Badge', { size: 13, color: '666666' }), p(data.no_badge || '—', { size: 16, bold: true, color: NAVY })],
                  { w: Math.floor(CONTENT_W/2), borders: { top: nb(), bottom: sb('eeeeee', 2), left: nb(), right: nb() }, m: { top: 80, bottom: 80, left: 120, right: 120 } }),
              ] }),
              new TableRow({ children: [
                cell([p('Jabatan', { size: 13, color: '666666' }), p(data.jabatan || '—', { size: 16, bold: true })],
                  { w: Math.floor(CONTENT_W/2), borders: { top: nb(), bottom: sb('eeeeee', 2), left: nb(), right: sb('eeeeee', 2) }, m: { top: 80, bottom: 80, left: 120, right: 120 } }),
                cell([p('Company', { size: 13, color: '666666' }), p(data.company || '—', { size: 16, bold: true })],
                  { w: Math.floor(CONTENT_W/2), borders: { top: nb(), bottom: sb('eeeeee', 2), left: nb(), right: nb() }, m: { top: 80, bottom: 80, left: 120, right: 120 } }),
              ] }),
              new TableRow({ children: [
                cell([p('Kategori SIMPER', { size: 13, color: '666666' }), p(data.kategori || '—', { size: 16, bold: true, color: NAVY })],
                  { w: Math.floor(CONTENT_W/2), borders: { top: nb(), bottom: nb(), left: nb(), right: sb('eeeeee', 2) }, m: { top: 80, bottom: 80, left: 120, right: 120 } }),
                cell([p('Masa Berlaku s/d', { size: 13, color: '666666' }), p(data.exp_date || '—', { size: 16, bold: true, color: 'cc0000' })],
                  { w: Math.floor(CONTENT_W/2), borders: { top: nb(), bottom: nb(), left: nb(), right: nb() }, m: { top: 80, bottom: 80, left: 120, right: 120 } }),
              ] }),
            ],
          }),

          new Paragraph({ spacing: { before: 240, after: 80 }, children: [new TextRun({ text: 'Kartu SIMPER:', bold: true, size: 17, color: '444444', font: 'Arial' })] }),

          // ── Dua kartu side by side ──
          new Table({
            width: { size: CONTENT_W, type: WidthType.DXA },
            columnWidths: [CARD_W + 300, CARD_W + 300],
            borders: allNone,
            rows: [new TableRow({ children: [
              cell([
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, children: [new TextRun({ text: 'Sisi Depan', size: 13, color: '888888', font: 'Arial' })] }),
                ...(logoBytes ? [buildFront(data, logoBytes, fotoBytes)] : [p('(Logo tidak tersedia)', { size: 12, color: '999999' })]),
              ], { w: CARD_W + 300, borders: allNone, m: { top: 0, bottom: 0, left: 0, right: 300 }, va: VerticalAlign.TOP }),
              cell([
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, children: [new TextRun({ text: 'Sisi Belakang', size: 13, color: '888888', font: 'Arial' })] }),
                ...(logoBytes ? [buildBack(data, logoBytes)] : [p('(Logo tidak tersedia)', { size: 12, color: '999999' })]),
              ], { w: CARD_W + 300, borders: allNone, m: { top: 0, bottom: 0, left: 0, right: 0 }, va: VerticalAlign.TOP }),
            ] })],
          }),

          new Paragraph({ spacing: { before: 360, after: 80 },
            border: { top: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 1 } },
            children: [new TextRun({ text: 'Tanda Tangan Pengesahan', bold: true, size: 19, color: NAVY, font: 'Arial' })] }),

          buildSignature(),
        ],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = `SIMPER_${(data.nama || '').replace(/\s+/g, '_')}_${(data.simper_num || '').replace(/\//g, '-')}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.status(200).send(buffer);

  } catch (err) {
    console.error('simper API error:', err);
    res.status(500).json({ status: 'error', message: err.message });
  }
};
