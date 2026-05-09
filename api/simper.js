// api/simper.js — Vercel Serverless Function
// Download template_simper.docx dari GitHub, replace placeholders, return docx

const https = require('https');
const { Readable } = require('stream');

// ── Download file dari URL sebagai Buffer ──────────────────────
function downloadBuffer(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        return downloadBuffer(res.headers.location).then(resolve).catch(reject);
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

// ── Escape XML special chars ───────────────────────────────────
function escXml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ── Replace placeholder yang mungkin terfragmentasi oleh XML tags ──
// Word sering memecah text runs, misal {{NAMA}} jadi <w:t>{{</w:t><w:t>NAMA</w:t><w:t>}}</w:t>
// Solusi: clean XML dulu, replace, lalu restore
function replacePlaceholders(xml, data) {
  // Step 1: Gabungkan fragmented runs dalam satu paragraph
  // Hapus penutup/pembuka run tag yang memisahkan placeholder
  // Pattern: tutup w:t, tutup w:r, buka w:r baru, buka w:t
  let cleaned = xml;

  // Step 2: Dulu bersihkan karakter invisible di antara {{ dan }}
  // Ganti semua {{...}} yang mungkin terpotong XML tags
  // Teknik: hapus semua XML tags DI DALAM area placeholder
  cleaned = cleaned.replace(/\{\{([^}]*)\}\}/g, (match) => match); // no-op dulu

  // Step 3: Handle fragmented placeholders
  // Word memecah text seperti: {</w:t></w:r><w:r><w:t>{NAMA}}</w:t>
  // Kita remove closing/opening tags di antara karakter placeholder
  const placeholderPattern = /(\{)(<\/w:t>(?:<\/w:r>)?(?:<w:r[^>]*>)?(?:<w:rPr>[^<]*(?:<[^>]+>[^<]*)*<\/w:rPr>)?(?:<w:t[^>]*>)?)+(\{[A-Z_]+\}\})/g;
  cleaned = cleaned.replace(placeholderPattern, '$1$3');

  const closingPattern = /(\{\{[A-Z_]+\})(<\/w:t>(?:<\/w:r>)?(?:<w:r[^>]*>)?(?:<w:rPr>[^<]*(?:<[^>]+>[^<]*)*<\/w:rPr>)?(?:<w:t[^>]*>)?)+(\})/g;
  cleaned = cleaned.replace(closingPattern, '$1$3');

  // Step 4: Replace placeholders dengan data
  const map = {
    '{{NAMA}}': escXml(data.nama),
    '{{NO_BADGE}}': escXml(data.no_badge),
    '{{COMPANY}}': escXml(data.company),
    '{{JABATAN}}': escXml(data.jabatan),
    '{{SIMPER_NUM}}': escXml(data.simper_num),
    '{{KATEGORI}}': escXml(data.kategori),
    '{{EXP_DATE}}': escXml(data.exp_date),
    '{{ISSUED_DATE}}': escXml(data.issued_date),
    '{{VEHICLE_LIST}}': escXml((data.vehicle_list || []).join(', ')),
    '{{TAHUN}}': escXml(String(new Date().getFullYear())),
  };

  for (const [placeholder, value] of Object.entries(map)) {
    cleaned = cleaned.split(placeholder).join(value);
  }

  return cleaned;
}

// ── Main Handler ───────────────────────────────────────────────
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ status: 'error', message: 'Method not allowed' });

  try {
    const data = req.body;
    if (!data || !data.nama) {
      return res.status(400).json({ status: 'error', message: 'Data tidak lengkap' });
    }

    // Download template dari GitHub
    const templateUrl = 'https://raw.githubusercontent.com/decko28/portalhsr/main/template_simper.docx';
    const templateBuffer = await downloadBuffer(templateUrl);

    // Unzip docx (docx adalah ZIP)
    const AdmZip = require('adm-zip');
    const zip = new AdmZip(templateBuffer);
    const zipEntries = zip.getEntries();

    // Replace placeholders di word/document.xml
    const docXmlEntry = zip.getEntry('word/document.xml');
    if (!docXmlEntry) throw new Error('word/document.xml tidak ditemukan di template');

    let docXml = docXmlEntry.getData().toString('utf8');
    docXml = replacePlaceholders(docXml, data);
    zip.updateFile('word/document.xml', Buffer.from(docXml, 'utf8'));

    // Juga replace di header/footer jika ada
    for (const entry of zipEntries) {
      const name = entry.entryName;
      if (name.startsWith('word/header') || name.startsWith('word/footer')) {
        try {
          let content = entry.getData().toString('utf8');
          content = replacePlaceholders(content, data);
          zip.updateFile(name, Buffer.from(content, 'utf8'));
        } catch(e) { /* skip */ }
      }
    }

    // Output sebagai buffer
    const outputBuffer = zip.toBuffer();
    const filename = `SIMPER_${(data.nama || '').replace(/\s+/g, '_')}_${(data.simper_num || '').replace(/\//g, '-')}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', outputBuffer.length);
    return res.status(200).send(outputBuffer);

  } catch (err) {
    console.error('simper API error:', err);
    return res.status(500).json({ status: 'error', message: err.message });
  }
};
