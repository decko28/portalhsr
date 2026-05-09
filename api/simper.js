// api/simper.js — Vercel Serverless Function
// Download template_simper.docx dari GitHub, replace placeholders + foto, return docx

const https = require('https');

const TEMPLATE_URL = 'https://raw.githubusercontent.com/decko28/portalhsr/main/template_simper.docx';
// image2.png = foto karyawan di template
const FOTO_IMAGE_NAME = 'word/media/image2.png';

function downloadBuffer(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        return downloadBuffer(res.headers.location).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) return reject(new Error('HTTP ' + res.statusCode));
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

function escXml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Replace placeholders — handle fragmentasi XML Word
function replacePlaceholders(xml, data) {
  // Bersihkan fragmentasi: Word kadang pecah {{NAMA}} jadi
  // {{<tag>NAM</tag>A}} dst. Gabungkan dulu dalam w:t runs.
  // Strategi: hapus XML tags DI ANTARA karakter { } [A-Z_]
  // Step 1: collapse fragmented runs inside placeholder context
  let out = xml;

  // Gabungkan runs yang memisahkan placeholder
  // Pattern: closing w:t + optional w:r tags + opening w:t
  const runSep = /(<\/w:t>(?:<\/w:r>)?(?:<w:r[^>]*>)?(?:<w:rPr>[\s\S]*?<\/w:rPr>)?(?:<w:t[^>]*>)?)/g;

  // Cari semua {{ ... }} termasuk yang terpecah tag XML
  // Kita extract text content dulu, cari posisi placeholder, lalu replace di raw XML
  // Teknik paling robust: hapus tags di dalam window {{ ... }}
  out = out.replace(/\{\{([\s\S]*?)\}\}/g, (match) => {
    // Hapus XML tags di dalam placeholder
    const clean = match.replace(/<[^>]+>/g, '');
    return clean; // misal {{KATEGORI}} setelah dibersihkan
  });

  const map = {
    '{{NAMA}}':         escXml(data.nama),
    '{{NO_BADGE}}':     escXml(data.no_badge),
    '{{COMPANY}}':      escXml(data.company),
    '{{JABATAN}}':      escXml(data.jabatan),
    '{{SIMPER_NUM}}':   escXml(data.simper_num),
    '{{KATEGORI}}':     escXml(data.kategori),
    '{{EXP_DATE}}':     escXml(data.exp_date),
    '{{ISSUED_DATE}}':  escXml(data.issued_date),
    '{{VEHICLE_LIST}}': escXml((data.vehicle_list || []).join(', ')),
    '{{TAHUN}}':        escXml(String(new Date().getFullYear())),
  };

  for (const [key, val] of Object.entries(map)) {
    out = out.split(key).join(val);
  }
  return out;
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ status: 'error', message: 'Method not allowed' });

  try {
    const data = req.body;
    if (!data || !data.nama) return res.status(400).json({ status: 'error', message: 'Data tidak lengkap' });

    // Download template
    const templateBuffer = await downloadBuffer(TEMPLATE_URL);

    // Proses ZIP (docx = ZIP)
    const AdmZip = require('adm-zip');
    const zip = new AdmZip(templateBuffer);

    // 1. Replace placeholders di document.xml
    const docEntry = zip.getEntry('word/document.xml');
    if (!docEntry) throw new Error('word/document.xml tidak ditemukan');
    let docXml = docEntry.getData().toString('utf8');
    docXml = replacePlaceholders(docXml, data);
    zip.updateFile('word/document.xml', Buffer.from(docXml, 'utf8'));

    // 2. Replace juga di header/footer jika ada
    for (const entry of zip.getEntries()) {
      const name = entry.entryName;
      if ((name.startsWith('word/header') || name.startsWith('word/footer')) && name.endsWith('.xml')) {
        try {
          let content = entry.getData().toString('utf8');
          content = replacePlaceholders(content, data);
          zip.updateFile(name, Buffer.from(content, 'utf8'));
        } catch(e) { /* skip */ }
      }
    }

    // 3. Replace foto karyawan (image2.png) jika ada pas_foto
    if (data.foto_b64 && data.foto_b64.startsWith('data:')) {
      const b64 = data.foto_b64.split(',')[1];
      const fotoBuffer = Buffer.from(b64, 'base64');

      // Tentukan ekstensi dari mime type
      const mimeType = data.foto_b64.split(';')[0].split(':')[1]; // image/jpeg atau image/png
      const ext = mimeType.includes('png') ? 'png' : 'jpeg';
      const targetName = ext === 'png' ? FOTO_IMAGE_NAME : FOTO_IMAGE_NAME.replace('.png', '.jpeg');

      if (ext === 'png') {
        // Langsung replace image2.png
        zip.updateFile(FOTO_IMAGE_NAME, fotoBuffer);
      } else {
        // Foto JPEG: update image2.png dengan bytes JPEG (Word tetap bisa baca)
        // Atau rename relationship
        zip.updateFile(FOTO_IMAGE_NAME, fotoBuffer);
      }
    }

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
