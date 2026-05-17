// api/simper.js — Vercel Serverless Function
// Menggunakan PizZip untuk preserve format/layout template dengan sempurna

const https = require('https');

const TEMPLATE_URL = 'https://raw.githubusercontent.com/decko28/portalhsr/main/template_simper.docx';
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

function replacePlaceholders(xml, data) {
  // Bersihkan fragmentasi XML di dalam placeholder
  let out = xml.replace(/\{\{([\s\S]*?)\}\}/g, (match) => {
    return match.replace(/<[^>]+>/g, '');
  });

  const map = {
    '{{NAMA}}':            escXml(data.nama),
    '{{NO_BADGE}}':        escXml(data.no_badge),
    '{{COMPANY}}':         escXml(data.company),
    '{{JABATAN}}':         escXml(data.jabatan),
    '{{SIMPER_NUM}}':      escXml(data.simper_num),
    '{{KATEGORI}}':        escXml(data.kategori),
    '{{EXP_DATE}}':        escXml(data.exp_date),
    '{{ISSUED_DATE}}':     escXml(data.issued_date),
    '{{JENIS_SIM}}':       escXml(data.jenis_sim       || '-'),
    '{{NO_SIM}}':          escXml(data.no_sim           || '-'),
    '{{SIM_DITERBITKAN}}': escXml(data.sim_diterbitkan  || '-'),
    '{{SIM_EXPIRY}}':      escXml(data.sim_expiry        || '-'),
    '{{TAHUN}}':           escXml(String(new Date().getFullYear())),
  };

  for (const [key, val] of Object.entries(map)) {
    out = out.split(key).join(val);
  }

  // VEHICLE_LIST — list ke bawah dengan bullet
  const vehicles = data.vehicle_list || [];
  if (vehicles.length === 0) {
    out = out.split('{{VEHICLE_LIST}}').join('-');
  } else {
    const vehicleXml = vehicles.map((v, i) => {
      const escaped = escXml('- ' + v);
      if (i === 0) return escaped;
      return `</w:t></w:r><w:r><w:br/></w:r><w:r><w:t xml:space="preserve">${escaped}`;
    }).join('');
    out = out.split('{{VEHICLE_LIST}}').join(vehicleXml);
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
    if (!data || !data.nama) {
      return res.status(400).json({ status: 'error', message: 'Data tidak lengkap' });
    }

    const templateBuffer = await downloadBuffer(TEMPLATE_URL);

    // PizZip — library yang preserve format docx
    const PizZip = require('pizzip');
    const zip = new PizZip(templateBuffer);

    // Replace placeholders di document.xml
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('word/document.xml tidak ditemukan di template');
    let docXml = docFile.asText();
    docXml = replacePlaceholders(docXml, data);
    zip.file('word/document.xml', docXml);

    // Replace di header/footer jika ada
    Object.keys(zip.files).forEach(name => {
      if (/^word\/(header|footer)\d*\.xml$/.test(name)) {
        try {
          let content = zip.file(name).asText();
          content = replacePlaceholders(content, data);
          zip.file(name, content);
        } catch(e) {}
      }
    });

    // Replace foto karyawan
    if (data.foto_b64 && data.foto_b64.startsWith('data:')) {
      const b64 = data.foto_b64.split(',')[1];
      if (b64 && zip.file(FOTO_IMAGE_NAME)) {
        zip.file(FOTO_IMAGE_NAME, Buffer.from(b64, 'base64'));
      }
    }

    // Generate — preserve format asli
    const outputBuffer = zip.generate({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });

    const filename = `SIMPER_${(data.nama||'').replace(/\s+/g,'_')}_${(data.simper_num||'').replace(/\//g,'-')}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', outputBuffer.length);
    return res.status(200).send(outputBuffer);

  } catch (err) {
    console.error('simper API error:', err);
    return res.status(500).json({ status: 'error', message: err.message });
  }
};
