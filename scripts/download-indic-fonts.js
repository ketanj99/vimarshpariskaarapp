/**
 * Download Noto Sans Gujarati & Devanagari TTF to public/fonts for PDF Indic script support.
 * Run once: node scripts/download-indic-fonts.js
 */
const fs = require('fs');
const path = require('path');
const https = require('https');

const FONTS_DIR = path.join(process.cwd(), 'public', 'fonts');
// Google Fonts repo has variable fonts; we save as -Regular.ttf for jsPDF
const URLS = {
  gujarati:
    'https://raw.githubusercontent.com/google/fonts/main/ofl/notosansgujarati/NotoSansGujarati%5Bwdth%2Cwght%5D.ttf',
  devanagari:
    'https://raw.githubusercontent.com/google/fonts/main/ofl/notosansdevanagari/NotoSansDevanagari%5Bwdth%2Cwght%5D.ttf',
};

function download(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, (res) => {
        if (res.statusCode === 302 || res.statusCode === 301) {
          return download(res.headers.location).then(resolve).catch(reject);
        }
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => resolve(Buffer.concat(chunks)));
        res.on('error', reject);
      })
      .on('error', reject);
  });
}

async function main() {
  if (!fs.existsSync(path.join(process.cwd(), 'public'))) {
    fs.mkdirSync(path.join(process.cwd(), 'public'), { recursive: true });
  }
  if (!fs.existsSync(FONTS_DIR)) {
    fs.mkdirSync(FONTS_DIR, { recursive: true });
    console.log('Created public/fonts');
  }

  const guPath = path.join(FONTS_DIR, 'NotoSansGujarati-Regular.ttf');
  const devPath = path.join(FONTS_DIR, 'NotoSansDevanagari-Regular.ttf');

  if (!fs.existsSync(guPath)) {
    try {
      console.log('Downloading Noto Sans Gujarati...');
      const buf = await download(URLS.gujarati);
      if (buf.length > 10000) {
        fs.writeFileSync(guPath, buf);
        console.log('Saved NotoSansGujarati-Regular.ttf');
      } else throw new Error('Download too small');
    } catch (e) {
      console.warn('Download failed:', e.message);
      console.log(
        'Download manually: https://fonts.google.com/noto/specimen/Noto+Sans+Gujarati → Download family → extract and put NotoSansGujarati-Regular.ttf (or the variable .ttf renamed) in public/fonts/'
      );
    }
  } else {
    console.log('NotoSansGujarati-Regular.ttf already exists');
  }

  if (!fs.existsSync(devPath)) {
    try {
      console.log('Downloading Noto Sans Devanagari...');
      const buf = await download(URLS.devanagari);
      if (buf.length > 10000) {
        fs.writeFileSync(devPath, buf);
        console.log('Saved NotoSansDevanagari-Regular.ttf');
      } else throw new Error('Download too small');
    } catch (e) {
      console.warn('Download failed:', e.message);
      console.log(
        'Download manually: https://fonts.google.com/noto/specimen/Noto+Sans+Devanagari → put TTF in public/fonts/NotoSansDevanagari-Regular.ttf'
      );
    }
  } else {
    console.log('NotoSansDevanagari-Regular.ttf already exists');
  }

  console.log('Done. Restart the app and generate PDF again for Gujarati/Hindi to display.');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
