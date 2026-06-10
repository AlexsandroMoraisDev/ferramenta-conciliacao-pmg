const sharp = require('sharp');
const fs = require('fs');
const svg = fs.readFileSync('./public/pwa-icon.svg', 'utf8');

sharp(Buffer.from(svg))
  .resize(192, 192)
  .png()
  .toFile('./public/pwa-192x192.png')
  .then(() => console.log('192 done'))
  .catch(err => console.error(err));

sharp(Buffer.from(svg))
  .resize(512, 512)
  .png()
  .toFile('./public/pwa-512x512.png')
  .then(() => console.log('512 done'))
  .catch(err => console.error(err));
