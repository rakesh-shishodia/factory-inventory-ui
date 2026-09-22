import { mkdir, copyFile } from 'node:fs/promises';

await mkdir(new URL('../public/vendor/', import.meta.url), { recursive: true });
await copyFile(
  new URL('../node_modules/html5-qrcode/html5-qrcode.min.js', import.meta.url),
  new URL('../public/vendor/html5-qrcode.min.js', import.meta.url)
);
