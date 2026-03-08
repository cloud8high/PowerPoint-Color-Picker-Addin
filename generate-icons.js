/**
 * generate-icons.js
 * Generates PNG icon files for the PowerPoint add-in ribbon.
 * Run with: node generate-icons.js
 */

'use strict';

const zlib = require('zlib');
const fs   = require('fs');

// CRC32 table for PNG chunk checksums
const crcTable = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
        let c = i;
        for (let j = 0; j < 8; j++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        t[i] = c;
    }
    return t;
})();

function crc32(buf) {
    let crc = 0xFFFFFFFF;
    for (const b of buf) crc = crcTable[(crc ^ b) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
}

function pngChunk(type, data) {
    const t   = Buffer.from(type);
    const len = Buffer.allocUnsafe(4);
    const crc = Buffer.allocUnsafe(4);
    len.writeUInt32BE(data.length);
    crc.writeUInt32BE(crc32(Buffer.concat([t, data])));
    return Buffer.concat([len, t, data, crc]);
}

function hexToRgb(hex) {
    return [
        parseInt(hex.slice(1, 3), 16),
        parseInt(hex.slice(3, 5), 16),
        parseInt(hex.slice(5, 7), 16),
    ];
}

// 4-quadrant icon colors
const QUADRANT_COLORS = [
    hexToRgb('#F5C000'), // top-left:     Golden Yellow
    hexToRgb('#EE6FA0'), // top-right:    Hot Pink
    hexToRgb('#20B888'), // bottom-left:  Teal
    hexToRgb('#6890C0'), // bottom-right: Cornflower Blue
];

function generatePNG(size) {
    const half = size >> 1;
    const gap  = size >= 32 ? 2 : 1; // white gap width between quadrants

    // Build raw scanlines (filter byte 0x00 + RGBA pixels)
    const rows = [];
    for (let y = 0; y < size; y++) {
        const row = Buffer.allocUnsafe(1 + size * 4); // 4 bytes per pixel (RGBA)
        row[0] = 0; // filter: None
        for (let x = 0; x < size; x++) {
            const inGap =
                (x >= half - gap && x < half) ||
                (y >= half - gap && y < half);
            const color = inGap
                ? [255, 255, 255]
                : QUADRANT_COLORS[(x >= half ? 1 : 0) + (y >= half ? 2 : 0)];
            row[1 + x * 4]     = color[0];
            row[1 + x * 4 + 1] = color[1];
            row[1 + x * 4 + 2] = color[2];
            row[1 + x * 4 + 3] = 255; // alpha: fully opaque
        }
        rows.push(row);
    }

    const raw        = Buffer.concat(rows);
    const compressed = zlib.deflateSync(raw, { level: 9 });

    // IHDR: width, height, bit-depth=8, color-type=6 (RGBA)
    const ihdr = Buffer.allocUnsafe(13);
    ihdr.writeUInt32BE(size, 0);
    ihdr.writeUInt32BE(size, 4);
    ihdr[8]  = 8; // bit depth
    ihdr[9]  = 6; // color type: RGBA
    ihdr[10] = 0; // compression
    ihdr[11] = 0; // filter
    ihdr[12] = 0; // interlace

    return Buffer.concat([
        Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), // PNG signature
        pngChunk('IHDR', ihdr),
        pngChunk('IDAT', compressed),
        pngChunk('IEND', Buffer.alloc(0)),
    ]);
}

// Generate all required sizes
const sizes = [
    { size: 16,  file: 'icon-16.png' },
    { size: 32,  file: 'icon-32.png' },
    { size: 80,  file: 'icon-80.png' },
];

sizes.forEach(({ size, file }) => {
    fs.writeFileSync(file, generatePNG(size));
    console.log(`Created ${file}  (${size}x${size})`);
});
