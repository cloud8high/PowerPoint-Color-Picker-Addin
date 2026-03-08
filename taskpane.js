'use strict';

Office.onReady(() => {
    buildColorGrid();
});

// ======================================================
// Base colors for 14 color families (middle bordered row)
// Adjust hex values here to match your palette
// ======================================================
const BASE_COLORS = [
    { hex: '#5C7A8C', name: 'Steel Blue Gray' },
    { hex: '#7A7A7A', name: 'Gray' },
    { hex: '#F5C000', name: 'Golden Yellow' },
    { hex: '#E07060', name: 'Salmon' },
    { hex: '#EE6FA0', name: 'Hot Pink' },
    { hex: '#F08020', name: 'Orange' },
    { hex: '#A0C030', name: 'Lime Green' },
    { hex: '#20B888', name: 'Teal' },
    { hex: '#00B5E0', name: 'Sky Blue' },
    { hex: '#6890C0', name: 'Cornflower Blue' },
    { hex: '#1E3C7A', name: 'Dark Navy' },
    { hex: '#C060B8', name: 'Orchid' },
    { hex: '#9E7035', name: 'Brown' },
    { hex: '#DEC090', name: 'Sand' },
];

// Row tint/shade factors
// Positive: mix with white (lighter)
// Negative: mix with black (darker)
const ROW_FACTORS = [
    0.72,   // Row 0: lightest pastel
    0.52,   // Row 1: light
    0.30,   // Row 2: slightly light
    0,      // Row 3: base color (bordered)
    -0.32,  // Row 4: dark
    -0.62,  // Row 5: darkest
];

// ======================================================
// Color calculation utilities
// ======================================================
function hexToRgb(hex) {
    return {
        r: parseInt(hex.slice(1, 3), 16),
        g: parseInt(hex.slice(3, 5), 16),
        b: parseInt(hex.slice(5, 7), 16),
    };
}

function rgbToHex(r, g, b) {
    return '#' + [r, g, b]
        .map(v => Math.max(0, Math.min(255, Math.round(v)))
            .toString(16)
            .padStart(2, '0')
            .toUpperCase())
        .join('');
}

function applyFactor(hex, factor) {
    const { r, g, b } = hexToRgb(hex);
    if (factor > 0) {
        // 白と混ぜて明るくする
        return rgbToHex(
            r + (255 - r) * factor,
            g + (255 - g) * factor,
            b + (255 - b) * factor
        );
    } else if (factor < 0) {
        // 黒と混ぜて暗くする
        const shade = Math.abs(factor);
        return rgbToHex(r * (1 - shade), g * (1 - shade), b * (1 - shade));
    }
    return hex.toUpperCase();
}

// ======================================================
// Build color grid
// ======================================================
function buildColorGrid() {
    const grid = document.getElementById('color-grid');

    ROW_FACTORS.forEach((factor, rowIndex) => {
        BASE_COLORS.forEach((color) => {
            const hex = applyFactor(color.hex, factor);

            const { r, g, b } = hexToRgb(hex);
            const swatch = document.createElement('div');
            swatch.className = 'color-swatch' + (rowIndex === 3 ? ' base-row' : '');
            swatch.style.backgroundColor = hex;
            swatch.title = `${color.name}\n${hex}\nRGB(${r}, ${g}, ${b})`;
            swatch.dataset.hex = hex;

            swatch.addEventListener('click', () => copyHex(hex));
            grid.appendChild(swatch);
        });
    });
}

// ======================================================
// Copy to clipboard
// ======================================================
let toastTimeout = null;

function copyHex(hex) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(hex)
            .then(() => showToast(hex))
            .catch(() => fallbackCopy(hex));
    } else {
        fallbackCopy(hex);
    }
}

function fallbackCopy(hex) {
    const el = document.createElement('textarea');
    el.value = hex;
    el.style.cssText = 'position:fixed;opacity:0;top:0;left:0;';
    document.body.appendChild(el);
    el.select();
    try {
        document.execCommand('copy');
        showToast(hex);
    } catch (e) {
        alert('Copy failed: ' + hex);
    }
    document.body.removeChild(el);
}

function showToast(hex) {
    const toast = document.getElementById('toast');
    const toastSwatch = document.getElementById('toast-swatch');
    const toastHex = document.getElementById('toast-hex');

    toastSwatch.style.backgroundColor = hex;
    toastHex.textContent = hex;

    toast.classList.remove('hidden');

    if (toastTimeout) clearTimeout(toastTimeout);
    toastTimeout = setTimeout(() => {
        toast.classList.add('hidden');
    }, 2000);
}
