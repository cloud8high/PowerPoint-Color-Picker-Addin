'use strict';

Office.onReady(() => {
    buildColorGrid();
    buildCorporateColorGrid();
    initTabs();
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
// Tab switching
// ======================================================
function initTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
            btn.classList.add('active');
            document.getElementById(btn.dataset.tab).classList.remove('hidden');
        });
    });
}

// ======================================================
// Corporate brand colors
// ======================================================
const CORP_MAIN = [
    { hex: '#000099', name: 'Corporate Blue' },
    { hex: '#000000', name: 'Black' },
    { hex: '#404040', name: 'Dark Gray' },
    { hex: '#FFFFFF', name: 'White' },
];

const CORP_NEUTRAL = [
    { hex: '#727171', name: 'Pantone 424C' },
    { hex: '#9fa0a0', name: 'Pantone Cool Gray 7C' },
    { hex: '#c9caca', name: 'Pantone Cool Gray 3C' },
    { hex: '#efefef', name: 'Pantone Cool Gray 1C' },
];

const CORP_ACCENT = [
    {
        label: 'DARK',
        colors: [
            { hex: '#206aa4', name: 'Pantone 2151C' },
            { hex: '#008caa', name: 'Pantone 3155C' },
            { hex: '#1c9576', name: 'Pantone 2417C' },
            { hex: '#669e1d', name: 'Pantone 370C' },
            { hex: '#e5c53f', name: 'Pantone 7751C' },
            { hex: '#c76b00', name: 'Pantone 2014C' },
            { hex: '#ba2947', name: 'Pantone 7419C' },
            { hex: '#b62a76', name: 'Pantone 2063C' },
            { hex: '#794b88', name: 'Pantone 7677C' },
            { hex: '#9e7220', name: 'Pantone 7558C' },
        ],
    },
    {
        label: 'BRIGHT',
        colors: [
            { hex: '#2980c4', name: 'Pantone 2143C' },
            { hex: '#00a6cb', name: 'Pantone 3125C' },
            { hex: '#27b28d', name: 'Pantone 2413C' },
            { hex: '#7cbd27', name: 'Pantone 368C' },
            { hex: '#ffdb46', name: 'Pantone 1215C' },
            { hex: '#f08300', name: 'Pantone 144C' },
            { hex: '#e03657', name: 'Pantone 710C' },
            { hex: '#db368d', name: 'Pantone 2038C' },
            { hex: '#915da3', name: 'Pantone 3593C' },
            { hex: '#be8a2b', name: 'Pantone 7556C' },
        ],
    },
    {
        label: 'MIDTONE',
        colors: [
            { hex: '#7aa2d6', name: 'Pantone 659C' },
            { hex: '#4dbfd8', name: 'Pantone 637C' },
            { hex: '#87cab2', name: 'Pantone 564C' },
            { hex: '#a9d06b', name: 'Pantone 2284C' },
            { hex: '#ffe787', name: 'Pantone 1205C' },
            { hex: '#f7aa53', name: 'Pantone 157C' },
            { hex: '#efa5a4', name: 'Pantone 701C' },
            { hex: '#e47fb0', name: 'Pantone 673C' },
            { hex: '#b08bbe', name: 'Pantone 521C' },
            { hex: '#d3ac64', name: 'Pantone 466C' },
        ],
    },
    {
        label: 'LIGHT',
        colors: [
            { hex: '#bbcce9', name: 'Pantone 658C' },
            { hex: '#acdcea', name: 'Pantone 2975C' },
            { hex: '#c2e2d2', name: 'Pantone 621C' },
            { hex: '#d2e6ae', name: 'Pantone 7485C' },
            { hex: '#fff2c2', name: 'Pantone 7499C' },
            { hex: '#fad09e', name: 'Pantone 712C' },
            { hex: '#f5c9bf', name: 'Pantone 495C' },
            { hex: '#f1bdd6', name: 'Pantone 217C' },
            { hex: '#d1bedc', name: 'Pantone 7437C' },
            { hex: '#e0cfbd', name: 'Pantone 2310C' },
        ],
    },
];

function isLightColor(hex) {
    const { r, g, b } = hexToRgb(hex);
    return (r * 0.299 + g * 0.587 + b * 0.114) > 210;
}

function makeSwatch(hex, name) {
    const { r, g, b } = hexToRgb(hex);
    const swatch = document.createElement('div');
    swatch.className = 'color-swatch' + (isLightColor(hex) ? ' swatch-outline' : '');
    swatch.style.backgroundColor = hex;
    swatch.title = `${name}\n${hex.toUpperCase()}\nRGB(${r}, ${g}, ${b})`;
    swatch.dataset.hex = hex;
    swatch.addEventListener('click', () => copyHex(hex.toUpperCase()));
    return swatch;
}

function buildCorpRow(container, labelText, colors) {
    container.className = 'corp-row';
    const labelEl = document.createElement('div');
    labelEl.className = 'corp-row-label';
    labelEl.textContent = labelText;
    container.appendChild(labelEl);
    colors.forEach(({ hex, name }) => container.appendChild(makeSwatch(hex, name)));
}

function buildCorporateColorGrid() {
    buildCorpRow(document.getElementById('corp-main-row'), 'MAIN', CORP_MAIN);
    buildCorpRow(document.getElementById('corp-neutral-row'), 'NEUTRAL', CORP_NEUTRAL);

    const accentGrid = document.getElementById('corp-accent-grid');
    CORP_ACCENT.forEach(({ label, colors }) => {
        const row = document.createElement('div');
        accentGrid.appendChild(row);
        buildCorpRow(row, label, colors);
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
