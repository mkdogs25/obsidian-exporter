/*
 * PDF Customiser — an Obsidian plugin
 *
 * Styles Obsidian's native "Export to PDF" (colours, fonts, designs, header,
 * footer, page numbers, watermark…) and adds Word (.docx) and PowerPoint
 * (.pptx) exporters. No build step: plain CommonJS against the Obsidian API.
 */
'use strict';

const obsidian = require('obsidian');
const {
	Plugin, PluginSettingTab, Setting, Notice, Modal, SuggestModal,
	TFile, Platform, normalizePath, debounce, setIcon,
} = obsidian;

/* ────────────────────────────────────────────────────────────────────────── *
 * Generic helpers
 * ────────────────────────────────────────────────────────────────────────── */

const isObj = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);
const clone = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)));

function deepMerge(target, src) {
	if (!isObj(src)) return target;
	for (const k of Object.keys(src)) {
		if (src[k] === undefined) continue;
		if (isObj(src[k]) && isObj(target[k])) deepMerge(target[k], src[k]);
		else target[k] = isObj(src[k]) || Array.isArray(src[k]) ? clone(src[k]) : src[k];
	}
	return target;
}

function getPath(obj, path) {
	return path.split('.').reduce((o, k) => (o == null ? undefined : o[k]), obj);
}

function setPath(obj, path, value) {
	const keys = path.split('.');
	const last = keys.pop();
	let o = obj;
	for (const k of keys) {
		if (!isObj(o[k])) o[k] = {};
		o = o[k];
	}
	o[last] = value;
}

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

const escXml = (s) => String(s == null ? '' : s)
	.replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[ch]))
	.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');

const cssString = (s) => '"' + String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\A ') + '"';

/* Colours ------------------------------------------------------------------ */

function normHex(h, fallback = '#000000') {
	if (typeof h !== 'string') return fallback;
	let s = h.trim();
	if (/^#?[0-9a-f]{3}$/i.test(s)) s = '#' + s.replace('#', '').split('').map((c) => c + c).join('');
	if (!s.startsWith('#')) s = '#' + s;
	return /^#[0-9a-f]{6}$/i.test(s) ? s.toLowerCase() : fallback;
}

function hexToRgb(h) {
	const n = normHex(h).slice(1);
	return [0, 2, 4].map((i) => parseInt(n.substr(i, 2), 16));
}

function rgbToHex(r, g, b) {
	return '#' + [r, g, b].map((v) => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('');
}

/** Blend colour a towards colour b by t (0 = a, 1 = b). */
function mix(a, b, t) {
	const A = hexToRgb(a);
	const B = hexToRgb(b);
	return rgbToHex(A[0] + (B[0] - A[0]) * t, A[1] + (B[1] - A[1]) * t, A[2] + (B[2] - A[2]) * t);
}

function luminance(h) {
	const [r, g, b] = hexToRgb(h).map((v) => {
		v /= 255;
		return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
	});
	return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

const contrastText = (h) => (luminance(h) > 0.4 ? '#111111' : '#ffffff');
const rgba = (h, a) => { const [r, g, b] = hexToRgb(h); return `rgba(${r}, ${g}, ${b}, ${a})`; };
/** Hex without '#', upper case — the form OOXML wants. */
const hx = (h) => normHex(h).slice(1).toUpperCase();

/* Fonts -------------------------------------------------------------------- */

const GENERIC_FONTS = ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-serif', 'ui-sans-serif', 'ui-monospace', 'inherit', 'initial'];

const FONT_SUGGESTIONS = [
	'Inter', 'Roboto', 'Open Sans', 'Lato', 'Montserrat', 'Poppins', 'Nunito', 'Raleway', 'Source Sans 3', 'DM Sans',
	'Space Grotesk', 'IBM Plex Sans', 'Helvetica Neue', 'Helvetica', 'Arial', 'Segoe UI', 'Calibri', 'Verdana', 'Tahoma',
	'Trebuchet MS', 'Gill Sans', 'Futura', 'Avenir', 'Oswald', 'Bebas Neue', 'Georgia', 'Times New Roman', 'Garamond',
	'EB Garamond', 'Merriweather', 'Playfair Display', 'DM Serif Display', 'Libre Baskerville', 'Baskerville', 'Lora',
	'Crimson Text', 'Cambria', 'Palatino Linotype', 'Book Antiqua', 'Didot', 'IBM Plex Serif', 'Courier New', 'Consolas',
	'JetBrains Mono', 'Fira Code', 'Source Code Pro', 'IBM Plex Mono', 'Menlo', 'Comic Sans MS',
];

function splitFonts(stack) {
	return String(stack || '').split(',').map((f) => f.trim().replace(/^['"]|['"]$/g, '')).filter(Boolean);
}

/** First concrete family in a CSS font stack (for Word/PowerPoint). */
function primaryFont(stack) {
	return splitFonts(stack).find((f) => !GENERIC_FONTS.includes(f.toLowerCase()) && !f.startsWith('var(')) || '';
}

function guessGeneric(name) {
	const n = String(name).toLowerCase();
	if (/mono|code|courier|consol|menlo/.test(n)) return 'monospace';
	if (/sans/.test(n)) return 'sans-serif';
	if (/serif|georgia|times|garamond|baskerville|playfair|lora|merriweather|cambria|palatino|didot|crimson|antiqua/.test(n)) return 'serif';
	return 'sans-serif';
}

/** Turn a user font entry ("Georgia" or a full stack) into a safe CSS font-family value. */
function cssFont(stack, fallback) {
	const parts = splitFonts(stack);
	if (!parts.length) return fallback || 'inherit';
	const out = parts.map((p) => (GENERIC_FONTS.includes(p.toLowerCase()) || p.startsWith('var(') ? p : `"${p.replace(/"/g, '')}"`));
	if (!parts.some((p) => GENERIC_FONTS.includes(p.toLowerCase()))) out.push(guessGeneric(parts[0]));
	return out.join(', ');
}

/* ────────────────────────────────────────────────────────────────────────── *
 * Designs (the "template gallery")
 * ────────────────────────────────────────────────────────────────────────── */

const BASE_COLOURS = {
	text: '#1f2328', background: '#ffffff',
	h1: '#111111', h2: '#1f2328', h3: '#1f2328', h4: '#1f2328', h5: '#374151', h6: '#4b5563',
	accent: '#2563eb', link: '#2563eb', highlight: '#fff3a3',
	blockquoteText: '#4b5563', blockquoteBorder: '#d1d5db', blockquoteBg: '#f9fafb',
	codeText: '#be185d', codeBg: '#f3f4f6',
	tableHeaderBg: '#f3f4f6', tableHeaderText: '#111111', tableBorder: '#d1d5db', tableStripe: '#f9fafb',
	hr: '#e5e7eb', headerFooter: '#6b7280',
};

const BASE_DESIGN = {
	headingStyle: 'none',      // none | underline | double-underline | overline | leftbar | band | boxed | pill | centered
	decorationLevels: 2,       // decoration applies to H1…Hn
	textAlign: 'left',         // left | justify
	paragraphSpacing: 0.75,    // em
	blockquoteStyle: 'bar',    // bar | card | quote | minimal
	tableStyle: 'grid',        // grid | striped | minimal | header
	imageStyle: 'plain',       // plain | rounded | shadow | framed
	listMarkerAccent: false,
	accentBar: 'none',         // none | top | left
	pageBorder: 'none',        // none | thin | double | thick
	backgroundPattern: 'none', // none | dots | grid | lines
	dropCap: false,
	radius: 6,
	linkUnderline: true,
};

function makeTemplate(def) {
	const f = def.fonts || {};
	const sizes = def.sizes || [26, 20, 16, 13.5, 12, 11];
	const colours = Object.assign({}, BASE_COLOURS);
	const dc = Object.assign({}, def.colours);
	if (dc.headings) for (let n = 1; n <= 6; n++) colours['h' + n] = dc.headings;
	delete dc.headings;
	Object.assign(colours, dc);
	const headings = {};
	for (let n = 1; n <= 6; n++) {
		headings['h' + n] = Object.assign({
			font: f.heading || '', size: sizes[n - 1], weight: def.headingWeight || 700,
			italic: false, uppercase: false, letterSpacing: 0, align: 'inherit',
		}, (def.headings || {})['h' + n]);
	}
	return {
		id: def.id, name: def.name, category: def.category, description: def.description,
		data: {
			colours,
			typography: {
				body: { font: f.body || 'Inter, "Segoe UI", Helvetica, Arial, sans-serif', size: def.bodySize || 11, weight: 400, lineHeight: def.lineHeight || 1.6 },
				headings,
				code: { font: f.code || '"JetBrains Mono", Consolas, monospace', size: def.codeSize || 9.5 },
			},
			design: Object.assign({}, BASE_DESIGN, def.design),
		},
	};
}

const TEMPLATES = [
	makeTemplate({
		id: 'minimal', name: 'Minimal', category: 'professional', description: 'Clean, quiet and neutral.',
	}),
	makeTemplate({
		id: 'editorial', name: 'Editorial', category: 'creative', description: 'Magazine serif with drop caps.',
		fonts: { heading: '"Playfair Display", Georgia, serif', body: 'Georgia, "Times New Roman", serif', code: '"Courier New", monospace' },
		sizes: [30, 22, 17, 14, 12, 11], bodySize: 11.5, lineHeight: 1.7,
		colours: {
			text: '#2b2b2b', background: '#fffdf8', headings: '#1a1a1a', h1: '#7f1d1d', accent: '#9f1239', link: '#9f1239',
			blockquoteBorder: '#9f1239', blockquoteBg: '#fdf2f4', blockquoteText: '#5b4b4b', tableHeaderBg: '#f5ece6',
			hr: '#d6c7bd', codeBg: '#f5efe6', codeText: '#7f1d1d', headerFooter: '#8a7f76', tableBorder: '#d6c7bd', tableStripe: '#fbf6f0',
		},
		headings: { h1: { align: 'center' }, h2: { italic: true } },
		design: { headingStyle: 'centered', decorationLevels: 1, textAlign: 'justify', dropCap: true, blockquoteStyle: 'quote', tableStyle: 'minimal', imageStyle: 'framed' },
	}),
	makeTemplate({
		id: 'corporate', name: 'Corporate', category: 'professional', description: 'Navy banding for reports.',
		fonts: { heading: '"Segoe UI", Calibri, Arial, sans-serif', body: '"Segoe UI", Calibri, Arial, sans-serif' },
		colours: {
			headings: '#1e3a8a', h2: '#1d4ed8', accent: '#1d4ed8', link: '#1d4ed8', tableHeaderBg: '#1e3a8a', tableHeaderText: '#ffffff',
			tableStripe: '#eff6ff', tableBorder: '#cbd5e1', blockquoteBorder: '#1d4ed8', blockquoteBg: '#eff6ff', codeBg: '#f1f5f9', codeText: '#0f172a',
		},
		design: { headingStyle: 'band', decorationLevels: 1, tableStyle: 'header', accentBar: 'top', listMarkerAccent: true },
	}),
	makeTemplate({
		id: 'academic', name: 'Academic', category: 'academic', description: 'Times, justified, paper-ready.',
		fonts: { heading: '"Times New Roman", Times, serif', body: '"Times New Roman", Times, serif', code: '"Courier New", monospace' },
		sizes: [18, 15, 13, 12, 12, 12], bodySize: 12, lineHeight: 1.8,
		colours: {
			text: '#111111', headings: '#000000', accent: '#000000', link: '#1a0dab', blockquoteBorder: '#999999', blockquoteBg: '#ffffff',
			blockquoteText: '#333333', tableHeaderBg: '#ffffff', tableBorder: '#555555', tableStripe: '#ffffff', codeBg: '#f6f6f6', codeText: '#111111', headerFooter: '#555555',
		},
		headings: { h1: { align: 'center' }, h3: { italic: true }, h4: { italic: true, weight: 400 } },
		design: { textAlign: 'justify', tableStyle: 'minimal', blockquoteStyle: 'minimal', paragraphSpacing: 0.6 },
	}),
	makeTemplate({
		id: 'modern', name: 'Modern Bold', category: 'creative', description: 'Uppercase violet statements.',
		fonts: { heading: 'Montserrat, "Segoe UI", sans-serif', body: '"Open Sans", "Segoe UI", sans-serif' },
		sizes: [28, 18, 15, 13, 12, 11], headingWeight: 800,
		colours: {
			text: '#1f1f2e', headings: '#1f1f2e', h1: '#6d28d9', h2: '#7c3aed', accent: '#7c3aed', link: '#7c3aed', highlight: '#ede9fe',
			blockquoteBorder: '#7c3aed', blockquoteBg: '#f5f3ff', tableHeaderBg: '#7c3aed', tableHeaderText: '#ffffff', tableStripe: '#faf5ff',
			tableBorder: '#ddd6fe', codeBg: '#f5f3ff', codeText: '#6d28d9',
		},
		headings: { h1: { uppercase: true, letterSpacing: 0.04 }, h2: { uppercase: true, letterSpacing: 0.03 } },
		design: { headingStyle: 'leftbar', decorationLevels: 2, tableStyle: 'header', listMarkerAccent: true, imageStyle: 'rounded', radius: 10, blockquoteStyle: 'card' },
	}),
	makeTemplate({
		id: 'pastel', name: 'Pastel Notebook', category: 'creative', description: 'Soft colours on dotted paper.',
		fonts: { heading: 'Nunito, "Segoe UI", sans-serif', body: 'Nunito, "Segoe UI", sans-serif' },
		colours: {
			background: '#fffafc', text: '#4a4453', h1: '#d6336c', h2: '#9c36b5', h3: '#4263eb', h4: '#0c8599', h5: '#5c5f66', h6: '#5c5f66',
			accent: '#f06595', link: '#ae3ec9', highlight: '#ffe3ec', blockquoteBorder: '#faa2c1', blockquoteBg: '#fff0f6', blockquoteText: '#6b5b73',
			tableHeaderBg: '#f3d9fa', tableHeaderText: '#5f3dc4', tableStripe: '#fdf4ff', tableBorder: '#eebefa', codeBg: '#fff0f6', codeText: '#c2255c',
			hr: '#fcc2d7', headerFooter: '#b197c4',
		},
		design: { headingStyle: 'pill', decorationLevels: 1, backgroundPattern: 'dots', imageStyle: 'rounded', radius: 14, blockquoteStyle: 'card', tableStyle: 'striped', listMarkerAccent: true },
	}),
	makeTemplate({
		id: 'midnight', name: 'Midnight', category: 'dark', description: 'Dark navy with neon accents.',
		fonts: { heading: '"Space Grotesk", Inter, sans-serif', body: 'Inter, "Segoe UI", sans-serif' },
		colours: {
			background: '#0f172a', text: '#cbd5e1', h1: '#38bdf8', h2: '#7dd3fc', h3: '#a5b4fc', h4: '#c4b5fd', h5: '#e2e8f0', h6: '#e2e8f0',
			accent: '#38bdf8', link: '#22d3ee', highlight: '#854d0e', blockquoteText: '#94a3b8', blockquoteBorder: '#38bdf8', blockquoteBg: '#1e293b',
			codeText: '#f9a8d4', codeBg: '#1e293b', tableHeaderBg: '#1e293b', tableHeaderText: '#e2e8f0', tableBorder: '#334155', tableStripe: '#131d35',
			hr: '#334155', headerFooter: '#64748b',
		},
		design: { headingStyle: 'underline', decorationLevels: 1, tableStyle: 'striped', blockquoteStyle: 'card', listMarkerAccent: true },
	}),
	makeTemplate({
		id: 'typewriter', name: 'Typewriter', category: 'creative', description: 'Sepia paper, monospaced charm.',
		fonts: { heading: '"Courier New", Courier, monospace', body: '"Courier New", Courier, monospace', code: '"Courier New", Courier, monospace' },
		sizes: [22, 17, 14, 12, 11, 11],
		colours: {
			background: '#f4ecd8', text: '#3b2f2f', headings: '#2b1d14', accent: '#8b4513', link: '#8b4513', highlight: '#e9d8a6',
			blockquoteBorder: '#8b4513', blockquoteBg: '#efe3c8', blockquoteText: '#5a4636', codeBg: '#ebdfc4', codeText: '#5a3825',
			tableHeaderBg: '#e6d5b1', tableHeaderText: '#2b1d14', tableBorder: '#b9a27a', tableStripe: '#efe5cd', hr: '#b9a27a', headerFooter: '#7a6650',
		},
		headings: { h1: { uppercase: true, letterSpacing: 0.08 } },
		design: { headingStyle: 'double-underline', decorationLevels: 2, pageBorder: 'double', blockquoteStyle: 'minimal', tableStyle: 'grid' },
	}),
	makeTemplate({
		id: 'botanical', name: 'Botanical', category: 'creative', description: 'Calm greens and serif headings.',
		fonts: { heading: 'Lora, Georgia, serif', body: '"Source Sans 3", "Segoe UI", sans-serif' },
		colours: {
			background: '#fbfaf5', text: '#2f3a2f', h1: '#2d6a4f', h2: '#40916c', h3: '#52796f', h4: '#354f52', h5: '#354f52', h6: '#354f52',
			accent: '#52b788', link: '#2d6a4f', highlight: '#d8f3dc', blockquoteBorder: '#74c69d', blockquoteBg: '#f0f7f2', blockquoteText: '#40514a',
			tableHeaderBg: '#d8f3dc', tableHeaderText: '#1b4332', tableStripe: '#f3faf5', tableBorder: '#b7e4c7', codeBg: '#eef5ef', codeText: '#2d6a4f',
			hr: '#b7e4c7', headerFooter: '#74796d',
		},
		design: { headingStyle: 'underline', decorationLevels: 2, imageStyle: 'rounded', tableStyle: 'striped', accentBar: 'left' },
	}),
	makeTemplate({
		id: 'tech', name: 'Tech Docs', category: 'professional', description: 'Developer docs on a faint grid.',
		fonts: { heading: '"IBM Plex Sans", Inter, sans-serif', body: 'Inter, "Segoe UI", sans-serif', code: '"JetBrains Mono", "Fira Code", monospace' },
		colours: {
			headings: '#0f172a', h3: '#334155', h4: '#334155', accent: '#0ea5e9', link: '#0284c7', codeBg: '#eef2f7', codeText: '#0f172a',
			blockquoteBorder: '#0ea5e9', blockquoteBg: '#f0f9ff', tableHeaderBg: '#f1f5f9', tableBorder: '#e2e8f0', tableStripe: '#f8fafc',
		},
		design: { headingStyle: 'underline', decorationLevels: 2, backgroundPattern: 'grid', tableStyle: 'grid', blockquoteStyle: 'card', imageStyle: 'shadow' },
	}),
	makeTemplate({
		id: 'sunset', name: 'Sunset', category: 'creative', description: 'Warm orange-to-pink energy.',
		fonts: { heading: 'Poppins, "Segoe UI", sans-serif', body: 'Poppins, "Segoe UI", sans-serif' },
		bodySize: 10.5,
		colours: {
			h1: '#e8590c', h2: '#f76707', h3: '#e64980', h4: '#d6336c', h5: '#862e9c', h6: '#862e9c', accent: '#fd7e14', link: '#e8590c',
			highlight: '#ffe8cc', blockquoteBorder: '#fd7e14', blockquoteBg: '#fff4e6', tableHeaderBg: '#fd7e14', tableHeaderText: '#ffffff',
			tableStripe: '#fff9f2', tableBorder: '#ffd8a8', codeBg: '#fff4e6', codeText: '#d9480f', hr: '#ffc078',
		},
		design: { headingStyle: 'overline', decorationLevels: 2, tableStyle: 'header', imageStyle: 'shadow', accentBar: 'top', listMarkerAccent: true, blockquoteStyle: 'card' },
	}),
	makeTemplate({
		id: 'swiss', name: 'Swiss', category: 'professional', description: 'Bold Helvetica with red accents.',
		fonts: { heading: '"Helvetica Neue", Helvetica, Arial, sans-serif', body: '"Helvetica Neue", Helvetica, Arial, sans-serif' },
		sizes: [34, 20, 15, 12, 11, 11], headingWeight: 800,
		colours: {
			text: '#111111', headings: '#000000', accent: '#e10600', link: '#e10600', blockquoteBorder: '#e10600', blockquoteBg: '#ffffff',
			tableHeaderBg: '#111111', tableHeaderText: '#ffffff', tableBorder: '#111111', tableStripe: '#f4f4f4', hr: '#111111', codeBg: '#f2f2f2', codeText: '#111111',
		},
		headings: { h1: { letterSpacing: -0.02 }, h4: { uppercase: true, letterSpacing: 0.08 }, h5: { uppercase: true, letterSpacing: 0.06 }, h6: { uppercase: true, letterSpacing: 0.06 } },
		design: { headingStyle: 'underline', decorationLevels: 1, tableStyle: 'minimal', blockquoteStyle: 'minimal', listMarkerAccent: true },
	}),
];

/** Neutral Office-like styling used for Word/PowerPoint when "use PDF style" is off. */
const PLAIN_STYLE = makeTemplate({
	id: 'plain', name: 'Plain', category: 'hidden', description: '',
	fonts: { heading: 'Calibri Light, sans-serif', body: 'Calibri, sans-serif', code: 'Consolas, monospace' },
	colours: { text: '#000000', headings: '#2f5496', h1: '#1f3864', accent: '#2f5496', link: '#0563c1', codeText: '#000000' },
}).data;

const TEMPLATE_CATEGORIES = [
	['all', 'All'], ['professional', 'Professional'], ['creative', 'Creative'], ['academic', 'Academic'], ['dark', 'Dark'], ['mine', 'My presets'],
];

const getTemplate = (id) => TEMPLATES.find((t) => t.id === id);

/* ────────────────────────────────────────────────────────────────────────── *
 * Settings
 * ────────────────────────────────────────────────────────────────────────── */

const DEFAULT_SETTINGS = {
	enabled: true,
	activeTemplate: 'minimal',
	templateModified: false,
	menu: { pdf: true, docx: true, pptx: true, fileExplorer: false },
	output: { location: 'same', folder: 'Exports', filename: '{title}', overwrite: true, openAfter: false },
	info: { author: '', dateFormat: 'D MMMM YYYY' },
	content: { resolveEmbeds: true },
	colours: clone(TEMPLATES[0].data.colours),
	typography: clone(TEMPLATES[0].data.typography),
	design: clone(TEMPLATES[0].data.design),
	page: {
		controlMargins: true,
		margins: { top: 20, right: 18, bottom: 20, left: 18 },
		h1PageBreak: false,
		avoidBreaks: true,
		imageMaxWidth: 100,
		showLinkUrls: false,
	},
	headerFooter: {
		enabled: false,
		header: { left: '{title}', center: '', right: '{date}' },
		footer: { left: '{author}', center: '', right: '' },
		font: '',
		size: 8.5,
		separator: true,
		showOnFirstPage: true,
	},
	pageNumbers: { enabled: true, position: 'bottom-right', format: '{page} / {pages}', style: 'decimal', showOnFirstPage: true },
	watermark: { enabled: false, text: 'DRAFT', colour: '#888888', opacity: 0.08, angle: -30, size: 96 },
	docx: { useTheme: true, pageSize: 'A4', includeTitle: true, toc: false, headerFooter: true, pageNumbers: true, embedImages: true, pageBackground: false },
	pptx: {
		useTheme: true, aspect: '16:9', titleSize: 36, bodySize: 20,
		background: '#ffffff', titleColour: '#111111', textColour: '#333333', accent: '#2563eb',
		decoration: 'underline', titleSlides: true, splitOn: 'none', slideNumbers: true, footer: '',
	},
	customCss: '',
	presets: [],
};

/** Keys a design template sets. */
const TEMPLATE_KEYS = ['colours', 'typography', 'design'];
/** Keys a user preset captures. */
const PRESET_KEYS = ['colours', 'typography', 'design', 'page', 'headerFooter', 'pageNumbers', 'watermark', 'customCss'];

const SLOT_KEYS = ['top-left', 'top-center', 'top-right', 'bottom-left', 'bottom-center', 'bottom-right'];

/* ────────────────────────────────────────────────────────────────────────── *
 * Tokens, header & footer slots
 * ────────────────────────────────────────────────────────────────────────── */

const TOKENS = [
	['{title}', 'Note title (frontmatter "title" or file name)'],
	['{filename}', 'File name'],
	['{folder}', 'Parent folder'],
	['{date}', 'Export date (format set in General)'],
	['{time}', 'Export time'],
	['{author}', 'Frontmatter "author" or the default author'],
	['{vault}', 'Vault name'],
	['{page}', 'Current page number'],
	['{pages}', 'Total pages'],
	['{fm.key}', 'Any frontmatter field, e.g. {fm.client}'],
];

function expandTokens(str, ctx, pageVals) {
	return String(str || '').replace(/\{([\w.-]+)\}/g, (m, key) => {
		if (key === 'page' || key === 'pages') return pageVals ? String(pageVals[key]) : m;
		if (key.startsWith('fm.')) {
			const v = ctx.frontmatter ? ctx.frontmatter[key.slice(3)] : undefined;
			return v == null ? '' : Array.isArray(v) ? v.join(', ') : String(v);
		}
		return Object.prototype.hasOwnProperty.call(ctx.tokens, key) ? ctx.tokens[key] : m;
	});
}

const SAMPLE_CTX = {
	frontmatter: {},
	tokens: { title: 'Quarterly Field Notes', filename: 'Quarterly Field Notes', folder: 'Projects', date: '19 October 2026', time: '09:30', author: 'A. Writer', vault: 'My Vault' },
};

/** Work out what goes in each of the six page-margin slots (normal pages and the first page). */
function computeSlots(s) {
	const hf = s.headerFooter;
	const pn = s.pageNumbers;
	const map = {
		'top-left': hf.header.left, 'top-center': hf.header.center, 'top-right': hf.header.right,
		'bottom-left': hf.footer.left, 'bottom-center': hf.footer.center, 'bottom-right': hf.footer.right,
	};
	const normal = {};
	const first = {};
	for (const k of SLOT_KEYS) {
		normal[k] = hf.enabled ? map[k] || '' : '';
		first[k] = hf.enabled && hf.showOnFirstPage ? map[k] || '' : '';
	}
	if (pn.enabled && SLOT_KEYS.includes(pn.position)) {
		normal[pn.position] = pn.format;
		first[pn.position] = pn.showOnFirstPage ? pn.format : first[pn.position];
	}
	return { normal, first, differentFirst: SLOT_KEYS.some((k) => normal[k] !== first[k]) };
}

/* ────────────────────────────────────────────────────────────────────────── *
 * CSS generation (shared by the PDF export and the live preview)
 * ────────────────────────────────────────────────────────────────────────── */

function headingDecoration(style, level, c, d) {
	const w = level === 1 ? 3 : 2;
	const acc = c.accent;
	switch (style) {
		case 'underline': return `padding-bottom: .25em !important; border-bottom: ${w}px solid ${acc} !important;`;
		case 'double-underline': return `padding-bottom: .25em !important; border-bottom: ${w + 2}px double ${acc} !important;`;
		case 'overline': return `padding-top: .35em !important; border-top: ${w + 1}px solid ${acc} !important;`;
		case 'leftbar': return `padding-left: .55em !important; border-left: ${w + 3}px solid ${acc} !important;`;
		case 'band': return `padding: .3em .55em !important; background: ${mix(acc, c.background, 0.86)} !important; border-radius: ${d.radius}px;`;
		case 'boxed': return `padding: .3em .6em !important; border: 1.5px solid ${acc} !important; border-radius: ${d.radius}px;`;
		case 'pill': return `display: table !important; padding: .15em .8em !important; background: ${acc} !important; color: ${contrastText(acc)} !important; border-radius: 999px;`;
		case 'centered': return 'text-align: center !important;';
		default: return '';
	}
}

/** Styles for the rendered note itself, scoped under `R`. */
function buildContentCss(s, R) {
	const c = s.colours;
	const t = s.typography;
	const d = s.design;
	const p = s.page;
	const bodyFont = cssFont(t.body.font, 'sans-serif');
	const codeFont = cssFont(t.code.font, 'monospace');
	const L = [];
	const rule = (sel, body) => L.push(`${sel.split(',').map((x) => `${R} ${x.trim()}`).join(', ')} { ${body} }`);

	const vars = [
		`--text-normal: ${c.text}`, `--background-primary: ${c.background}`, `--link-color: ${c.link}`, `--link-external-color: ${c.link}`,
		`--text-accent: ${c.accent}`, `--blockquote-border-color: ${c.blockquoteBorder}`, `--code-background: ${c.codeBg}`, `--code-normal: ${c.codeText}`,
		`--font-text: ${bodyFont}`, `--font-monospace: ${codeFont}`, `--text-highlight-bg: ${c.highlight}`, `--hr-color: ${c.hr}`,
		`--table-border-color: ${c.tableBorder}`, `--table-header-background: ${c.tableHeaderBg}`,
	];
	for (let n = 1; n <= 6; n++) {
		const h = t.headings['h' + n];
		vars.push(`--h${n}-color: ${c['h' + n]}`, `--h${n}-font: ${cssFont(h.font, bodyFont)}`, `--h${n}-size: ${h.size}pt`, `--h${n}-weight: ${h.weight}`);
	}
	L.push(`${R} { ${vars.join('; ')}; color: ${c.text} !important; background-color: ${c.background} !important; font-family: ${bodyFont} !important; font-size: ${t.body.size}pt !important; font-weight: ${t.body.weight}; line-height: ${t.body.lineHeight} !important; }`);
	if (d.pageBorder !== 'none' || d.accentBar === 'left') L.push(`${R} { padding-left: 6mm !important; padding-right: 6mm !important; }`);

	rule('p, li, dd, dt, td, th', `font-family: ${bodyFont}; line-height: ${t.body.lineHeight};`);
	rule('p, li, td, dd', `color: ${c.text} !important; font-size: ${t.body.size}pt !important;`);
	rule('p', `margin: 0 0 ${d.paragraphSpacing}em 0 !important; text-align: ${d.textAlign} !important;`);
	rule('ul, ol', `margin-top: 0 !important; margin-bottom: ${d.paragraphSpacing}em !important;`);
	rule('strong, b', 'font-weight: 700;');

	for (let n = 1; n <= 6; n++) {
		const h = t.headings['h' + n];
		const deco = n <= d.decorationLevels ? headingDecoration(d.headingStyle, n, c, d) : '';
		const align = h.align && h.align !== 'inherit' ? `text-align: ${h.align} !important;` : '';
		rule(`h${n}`, [
			`font-family: ${cssFont(h.font, bodyFont)} !important`, `font-size: ${h.size}pt !important`, `font-weight: ${h.weight} !important`,
			`font-style: ${h.italic ? 'italic' : 'normal'} !important`, `color: ${c['h' + n]} !important`,
			`text-transform: ${h.uppercase ? 'uppercase' : 'none'} !important`, `letter-spacing: ${h.letterSpacing}em !important`,
			'line-height: 1.25 !important', `margin: ${n <= 2 ? 1.1 : 1}em 0 .5em 0 !important`, 'break-after: avoid', 'page-break-after: avoid',
		].join('; ') + `; ${align} ${deco}`);
		if (n <= d.decorationLevels && d.headingStyle === 'centered') {
			rule(`h${n}::after`, `content: ""; display: block; width: 3.2em; height: 2px; margin: .35em auto 0; background: ${c.accent};`);
		}
	}

	rule('a, a.external-link, a.internal-link', `color: ${c.link} !important; text-decoration: ${d.linkUnderline ? 'underline' : 'none'} !important; background-image: none !important;`);
	rule('mark, .cm-highlight', `background-color: ${c.highlight} !important; color: inherit !important; padding: 0 .1em;`);
	rule('code', `font-family: ${codeFont} !important; font-size: ${t.code.size}pt !important; color: ${c.codeText} !important; background-color: ${c.codeBg} !important; padding: .1em .3em; border-radius: ${Math.min(d.radius, 6)}px;`);
	rule('pre', `background-color: ${c.codeBg} !important; border: 1px solid ${mix(c.codeBg, c.text, 0.1)} !important; border-radius: ${d.radius}px !important; padding: .8em 1em !important; white-space: pre-wrap !important; word-break: break-word;`);
	rule('pre code', `background: transparent !important; padding: 0 !important; color: ${c.codeText} !important;`);

	const bq = {
		bar: `border-left: 4px solid ${c.blockquoteBorder} !important; background: ${c.blockquoteBg} !important; padding: .55em 1em !important; border-radius: 0 ${d.radius}px ${d.radius}px 0;`,
		card: `border: none !important; border-left: none !important; background: ${c.blockquoteBg} !important; padding: .9em 1.2em !important; border-radius: ${Math.max(d.radius, 8)}px; box-shadow: inset 0 0 0 1px ${mix(c.blockquoteBg, c.blockquoteBorder, 0.35)};`,
		quote: `border: none !important; border-left: none !important; background: transparent !important; position: relative; padding: .2em 1em .2em 2.4em !important; font-style: italic; font-size: 1.08em;`,
		minimal: `border-left: 2px solid ${c.blockquoteBorder} !important; background: transparent !important; padding: .1em 1em !important; font-style: italic;`,
	}[d.blockquoteStyle] || '';
	rule('blockquote', `color: ${c.blockquoteText} !important; margin: 1em 0 !important; ${bq}`);
	rule('blockquote p', `color: ${c.blockquoteText} !important;`);
	if (d.blockquoteStyle === 'quote') rule('blockquote::before', `content: "\\201C"; position: absolute; left: .1em; top: -.15em; font-size: 3.2em; line-height: 1; color: ${c.accent}; font-family: Georgia, serif; font-style: normal;`);

	rule('table', 'border-collapse: collapse !important; margin: 1em 0 !important;');
	rule('th, td', 'padding: .4em .7em !important; border: none !important;');
	rule('th', `font-weight: 700 !important; color: ${c.tableHeaderText} !important;`);
	switch (d.tableStyle) {
		case 'striped':
			rule('th', `background: ${c.tableHeaderBg} !important; border-bottom: 2px solid ${c.tableBorder} !important;`);
			rule('td', `border-bottom: 1px solid ${c.tableBorder} !important;`);
			rule('tbody tr:nth-child(even) td', `background: ${c.tableStripe} !important;`);
			break;
		case 'minimal':
			rule('th', `background: transparent !important; border-top: 2px solid ${c.text} !important; border-bottom: 1.5px solid ${c.text} !important; color: ${c.text} !important;`);
			rule('td', `border-bottom: 1px solid ${c.tableBorder} !important;`);
			break;
		case 'header':
			rule('th', `background: ${c.tableHeaderBg} !important; border: 1px solid ${c.tableHeaderBg} !important;`);
			rule('td', `border: 1px solid ${c.tableBorder} !important;`);
			rule('tbody tr:nth-child(even) td', `background: ${c.tableStripe} !important;`);
			break;
		default:
			rule('th', `background: ${c.tableHeaderBg} !important; border: 1px solid ${c.tableBorder} !important;`);
			rule('td', `border: 1px solid ${c.tableBorder} !important;`);
	}

	rule('hr', `border: none !important; border-top: 1px solid ${c.hr} !important; margin: 1.5em 0 !important;`);
	const img = {
		rounded: `border-radius: ${Math.max(d.radius, 10)}px;`,
		shadow: `border-radius: ${d.radius}px; box-shadow: 0 4px 16px rgba(0, 0, 0, .18);`,
		framed: `border: 6px solid #ffffff; outline: 1px solid ${c.tableBorder}; box-shadow: 0 2px 8px rgba(0, 0, 0, .12);`,
	}[d.imageStyle] || '';
	rule('img', `max-width: ${p.imageMaxWidth}% !important; height: auto; ${img}`);
	if (d.listMarkerAccent) rule('li::marker, .list-bullet', `color: ${c.accent} !important;`);
	if (d.dropCap) {
		const capFont = cssFont(t.headings.h1.font, bodyFont);
		rule('h1 + p::first-letter, div:has(> h1) + div > p:first-child::first-letter',
			`float: left; font-size: 3.3em; line-height: .85; padding: .06em .08em 0 0; font-family: ${capFont}; color: ${c.accent}; font-weight: 700;`);
	}
	if (p.avoidBreaks) rule('pre, blockquote, table, img, figure, .callout, li', 'break-inside: avoid; page-break-inside: avoid;');
	if (p.h1PageBreak) rule('h1', 'break-before: page; page-break-before: always;');
	if (p.showLinkUrls) rule('a.external-link::after', `content: " (" attr(href) ")"; font-size: .85em; color: ${c.headerFooter};`);
	return L.join('\n');
}

function backgroundPatternCss(s) {
	const c = s.colours;
	const ink = rgba(mix(c.background, c.text, 0.5), 0.18);
	switch (s.design.backgroundPattern) {
		case 'dots': return `background-image: radial-gradient(${ink} 1px, transparent 1.3px) !important; background-size: 16px 16px !important;`;
		case 'grid': return `background-image: linear-gradient(${ink} 1px, transparent 1px), linear-gradient(90deg, ${ink} 1px, transparent 1px) !important; background-size: 20px 20px !important;`;
		case 'lines': return `background-image: repeating-linear-gradient(to bottom, transparent 0, transparent 23px, ${ink} 23px, ${ink} 24px) !important;`;
		default: return '';
	}
}

function pageFrameCss(s) {
	const c = s.colours;
	const border = { thin: `1px solid ${c.accent}`, double: `4px double ${c.accent}`, thick: `3px solid ${c.accent}` }[s.design.pageBorder];
	const shadow = s.design.accentBar === 'top' ? `inset 0 6px 0 ${c.accent}` : s.design.accentBar === 'left' ? `inset 6px 0 0 ${c.accent}` : '';
	if (!border && !shadow) return '';
	return `${border ? `border: ${border};` : ''} ${shadow ? `box-shadow: ${shadow};` : ''}`;
}

function watermarkCss(s) {
	const w = s.watermark;
	return `content: ${cssString(w.text)}; transform: translate(-50%, -50%) rotate(${w.angle}deg); font-size: ${w.size}pt; font-weight: 800; font-family: ${cssFont(s.typography.headings.h1.font || s.typography.body.font)}; color: ${w.colour}; opacity: ${w.opacity}; white-space: nowrap; pointer-events: none; letter-spacing: .05em;`;
}

function cssContent(tpl, ctx, numStyle) {
	const str = expandTokens(tpl, ctx, null);
	const parts = str.split(/(\{page\}|\{pages\})/).filter((x) => x !== '');
	if (!parts.length) return '""';
	const style = numStyle && numStyle !== 'decimal' ? `, ${numStyle}` : '';
	return parts.map((x) => (x === '{page}' ? `counter(page${style})` : x === '{pages}' ? `counter(pages${style})` : cssString(x))).join(' ');
}

function marginBoxesCss(slots, s, ctx) {
	const hf = s.headerFooter;
	const font = cssFont(hf.font || s.typography.body.font, 'sans-serif');
	const hasTop = SLOT_KEYS.slice(0, 3).some((k) => slots[k]);
	const hasBottom = SLOT_KEYS.slice(3).some((k) => slots[k]);
	const out = [];
	for (const k of SLOT_KEYS) {
		const top = k.startsWith('top');
		const lineHere = hf.enabled && hf.separator && (top ? hasTop : hasBottom);
		let content = 'none';
		if (slots[k]) content = cssContent(slots[k], ctx, s.pageNumbers.style);
		else if (lineHere) content = '""';
		const align = k.endsWith('left') ? 'left' : k.endsWith('right') ? 'right' : 'center';
		const pos = top ? 'vertical-align: bottom; padding-bottom: 3.5mm;' : 'vertical-align: top; padding-top: 3.5mm;';
		const line = lineHere && content !== 'none' ? `${top ? 'border-bottom' : 'border-top'}: .5pt solid ${s.colours.headerFooter};` : '';
		out.push(`@${k} { content: ${content}; font-family: ${font}; font-size: ${hf.size}pt; color: ${s.colours.headerFooter}; text-align: ${align}; ${pos} ${line} }`);
	}
	return out.join(' ');
}

/** The full stylesheet injected for Obsidian's native PDF export. */
function buildPrintCss(s, ctx) {
	const c = s.colours;
	const m = s.page.margins;
	const out = ['/* PDF Customiser — generated print styles */', '@media print {'];
	out.push(`html, body { background-color: ${c.background} !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }`);
	const pattern = backgroundPatternCss(s);
	if (pattern) out.push(`html { ${pattern} }`);
	out.push(buildContentCss(s, 'body .markdown-rendered'));
	const frame = pageFrameCss(s);
	if (frame) out.push(`body::after { content: ""; position: fixed; top: 0; left: 0; right: 0; bottom: 0; pointer-events: none; z-index: 9998; ${frame} }`);
	if (s.watermark.enabled && s.watermark.text) out.push(`body::before { position: fixed; top: 50%; left: 50%; z-index: 9999; ${watermarkCss(s)} }`);
	if (s.customCss && s.customCss.trim()) out.push(`body .markdown-rendered { ${s.customCss} }`);
	out.push('}');

	const slots = computeSlots(s);
	const margin = s.page.controlMargins ? `margin: ${m.top}mm ${m.right}mm ${m.bottom}mm ${m.left}mm !important;` : '';
	out.push(`@page { ${margin} background-color: ${c.background}; ${marginBoxesCss(slots.normal, s, ctx)} }`);
	if (slots.differentFirst) out.push(`@page :first { ${marginBoxesCss(slots.first, s, ctx)} }`);
	return out.join('\n');
}

/** Page-level decoration for the settings preview (mirrors the print rules). */
function buildPreviewPageCss(s, P) {
	const c = s.colours;
	const out = [`${P} { background-color: ${c.background}; ${backgroundPatternCss(s)} }`];
	const frame = pageFrameCss(s);
	out.push(`${P} .pdfc-pv-frame { display: ${frame ? 'block' : 'none'}; ${frame} }`);
	out.push(`${P} .pdfc-pv-watermark { display: ${s.watermark.enabled && s.watermark.text ? 'block' : 'none'}; }`);
	out.push(`${P} .pdfc-pv-watermark::before { position: absolute; top: 50%; left: 50%; ${watermarkCss(s)} }`);
	return out.join('\n');
}

/* ────────────────────────────────────────────────────────────────────────── *
 * Markdown → block tree (used for Word & PowerPoint)
 * ────────────────────────────────────────────────────────────────────────── */

const IMAGE_EXT = /\.(png|jpe?g|gif|webp|svg|bmp|avif)$/i;
const LIST_RE = /^(\s*)([-*+]|\d{1,9}[.)])(?:\s+(.*))?$/;
const HR_RE = /^\s{0,3}([-*_])(?:\s*\1){2,}\s*$/;
const TABLE_SEP_RE = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;

const isBlank = (l) => /^\s*$/.test(l);
const indentWidth = (s) => s.replace(/\t/g, '    ').length;

function isBlockStart(l) {
	return /^\s{0,3}(#{1,6}\s|>|```|~~~|\$\$)/.test(l) || HR_RE.test(l) || LIST_RE.test(l);
}

function stripFrontmatter(md) {
	return md.replace(/^﻿?---\r?\n[\s\S]*?\r?\n---[ \t]*(\r?\n|$)/, '');
}

function splitTableRow(line) {
	let s = line.trim();
	if (s.startsWith('|')) s = s.slice(1);
	if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1);
	const cells = [];
	let cur = '';
	let inCode = false;
	for (let k = 0; k < s.length; k++) {
		const ch = s[k];
		if (ch === '\\' && s[k + 1] === '|') { cur += '|'; k++; continue; }
		if (ch === '`') inCode = !inCode;
		if (ch === '|' && !inCode) { cells.push(cur.trim()); cur = ''; continue; }
		cur += ch;
	}
	cells.push(cur.trim());
	return cells;
}

/** Parse "path|300", "path|alt|300x200" etc. */
function parseImageTarget(target, alt) {
	const parts = target.split('|');
	const img = { type: 'image', src: parts[0].split('#')[0].trim(), alt: alt || '', width: 0, height: 0 };
	for (const p of parts.slice(1)) {
		const m = /^\s*(\d+)(?:x(\d+))?\s*$/.exec(p);
		if (m) { img.width = +m[1]; img.height = m[2] ? +m[2] : 0; } else if (p.trim()) img.alt = p.trim();
	}
	return img;
}

function imagesOnlyLine(text) {
	const re = /!\[\[([^\]]+)\]\]|!\[([^\]]*)\]\(\s*<?([^)>]+?)>?(?:\s+"[^"]*")?\s*\)/g;
	if (text.replace(re, '').trim()) return null;
	const out = [];
	let m;
	while ((m = re.exec(text))) {
		if (m[1] !== undefined) {
			const img = parseImageTarget(m[1], '');
			if (!IMAGE_EXT.test(img.src)) return null;
			out.push(img);
		} else {
			const altParts = m[2].split('|');
			const img = parseImageTarget(m[3] + (altParts[1] ? '|' + altParts[1] : ''), altParts[0]);
			out.push(img);
		}
	}
	return out.length ? out : null;
}

function parseMarkdown(src) {
	return parseBlocks(String(src).replace(/\r\n?/g, '\n').split('\n'));
}

function parseBlocks(lines) {
	const blocks = [];
	let i = 0;
	while (i < lines.length) {
		const line = lines[i];
		if (isBlank(line)) { i++; continue; }

		// Fenced code
		let m = /^(\s*)(`{3,}|~{3,})\s*([^\s`]*)/.exec(line);
		if (m) {
			const fence = m[2];
			const indent = m[1].length;
			const lang = (m[3] || '').toLowerCase();
			const close = new RegExp('^\\s*' + (fence[0] === '`' ? '`' : '~') + '{' + fence.length + ',}\\s*$');
			const body = [];
			i++;
			while (i < lines.length && !close.test(lines[i])) {
				const l = lines[i];
				const lead = l.length - l.trimStart().length;
				body.push(l.slice(Math.min(indent, lead)));
				i++;
			}
			i++;
			blocks.push(lang === 'math' ? { type: 'math', text: body.join('\n') } : { type: 'code', lang, text: body.join('\n') });
			continue;
		}

		// Display maths
		if (/^\s*\$\$/.test(line)) {
			const first = line.trim().slice(2);
			if (first.endsWith('$$') && first.length >= 2) { blocks.push({ type: 'math', text: first.slice(0, -2).trim() }); i++; continue; }
			const body = first ? [first] : [];
			i++;
			while (i < lines.length && !/\$\$\s*$/.test(lines[i])) body.push(lines[i++]);
			if (i < lines.length) { const last = lines[i].replace(/\$\$\s*$/, ''); if (last.trim()) body.push(last); i++; }
			blocks.push({ type: 'math', text: body.join('\n').trim() });
			continue;
		}

		// ATX heading
		m = /^\s{0,3}(#{1,6})\s+(.*?)(?:\s+#+)?\s*$/.exec(line);
		if (m) { blocks.push({ type: 'heading', level: m[1].length, text: m[2] }); i++; continue; }

		if (HR_RE.test(line)) { blocks.push({ type: 'hr' }); i++; continue; }

		// Table
		if (line.includes('|') && i + 1 < lines.length && TABLE_SEP_RE.test(lines[i + 1]) && lines[i + 1].includes('-')) {
			const header = splitTableRow(line);
			const align = splitTableRow(lines[i + 1]).map((c) => (/^:-+:$/.test(c) ? 'center' : /^-+:$/.test(c) ? 'right' : 'left'));
			const rows = [];
			i += 2;
			while (i < lines.length && !isBlank(lines[i]) && lines[i].includes('|')) rows.push(splitTableRow(lines[i++]));
			blocks.push({ type: 'table', header, align, rows });
			continue;
		}

		// Blockquote / callout
		if (/^\s{0,3}>/.test(line)) {
			const inner = [];
			while (i < lines.length && /^\s{0,3}>/.test(lines[i])) inner.push(lines[i++].replace(/^\s{0,3}>\s?/, ''));
			const cm = inner.length ? /^\[!([\w-]+)\][+-]?\s*(.*)$/.exec(inner[0].trim()) : null;
			if (cm) {
				const kind = cm[1].toLowerCase();
				blocks.push({ type: 'callout', kind, title: cm[2] || kind.charAt(0).toUpperCase() + kind.slice(1), children: parseBlocks(inner.slice(1)) });
			} else {
				blocks.push({ type: 'blockquote', children: parseBlocks(inner) });
			}
			continue;
		}

		// List
		if (LIST_RE.test(line)) {
			const res = parseList(lines, i);
			blocks.push(res.block);
			i = res.next;
			continue;
		}

		// Paragraph
		const para = [];
		while (i < lines.length && !isBlank(lines[i])) {
			if (para.length && isBlockStart(lines[i])) break;
			if (para.length && lines[i].includes('|') && i + 1 < lines.length && TABLE_SEP_RE.test(lines[i + 1]) && lines[i + 1].includes('-')) break;
			para.push(lines[i].trim());
			i++;
		}
		const text = para.join('\n');
		const imgs = imagesOnlyLine(text.replace(/\n/g, ' '));
		if (imgs) blocks.push(...imgs);
		else blocks.push({ type: 'paragraph', text });
	}
	return blocks;
}

function parseList(lines, start) {
	const flat = [];
	let i = start;
	while (i < lines.length) {
		const l = lines[i];
		const m = LIST_RE.exec(l);
		if (m && !HR_RE.test(l)) {
			// A switch between bullets and numbers at the top level starts a new list.
			if (flat.length && indentWidth(m[1]) <= flat[0].indent && /\d/.test(m[2]) !== flat[0].ordered) break;
			let text = m[3] || '';
			let task = null;
			const t = /^\[([ xX/-])\](?:\s+(.*))?$/.exec(text);
			if (t) { task = t[1] !== ' '; text = t[2] || ''; }
			flat.push({ indent: indentWidth(m[1]), ordered: /\d/.test(m[2]), num: parseInt(m[2], 10) || 1, text, task });
			i++;
			continue;
		}
		if (isBlank(l)) {
			let j = i;
			while (j < lines.length && isBlank(lines[j])) j++;
			if (j < lines.length && (LIST_RE.test(lines[j]) && !HR_RE.test(lines[j]) || /^\s{2,}\S/.test(lines[j]))) { i = j; continue; }
			break;
		}
		if (flat.length && (/^\s+\S/.test(l) || !isBlockStart(l))) {
			flat[flat.length - 1].text += '\n' + l.trim();
			i++;
			continue;
		}
		break;
	}
	const root = { type: 'list', ordered: flat[0].ordered, start: flat[0].num, items: [] };
	const stack = [{ indent: flat[0].indent, list: root }];
	for (const it of flat) {
		while (stack.length > 1 && it.indent < stack[stack.length - 1].indent) stack.pop();
		let top = stack[stack.length - 1];
		if (it.indent > top.indent && top.list.items.length) {
			const parent = top.list.items[top.list.items.length - 1];
			const sub = { type: 'list', ordered: it.ordered, start: it.num, items: [] };
			parent.children.push(sub);
			stack.push({ indent: it.indent, list: sub });
			top = stack[stack.length - 1];
		}
		top.list.items.push({ text: it.text, task: it.task, children: [] });
	}
	return { block: root, next: i };
}

/* Inline ------------------------------------------------------------------- */

const INLINE_DELIMS = [
	['***', { bold: true, italic: true }], ['**', { bold: true }], ['__', { bold: true }],
	['~~', { strike: true }], ['==', { highlight: true }], ['*', { italic: true }], ['_', { italic: true }],
];

function findClose(text, delim, from) {
	let j = from;
	for (;;) {
		j = text.indexOf(delim, j);
		if (j < 0) return -1;
		if (j === from || /\s/.test(text[j - 1])) { j += 1; continue; }
		if (delim.length === 1) {
			if (text[j + 1] === delim) { j += 2; continue; }
			if (delim === '_' && /[A-Za-z0-9]/.test(text[j + 1] || '')) { j += 1; continue; }
		}
		return j;
	}
}

/** Parse inline Markdown into styled runs: { text, bold, italic, strike, highlight, code, math, link, internal, sup } | { br: true } */
function parseInline(text, style = {}) {
	const runs = [];
	let buf = '';
	let i = 0;
	const flush = () => { if (buf) { runs.push(Object.assign({}, style, { text: buf })); buf = ''; } };
	const s = String(text || '');
	outer:
	while (i < s.length) {
		const c = s[i];
		const rest = s.slice(i);
		let m;
		if (c === '\\' && i + 1 < s.length && /[\\`*_{}[\]()#+\-.!~=|<>$]/.test(s[i + 1])) { buf += s[i + 1]; i += 2; continue; }
		if (c === '\n') { flush(); runs.push({ br: true }); i++; continue; }
		if (c === '`' && (m = /^(`+)([\s\S]*?[^`])\1(?!`)/.exec(rest))) {
			flush(); runs.push(Object.assign({}, style, { code: true, text: m[2] })); i += m[0].length; continue;
		}
		if (c === '$' && s[i + 1] !== '$' && (m = /^\$([^\s$](?:[^$\n]*[^\s$])?)\$(?!\d)/.exec(rest))) {
			flush(); runs.push(Object.assign({}, style, { math: true, text: m[1] })); i += m[0].length; continue;
		}
		if (c === '!' && (m = /^!\[\[([^\]]+)\]\]/.exec(rest))) {
			const img = parseImageTarget(m[1], '');
			buf += img.alt && !/^\d/.test(img.alt) ? img.alt : '';
			i += m[0].length; continue;
		}
		if (c === '!' && (m = /^!\[([^\]]*)\]\([^)]*\)/.exec(rest))) { buf += m[1].split('|')[0]; i += m[0].length; continue; }
		if (c === '[' && (m = /^\[\[([^\]|]+)(?:\|([^\]]+))?\]\]/.exec(rest))) {
			flush();
			const label = m[2] || m[1].replace(/#\^?/g, ' › ').replace(/^ › /, '');
			runs.push(Object.assign({}, style, { text: label, internal: true }));
			i += m[0].length; continue;
		}
		if (c === '[' && (m = /^\[\^([^\]]+)\]/.exec(rest))) { flush(); runs.push(Object.assign({}, style, { text: m[1], sup: true })); i += m[0].length; continue; }
		if (c === '[' && (m = /^\[((?:[^\][]|\[[^\]]*\])+)\]\(\s*<?([^)\s>]+)>?(?:\s+"[^"]*")?\s*\)/.exec(rest))) {
			flush(); runs.push(...parseInline(m[1], Object.assign({}, style, { link: m[2] }))); i += m[0].length; continue;
		}
		if (c === '<') {
			if ((m = /^<(https?:\/\/[^>\s]+)>/.exec(rest))) { flush(); runs.push(Object.assign({}, style, { text: m[1], link: m[1] })); i += m[0].length; continue; }
			if ((m = /^<br\s*\/?>/i.exec(rest))) { flush(); runs.push({ br: true }); i += m[0].length; continue; }
			if ((m = /^<\/?[a-zA-Z][^>]*>/.exec(rest))) { i += m[0].length; continue; }
		}
		if (c === 'h' && (i === 0 || /[\s(]/.test(s[i - 1])) && (m = /^https?:\/\/[^\s<>()]*[^\s<>().,;:!?'"]/.exec(rest))) {
			flush(); runs.push(Object.assign({}, style, { text: m[0], link: m[0] })); i += m[0].length; continue;
		}
		for (const [d, flags] of INLINE_DELIMS) {
			if (!s.startsWith(d, i)) continue;
			const after = s[i + d.length];
			if (!after || /\s/.test(after)) break;
			if (d[0] === '_' && i > 0 && /[A-Za-z0-9]/.test(s[i - 1])) break;
			const j = findClose(s, d, i + d.length);
			if (j < 0) break;
			flush();
			runs.push(...parseInline(s.slice(i + d.length, j), Object.assign({}, style, flags)));
			i = j + d.length;
			continue outer;
		}
		buf += c;
		i++;
	}
	flush();
	return runs;
}

const runsText = (runs) => runs.map((r) => (r.br ? '\n' : r.text || '')).join('');

/** Pull a heading section (or ^block) out of a note for ![[Note#Section]] embeds. */
function extractSection(content, sub) {
	const lines = content.split('\n');
	if (sub.startsWith('^')) {
		const id = sub.slice(1);
		const line = lines.find((l) => new RegExp('\\s\\^' + id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*$').test(l));
		return line || '';
	}
	const want = sub.trim().toLowerCase();
	const startIdx = lines.findIndex((l) => { const m = /^(#{1,6})\s+(.*?)\s*#*\s*$/.exec(l); return m && m[2].trim().toLowerCase() === want; });
	if (startIdx < 0) return content;
	const level = /^(#+)/.exec(lines[startIdx])[1].length;
	let end = startIdx + 1;
	while (end < lines.length) {
		const m = /^(#{1,6})\s/.exec(lines[end]);
		if (m && m[1].length <= level) break;
		end++;
	}
	return lines.slice(startIdx, end).join('\n');
}

/** Split Markdown into slides on `---` lines (outside code fences), optionally on headings too. */
function splitSlides(md, splitOn) {
	const lines = md.replace(/\r\n?/g, '\n').split('\n');
	const slides = [[]];
	let fence = null;
	const maxLevel = splitOn === 'h1' ? 1 : splitOn === 'h2' ? 2 : 0;
	for (const l of lines) {
		const fm = /^\s*(`{3,}|~{3,})/.exec(l);
		if (fm) {
			if (!fence) fence = fm[1];
			else if (fm[1][0] === fence[0] && fm[1].length >= fence.length && /^\s*[`~]+\s*$/.test(l)) fence = null;
		}
		if (!fence && !fm && /^\s{0,3}-{3,}\s*$/.test(l)) { slides.push([]); continue; }
		if (!fence && maxLevel) {
			const hm = /^(#{1,6})\s/.exec(l);
			if (hm && hm[1].length <= maxLevel && slides[slides.length - 1].some((x) => x.trim())) slides.push([]);
		}
		slides[slides.length - 1].push(l);
	}
	return slides.map((x) => x.join('\n')).filter((x) => x.trim());
}

/* ────────────────────────────────────────────────────────────────────────── *
 * ZIP writer (for .docx / .pptx containers)
 * ────────────────────────────────────────────────────────────────────────── */

const CRC_TABLE = (() => {
	const t = new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		t[n] = c >>> 0;
	}
	return t;
})();

function crc32(bytes) {
	let c = 0xffffffff;
	for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
	return (c ^ 0xffffffff) >>> 0;
}

async function deflateRaw(data) {
	if (typeof CompressionStream === 'undefined') return null;
	try {
		const stream = new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));
		return new Uint8Array(await new Response(stream).arrayBuffer());
	} catch (e) {
		return null;
	}
}

async function createZip(files) {
	const enc = new TextEncoder();
	const chunks = [];
	const central = [];
	let offset = 0;
	const now = new Date();
	const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
	const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
	for (const f of files) {
		const name = enc.encode(f.name);
		const raw = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
		const crc = crc32(raw);
		let method = 0;
		let data = raw;
		if (!f.store && raw.length > 256) {
			const z = await deflateRaw(raw);
			if (z && z.length < raw.length) { data = z; method = 8; }
		}
		const lh = new DataView(new ArrayBuffer(30));
		lh.setUint32(0, 0x04034b50, true);
		lh.setUint16(4, 20, true);
		lh.setUint16(6, 0x0800, true);
		lh.setUint16(8, method, true);
		lh.setUint16(10, dosTime, true);
		lh.setUint16(12, dosDate, true);
		lh.setUint32(14, crc, true);
		lh.setUint32(18, data.length, true);
		lh.setUint32(22, raw.length, true);
		lh.setUint16(26, name.length, true);
		lh.setUint16(28, 0, true);
		chunks.push(new Uint8Array(lh.buffer), name, data);

		const ch = new DataView(new ArrayBuffer(46));
		ch.setUint32(0, 0x02014b50, true);
		ch.setUint16(4, 20, true);
		ch.setUint16(6, 20, true);
		ch.setUint16(8, 0x0800, true);
		ch.setUint16(10, method, true);
		ch.setUint16(12, dosTime, true);
		ch.setUint16(14, dosDate, true);
		ch.setUint32(16, crc, true);
		ch.setUint32(20, data.length, true);
		ch.setUint32(24, raw.length, true);
		ch.setUint16(28, name.length, true);
		ch.setUint32(42, offset, true);
		central.push(new Uint8Array(ch.buffer), name);
		offset += 30 + name.length + data.length;
	}
	const cdSize = central.reduce((n, c) => n + c.length, 0);
	const end = new DataView(new ArrayBuffer(22));
	end.setUint32(0, 0x06054b50, true);
	end.setUint16(8, files.length, true);
	end.setUint16(10, files.length, true);
	end.setUint32(12, cdSize, true);
	end.setUint32(16, offset, true);
	const all = [...chunks, ...central, new Uint8Array(end.buffer)];
	const out = new Uint8Array(all.reduce((n, c) => n + c.length, 0));
	let p = 0;
	for (const c of all) { out.set(c, p); p += c.length; }
	return out;
}

/* ────────────────────────────────────────────────────────────────────────── *
 * Images
 * ────────────────────────────────────────────────────────────────────────── */

function imageInfo(b) {
	if (!b || b.length < 24) return null;
	if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
		const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
		return { ext: 'png', width: dv.getUint32(16), height: dv.getUint32(20) };
	}
	if (b[0] === 0xff && b[1] === 0xd8) {
		let i = 2;
		while (i + 9 < b.length) {
			if (b[i] !== 0xff) { i++; continue; }
			const marker = b[i + 1];
			if (marker === 0xff) { i++; continue; }
			if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { i += 2; continue; }
			const len = (b[i + 2] << 8) | b[i + 3];
			if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
				return { ext: 'jpeg', height: (b[i + 5] << 8) | b[i + 6], width: (b[i + 7] << 8) | b[i + 8] };
			}
			i += 2 + len;
		}
		return { ext: 'jpeg', width: 800, height: 600 };
	}
	if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { ext: 'gif', width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) };
	return null;
}

function sniffMime(b) {
	const head = new TextDecoder().decode(b.slice(0, 256)).trimStart();
	if (head.startsWith('<svg') || head.startsWith('<?xml')) return 'image/svg+xml';
	if (b[0] === 0x52 && b[1] === 0x49 && b[8] === 0x57 && b[9] === 0x45) return 'image/webp';
	if (b[0] === 0x42 && b[1] === 0x4d) return 'image/bmp';
	return 'image/avif';
}

/** Convert formats Office can't embed (SVG, WebP, BMP, AVIF) to PNG via a canvas. */
async function rasterise(bytes) {
	if (typeof document === 'undefined' || typeof Image === 'undefined') return null;
	const mime = sniffMime(bytes);
	const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
	try {
		const img = new Image();
		await new Promise((res, rej) => { img.onload = res; img.onerror = rej; img.src = url; });
		const w = img.naturalWidth || 800;
		const h = img.naturalHeight || 600;
		const scale = mime === 'image/svg+xml' ? 2 : 1;
		const canvas = document.createElement('canvas');
		canvas.width = w * scale;
		canvas.height = h * scale;
		canvas.getContext('2d').drawImage(img, 0, 0, w * scale, h * scale);
		const blob = await new Promise((r) => canvas.toBlob(r, 'image/png'));
		return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: 'png', width: w, height: h };
	} catch (e) {
		return null;
	} finally {
		URL.revokeObjectURL(url);
	}
}

/* ────────────────────────────────────────────────────────────────────────── *
 * Shared theme for Office exports
 * ────────────────────────────────────────────────────────────────────────── */

function deriveTheme(style) {
	const t = style.typography;
	const body = primaryFont(t.body.font) || 'Calibri';
	const headings = {};
	for (let n = 1; n <= 6; n++) {
		const h = t.headings['h' + n];
		headings[n] = {
			font: primaryFont(h.font) || body, size: h.size, bold: h.weight >= 600, italic: !!h.italic,
			caps: !!h.uppercase, letterSpacing: h.letterSpacing || 0, align: h.align, colour: hx(style.colours['h' + n]),
		};
	}
	const c = {};
	for (const k of Object.keys(style.colours)) c[k] = hx(style.colours[k]);
	return { body: { font: body, size: t.body.size, lineHeight: t.body.lineHeight }, code: { font: primaryFont(t.code.font) || 'Consolas', size: t.code.size }, headings, c };
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function relsXml(rels) {
	return XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		+ rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${escXml(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')
		+ '</Relationships>';
}

function corePropsXml(title, author) {
	const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
	return XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
		+ `<dc:title>${escXml(title)}</dc:title><dc:creator>${escXml(author || 'PDF Customiser')}</dc:creator>`
		+ `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`
		+ '</cp:coreProperties>';
}

const APP_PROPS = XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Obsidian PDF Customiser</Application></Properties>';

const isExternalLink = (url) => /^(https?:|mailto:)/i.test(url || '');

/* ────────────────────────────────────────────────────────────────────────── *
 * Word (.docx) builder
 * ────────────────────────────────────────────────────────────────────────── */

const PAGE_SIZES_TWIPS = { A4: [11906, 16838], Letter: [12240, 15840], Legal: [12240, 20160], A5: [8391, 11906], A3: [16838, 23811] };
const tw = (mm) => Math.round(mm * 56.6929);
const hp = (pt) => Math.round(pt * 2);
const DOCX_NUM_FMT = { decimal: 'decimal', 'lower-roman': 'lowerRoman', 'upper-roman': 'upperRoman', 'lower-alpha': 'lowerLetter', 'upper-alpha': 'upperLetter', 'decimal-leading-zero': 'decimalZero' };

class DocxBuilder {
	constructor(plugin, file, ctx) {
		this.plugin = plugin;
		this.file = file;
		this.ctx = ctx;
		this.settings = plugin.settings;
		this.opts = plugin.settings.docx;
		this.style = this.opts.useTheme ? plugin.settings : PLAIN_STYLE;
		this.th = deriveTheme(this.style);
		this.rels = [];
		this.relSeq = 100;
		this.media = [];
		this.docPr = 1;
		this.orderedNums = [];
		this.emitted = 0;
		const size = PAGE_SIZES_TWIPS[this.opts.pageSize] || PAGE_SIZES_TWIPS.A4;
		const mm = this.settings.page.controlMargins ? this.settings.page.margins : { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 };
		this.page = { w: size[0], h: size[1], mt: tw(mm.top), mr: tw(mm.right), mb: tw(mm.bottom), ml: tw(mm.left) };
		this.contentW = this.page.w - this.page.ml - this.page.mr;
	}

	addRel(type, target, external) {
		const id = 'rId' + this.relSeq++;
		this.rels.push({ id, type: REL + type, target, external });
		return id;
	}

	p(pPr, inner) {
		return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
	}

	textXml(text) {
		return String(text).split('\t').map((seg) => `<w:t xml:space="preserve">${escXml(seg)}</w:t>`).join('<w:tab/>');
	}

	run(r) {
		if (r.br) return '<w:r><w:br/></w:r>';
		if (!r.text) return '';
		const c = this.th.c;
		const ext = r.link && isExternalLink(r.link);
		let rPr = '';
		if (ext) rPr += '<w:rStyle w:val="Hyperlink"/>';
		if (r.code || r.math) { const f = escXml(this.th.code.font); rPr += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`; }
		if (r.bold) rPr += '<w:b/>';
		if (r.italic || r.math) rPr += '<w:i/>';
		if (r.strike) rPr += '<w:strike/>';
		if (r.colour) rPr += `<w:color w:val="${r.colour}"/>`;
		else if (r.code) rPr += `<w:color w:val="${c.codeText}"/>`;
		else if ((r.link && !ext) || r.internal) rPr += `<w:color w:val="${c.link}"/>`;
		if (r.code) rPr += `<w:sz w:val="${hp(this.th.code.size)}"/><w:szCs w:val="${hp(this.th.code.size)}"/>`;
		if (r.highlight) rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${c.highlight}"/>`;
		else if (r.code) rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${c.codeBg}"/>`;
		if (r.sup) rPr += '<w:vertAlign w:val="superscript"/>';
		const xml = `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${this.textXml(r.text)}</w:r>`;
		return ext ? `<w:hyperlink r:id="${this.addRel('hyperlink', r.link, true)}" w:history="1">${xml}</w:hyperlink>` : xml;
	}

	runs(runs) {
		return runs.map((r) => this.run(r)).join('');
	}

	newOrderedNum(level, start) {
		const id = 2 + this.orderedNums.length;
		this.orderedNums.push({ id, level, start });
		return id;
	}

	async blocks(list, ctx) {
		const out = [];
		for (const b of list) {
			const top = !ctx.indent && !ctx.quote;
			switch (b.type) {
				case 'heading': {
					const pb = this.settings.page.h1PageBreak && b.level === 1 && this.emitted > 0 ? '<w:pageBreakBefore/>' : '';
					out.push(this.p(`<w:pStyle w:val="Heading${b.level}"/>${pb}`, this.runs(parseInline(b.text))));
					break;
				}
				case 'paragraph': {
					const pPr = (ctx.quote ? '<w:pStyle w:val="Quote"/>' : '') + (ctx.indent ? `<w:ind w:left="${ctx.indent}"/>` : '');
					out.push(this.p(pPr, this.runs(parseInline(b.text))));
					break;
				}
				case 'list':
					out.push(...await this.listXml(b, ctx, 0));
					break;
				case 'blockquote':
					out.push(...await this.blocks(b.children, Object.assign({}, ctx, { quote: true })));
					break;
				case 'callout': {
					const icon = { warning: '⚠', caution: '⚠', danger: '⛔', error: '⛔', tip: '💡', hint: '💡', important: '❗', question: '❓', success: '✔', check: '✔', info: 'ℹ', note: '✎', example: '▸', quote: '❝', bug: '🐞', abstract: '≡', summary: '≡', todo: '☐', failure: '✘' }[b.kind] || '✎';
					out.push(this.p('<w:pStyle w:val="Quote"/><w:keepNext/>', this.runs([{ text: `${icon} ${b.title}`, bold: true }])));
					out.push(...await this.blocks(b.children, Object.assign({}, ctx, { quote: true })));
					break;
				}
				case 'code': {
					const lines = b.text.split('\n');
					const inner = lines.map((ln, i) => (i ? '<w:r><w:br/></w:r>' : '') + (ln ? `<w:r>${this.textXml(ln)}</w:r>` : '')).join('');
					out.push(this.p('<w:pStyle w:val="CodeBlock"/>' + (ctx.indent ? `<w:ind w:left="${ctx.indent}"/>` : ''), inner || '<w:r><w:t></w:t></w:r>'));
					break;
				}
				case 'math':
					out.push(this.p('<w:jc w:val="center"/>', this.runs(b.text.split('\n').flatMap((ln, i) => (i ? [{ br: true }, { text: ln, math: true }] : [{ text: ln, math: true }])))));
					break;
				case 'table':
					out.push(this.tableXml(b));
					break;
				case 'hr':
					out.push(this.p(`<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="${this.th.c.hr}"/></w:pBdr>`, ''));
					break;
				case 'image':
					out.push(await this.imageXml(b));
					break;
				default:
					break;
			}
			if (top) this.emitted++;
		}
		return out;
	}

	async listXml(list, ctx, lvl) {
		const out = [];
		const level = Math.min(lvl, 8);
		const numId = list.ordered ? this.newOrderedNum(level, list.start) : 1;
		for (const item of list.items) {
			const runs = this.runs(parseInline(item.text));
			if (item.task !== null && item.task !== undefined) {
				const box = this.run({ text: item.task ? '☑ ' : '☐ ', colour: this.style.design.listMarkerAccent ? this.th.c.accent : undefined });
				out.push(this.p(`<w:pStyle w:val="ListParagraph"/><w:ind w:left="${720 * (level + 1)}" w:hanging="360"/>`, box + runs));
			} else {
				out.push(this.p(`<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`, runs));
			}
			for (const child of item.children) {
				if (child.type === 'list') out.push(...await this.listXml(child, ctx, level + 1));
				else out.push(...await this.blocks([child], Object.assign({}, ctx, { indent: 720 * (level + 1) })));
			}
		}
		return out;
	}

	tableXml(tbl) {
		const c = this.th.c;
		const ts = this.style.design.tableStyle;
		const cols = Math.max(tbl.header.length, ...tbl.rows.map((r) => r.length));
		const colW = Math.floor(this.contentW / cols);
		const b = (side, val, sz, col) => `<w:${side} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${col}"/>`;
		let borders;
		if (ts === 'grid' || ts === 'header') borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => b(s, 'single', 4, c.tableBorder)).join('');
		else if (ts === 'striped') borders = b('top', 'single', 4, c.tableBorder) + b('bottom', 'single', 4, c.tableBorder) + b('insideH', 'single', 4, c.tableBorder);
		else borders = b('top', 'single', 12, c.text) + b('bottom', 'single', 8, c.text) + b('insideH', 'single', 4, c.tableBorder);
		const rowXml = (cells, ri) => {
			const header = ri === 0;
			let fill = '';
			if (header && ts !== 'minimal') fill = c.tableHeaderBg;
			else if (!header && (ts === 'striped' || ts === 'header') && ri % 2 === 0) fill = c.tableStripe;
			const tcs = [];
			for (let ci = 0; ci < cols; ci++) {
				const runs = parseInline(cells[ci] || '').map((r) => (header ? Object.assign({}, r, { bold: true, colour: r.br ? undefined : ts === 'minimal' ? c.text : c.tableHeaderText }) : r));
				const align = tbl.align[ci] === 'center' ? 'center' : tbl.align[ci] === 'right' ? 'right' : '';
				const tcBorder = header && ts === 'minimal' ? `<w:tcBorders>${b('bottom', 'single', 8, c.text)}</w:tcBorders>` : '';
				tcs.push(`<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/>${tcBorder}${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}</w:tcPr>`
					+ this.p(`<w:spacing w:before="40" w:after="40"/>${align ? `<w:jc w:val="${align}"/>` : ''}`, this.runs(runs)) + '</w:tc>');
			}
			return `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${tcs.join('')}</w:tr>`;
		};
		return '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>'
			+ `<w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/>`
			+ '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
			+ '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
			+ `<w:tblGrid>${Array.from({ length: cols }, () => `<w:gridCol w:w="${colW}"/>`).join('')}</w:tblGrid>`
			+ [tbl.header, ...tbl.rows].map(rowXml).join('')
			+ '</w:tbl>' + this.p('<w:spacing w:after="0"/>', '');
	}

	async imageXml(b) {
		const img = this.opts.embedImages ? await this.plugin.loadImage(b.src, this.file.path) : null;
		if (!img) return this.p('<w:jc w:val="center"/>', this.run({ text: `[Image: ${b.alt || b.src}]`, italic: true, colour: '888888' }));
		const name = `image${this.media.length + 1}.${img.ext}`;
		this.media.push({ name, bytes: img.bytes });
		const rid = this.addRel('image', `media/${name}`);
		let wpx = b.width || img.width;
		let hpx = b.height || (b.width ? (img.height * b.width) / img.width : img.height);
		const maxW = (this.contentW / 1440) * 96 * (this.settings.page.imageMaxWidth / 100);
		if (wpx > maxW) { hpx = (hpx * maxW) / wpx; wpx = maxW; }
		const maxH = ((this.page.h - this.page.mt - this.page.mb) / 1440) * 96 * 0.85;
		if (hpx > maxH) { wpx = (wpx * maxH) / hpx; hpx = maxH; }
		const cx = Math.round(wpx * 9525);
		const cy = Math.round(hpx * 9525);
		const id = this.docPr++;
		const drawing = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
			+ `<wp:docPr id="${id}" name="Picture ${id}" descr="${escXml(b.alt)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS_A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>`
			+ `<a:graphic xmlns:a="${NS_A}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`
			+ `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
			+ `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
		return this.p('<w:keepNext/><w:spacing w:before="120" w:after="120"/><w:jc w:val="center"/>', `<w:r>${drawing}</w:r>`);
	}

	tocXml() {
		return [
			this.p('<w:pStyle w:val="TOCHeading"/>', this.run({ text: 'Contents' })),
			'<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
			+ '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/><w:color w:val="888888"/></w:rPr><w:t>Right-click and choose “Update Field” to build the table of contents.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
			'<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
		];
	}

	tokenRuns(tpl) {
		const str = expandTokens(tpl, this.ctx, null);
		return str.split(/(\{page\}|\{pages\})/).filter(Boolean).map((part) => {
			if (part === '{page}' || part === '{pages}') {
				const instr = part === '{page}' ? 'PAGE' : 'NUMPAGES';
				return `<w:fldSimple w:instr=" ${instr} "><w:r><w:t>1</w:t></w:r></w:fldSimple>`;
			}
			return `<w:r>${this.textXml(part)}</w:r>`;
		}).join('');
	}

	hdrFtrXml(kind, slots, prefix) {
		const tag = kind === 'hdr' ? 'w:hdr' : 'w:ftr';
		const parts = [slots[prefix + '-left'], slots[prefix + '-center'], slots[prefix + '-right']];
		const hf = this.settings.headerFooter;
		const any = parts.some(Boolean);
		const sep = hf.enabled && hf.separator && any
			? `<w:pBdr><w:${kind === 'hdr' ? 'bottom' : 'top'} w:val="single" w:sz="4" w:space="4" w:color="${this.th.c.headerFooter}"/></w:pBdr>` : '';
		const inner = any ? parts.map((t, i) => (i ? '<w:r><w:tab/></w:r>' : '') + this.tokenRuns(t || '')).join('') : '';
		return XML_HEAD + `<${tag} xmlns:w="${NS_W}" xmlns:r="${NS_R}">`
			+ this.p(`<w:pStyle w:val="${kind === 'hdr' ? 'Header' : 'Footer'}"/>${sep}<w:tabs><w:tab w:val="center" w:pos="${Math.round(this.contentW / 2)}"/><w:tab w:val="right" w:pos="${this.contentW}"/></w:tabs>`, inner)
			+ `</${tag}>`;
	}

	stylesXml() {
		const th = this.th;
		const c = th.c;
		const d = this.style.design;
		const sc = this.style.colours;
		const font = (f) => { const e = escXml(f); return `<w:rFonts w:ascii="${e}" w:hAnsi="${e}" w:cs="${e}" w:eastAsia="${e}"/>`; };
		const after = Math.round(d.paragraphSpacing * th.body.size * 20);
		const line = Math.round(th.body.lineHeight * 240);
		const heading = (n) => {
			const h = th.headings[n];
			const deco = n <= d.decorationLevels ? d.headingStyle : 'none';
			const acc = c.accent;
			let pBdr = '';
			let shd = '';
			let jc = '';
			let colour = h.colour;
			switch (deco) {
				case 'underline': pBdr = `<w:bottom w:val="single" w:sz="${n === 1 ? 18 : 12}" w:space="4" w:color="${acc}"/>`; break;
				case 'double-underline': pBdr = `<w:bottom w:val="double" w:sz="6" w:space="4" w:color="${acc}"/>`; break;
				case 'overline': pBdr = `<w:top w:val="single" w:sz="18" w:space="6" w:color="${acc}"/>`; break;
				case 'leftbar': pBdr = `<w:left w:val="single" w:sz="36" w:space="8" w:color="${acc}"/>`; break;
				case 'boxed': pBdr = ['top', 'left', 'bottom', 'right'].map((s) => `<w:${s} w:val="single" w:sz="8" w:space="4" w:color="${acc}"/>`).join(''); break;
				case 'band': shd = `<w:shd w:val="clear" w:color="auto" w:fill="${hx(mix(sc.accent, sc.background, 0.86))}"/>`; break;
				case 'pill': shd = `<w:shd w:val="clear" w:color="auto" w:fill="${acc}"/>`; colour = hx(contrastText(sc.accent)); break;
				case 'centered': jc = '<w:jc w:val="center"/>'; break;
				default: break;
			}
			if (['left', 'center', 'right'].includes(h.align)) jc = `<w:jc w:val="${h.align}"/>`;
			const spacing = `<w:spacing w:before="${Math.round(h.size * 20 * 1.0)}" w:after="${Math.round(h.size * 20 * 0.35)}" w:line="276" w:lineRule="auto"/>`;
			const rPr = font(h.font) + (h.bold ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/>') + (h.italic ? '<w:i/>' : '') + (h.caps ? '<w:caps/>' : '')
				+ `<w:color w:val="${colour}"/>` + (h.letterSpacing ? `<w:spacing w:val="${Math.round(h.letterSpacing * h.size * 20)}"/>` : '')
				+ `<w:sz w:val="${hp(h.size)}"/><w:szCs w:val="${hp(h.size)}"/>`;
			return `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>`
				+ `<w:pPr><w:keepNext/><w:keepLines/>${pBdr ? `<w:pBdr>${pBdr}</w:pBdr>` : ''}${shd}${spacing}${jc}<w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr>${rPr}</w:rPr></w:style>`;
		};
		const h1 = th.headings[1];
		const bqs = d.blockquoteStyle;
		const bqBorder = bqs === 'bar' ? `<w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="${c.blockquoteBorder}"/></w:pBdr>`
			: bqs === 'minimal' ? `<w:pBdr><w:left w:val="single" w:sz="8" w:space="8" w:color="${c.blockquoteBorder}"/></w:pBdr>` : '';
		const bqShd = bqs === 'bar' || bqs === 'card' ? `<w:shd w:val="clear" w:color="auto" w:fill="${c.blockquoteBg}"/>` : '';
		const hfFont = primaryFont(this.settings.headerFooter.font) || th.body.font;
		const hfStyle = (id, name) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>`
			+ `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr><w:rPr>${font(hfFont)}<w:color w:val="${c.headerFooter}"/><w:sz w:val="${hp(this.settings.headerFooter.size)}"/><w:szCs w:val="${hp(this.settings.headerFooter.size)}"/></w:rPr></w:style>`;
		return XML_HEAD + `<w:styles xmlns:w="${NS_W}" xmlns:r="${NS_R}">`
			+ `<w:docDefaults><w:rPrDefault><w:rPr>${font(th.body.font)}<w:color w:val="${c.text}"/><w:sz w:val="${hp(th.body.size)}"/><w:szCs w:val="${hp(th.body.size)}"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>`
			+ `<w:pPrDefault><w:pPr><w:spacing w:after="${after}" w:line="${line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
			+ `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>${d.textAlign === 'justify' ? '<w:pPr><w:jc w:val="both"/></w:pPr>' : ''}</w:style>`
			+ '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
			+ [1, 2, 3, 4, 5, 6].map(heading).join('')
			+ `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>`
			+ `<w:pPr><w:spacing w:after="360"/>${h1.align === 'center' || d.headingStyle === 'centered' ? '<w:jc w:val="center"/>' : ''}</w:pPr>`
			+ `<w:rPr>${font(h1.font)}${h1.bold ? '<w:b/>' : ''}${h1.caps ? '<w:caps/>' : ''}<w:color w:val="${h1.colour}"/><w:sz w:val="${hp(h1.size * 1.3)}"/><w:szCs w:val="${hp(h1.size * 1.3)}"/></w:rPr></w:style>`
			+ `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/>`
			+ `<w:pPr>${bqBorder}${bqShd}<w:spacing w:before="60" w:after="60"/><w:ind w:left="360" w:right="360"/></w:pPr>`
			+ `<w:rPr>${bqs === 'quote' || bqs === 'minimal' ? '<w:i/>' : ''}<w:color w:val="${c.blockquoteText}"/></w:rPr></w:style>`
			+ `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>`
			+ `<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="39"/><w:qFormat/>`
			+ `<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="${c.codeBg}"/><w:spacing w:before="60" w:after="200" w:line="240" w:lineRule="auto"/><w:ind w:left="144" w:right="144"/><w:jc w:val="left"/></w:pPr>`
			+ `<w:rPr>${font(th.code.font)}<w:color w:val="${c.codeText}"/><w:sz w:val="${hp(th.code.size)}"/><w:szCs w:val="${hp(th.code.size)}"/></w:rPr></w:style>`
			+ `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="${c.link}"/>${d.linkUnderline ? '<w:u w:val="single"/>' : ''}</w:rPr></w:style>`
			+ hfStyle('Header', 'header') + hfStyle('Footer', 'footer')
			+ '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>'
			+ '</w:styles>';
	}

	numberingXml() {
		const accent = this.style.design.listMarkerAccent ? `<w:color w:val="${this.th.c.accent}"/>` : '';
		const bullets = ['•', '◦', '▪'];
		const fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
		let bulletLvls = '';
		let orderedLvls = '';
		for (let i = 0; i < 9; i++) {
			const ind = `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>`;
			bulletLvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bullets[i % 3]}"/><w:lvlJc w:val="left"/>${ind}<w:rPr>${accent}</w:rPr></w:lvl>`;
			orderedLvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmts[i % 3]}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/>${ind}<w:rPr>${accent}</w:rPr></w:lvl>`;
		}
		const nums = this.orderedNums.map((n) => `<w:num w:numId="${n.id}"><w:abstractNumId w:val="1"/>`
			+ Array.from({ length: 9 }, (_, l) => `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="${l === n.level ? n.start : 1}"/></w:lvlOverride>`).join('') + '</w:num>').join('');
		return XML_HEAD + `<w:numbering xmlns:w="${NS_W}">`
			+ `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLvls}</w:abstractNum>`
			+ `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${orderedLvls}</w:abstractNum>`
			+ `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>${nums}</w:numbering>`;
	}

	async build(blocks) {
		const body = [];
		if (this.opts.includeTitle) body.push(this.p('<w:pStyle w:val="Title"/>', this.runs(parseInline(this.ctx.tokens.title))));
		if (this.opts.toc) body.push(...this.tocXml());
		body.push(...await this.blocks(blocks, { indent: 0, quote: false }));
		if (!body.length || body[body.length - 1].startsWith('<w:tbl>')) body.push('<w:p/>');

		const s = this.settings;
		const slotSrc = {
			headerFooter: Object.assign({}, s.headerFooter, { enabled: s.headerFooter.enabled && this.opts.headerFooter }),
			pageNumbers: Object.assign({}, s.pageNumbers, { enabled: s.pageNumbers.enabled && this.opts.pageNumbers }),
		};
		const slots = computeSlots(slotSrc);
		const extra = [];
		let refs = '';
		const hasTop = (sl) => SLOT_KEYS.slice(0, 3).some((k) => sl[k]);
		const hasBottom = (sl) => SLOT_KEYS.slice(3).some((k) => sl[k]);
		const addPart = (kind, type, sl, n) => {
			const part = kind === 'hdr' ? 'header' : 'footer';
			const name = `${part}${n}.xml`;
			extra.push({ name: `word/${name}`, data: this.hdrFtrXml(kind, sl, kind === 'hdr' ? 'top' : 'bottom'), ct: `application/vnd.openxmlformats-officedocument.wordprocessingml.${part}+xml` });
			refs += `<w:${part}Reference w:type="${type}" r:id="${this.addRel(part, name)}"/>`;
		};
		if (hasTop(slots.normal)) addPart('hdr', 'default', slots.normal, 1);
		if (slots.differentFirst) addPart('hdr', 'first', slots.first, 2);
		if (hasBottom(slots.normal)) addPart('ftr', 'default', slots.normal, 1);
		if (slots.differentFirst) addPart('ftr', 'first', slots.first, 2);

		const pg = this.page;
		const numFmt = DOCX_NUM_FMT[s.pageNumbers.style] || 'decimal';
		const sectPr = `<w:sectPr>${refs}<w:pgSz w:w="${pg.w}" w:h="${pg.h}"/>`
			+ `<w:pgMar w:top="${pg.mt}" w:right="${pg.mr}" w:bottom="${pg.mb}" w:left="${pg.ml}" w:header="${Math.round(Math.min(708, pg.mt / 2))}" w:footer="${Math.round(Math.min(708, pg.mb / 2))}" w:gutter="0"/>`
			+ `${numFmt !== 'decimal' ? `<w:pgNumType w:fmt="${numFmt}"/>` : ''}${slots.differentFirst ? '<w:titlePg/>' : ''}</w:sectPr>`;
		const bg = this.opts.pageBackground && this.th.c.background !== 'FFFFFF';
		const doc = XML_HEAD + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${NS_A}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`
			+ (bg ? `<w:background w:color="${this.th.c.background}"/>` : '')
			+ `<w:body>${body.join('')}${sectPr}</w:body></w:document>`;

		const settingsXml = XML_HEAD + `<w:settings xmlns:w="${NS_W}">${bg ? '<w:displayBackgroundShape/>' : ''}<w:defaultTabStop w:val="720"/>${this.opts.toc ? '<w:updateFields w:val="true"/>' : ''}`
			+ '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';

		const docRels = [
			{ id: 'rId1', type: REL + 'styles', target: 'styles.xml' },
			{ id: 'rId2', type: REL + 'numbering', target: 'numbering.xml' },
			{ id: 'rId3', type: REL + 'settings', target: 'settings.xml' },
			...this.rels,
		];
		const W = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
		const contentTypes = XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			+ '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
			+ '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>'
			+ `<Override PartName="/word/document.xml" ContentType="${W}document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${W}styles+xml"/>`
			+ `<Override PartName="/word/numbering.xml" ContentType="${W}numbering+xml"/><Override PartName="/word/settings.xml" ContentType="${W}settings+xml"/>`
			+ extra.map((e) => `<Override PartName="/${e.name}" ContentType="${e.ct}"/>`).join('')
			+ '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
			+ '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';

		const files = [
			{ name: '[Content_Types].xml', data: contentTypes },
			{ name: '_rels/.rels', data: relsXml([
				{ id: 'rId1', type: REL + 'officeDocument', target: 'word/document.xml' },
				{ id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' },
				{ id: 'rId3', type: REL + 'extended-properties', target: 'docProps/app.xml' },
			]) },
			{ name: 'docProps/core.xml', data: corePropsXml(this.ctx.tokens.title, this.ctx.tokens.author) },
			{ name: 'docProps/app.xml', data: APP_PROPS },
			{ name: 'word/document.xml', data: doc },
			{ name: 'word/styles.xml', data: this.stylesXml() },
			{ name: 'word/numbering.xml', data: this.numberingXml() },
			{ name: 'word/settings.xml', data: settingsXml },
			{ name: 'word/_rels/document.xml.rels', data: relsXml(docRels) },
			...extra.map((e) => ({ name: e.name, data: e.data })),
			...this.media.map((m) => ({ name: `word/media/${m.name}`, data: m.bytes, store: true })),
		];
		return createZip(files);
	}
}

/* ────────────────────────────────────────────────────────────────────────── *
 * PowerPoint (.pptx) builder
 * ────────────────────────────────────────────────────────────────────────── */

const EMU_PT = 12700;
const SPTREE_HEAD = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const P_NS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
const ri = Math.round;

class PptxBuilder {
	constructor(plugin, file, ctx) {
		this.plugin = plugin;
		this.file = file;
		this.ctx = ctx;
		const s = plugin.settings;
		this.opts = s.pptx;
		const o = this.opts;
		const style = o.useTheme ? s : PLAIN_STYLE;
		this.design = style.design;
		this.th = deriveTheme(style);
		const sc = style.colours;
		if (o.useTheme) {
			this.col = {
				bg: hx(sc.background), title: hx(sc.h1), heading: hx(sc.h2), text: hx(sc.text), accent: hx(sc.accent), link: hx(sc.link),
				muted: hx(sc.headerFooter), code: hx(sc.codeText), quote: hx(sc.blockquoteText), highlight: hx(sc.highlight),
				tableHeaderBg: hx(sc.tableHeaderBg), tableHeaderText: hx(sc.tableHeaderText), tableBorder: hx(sc.tableBorder), tableStripe: hx(sc.tableStripe),
			};
		} else {
			const bg = normHex(o.background, '#ffffff');
			const text = normHex(o.textColour, '#333333');
			const acc = normHex(o.accent, '#2563eb');
			this.col = {
				bg: hx(bg), title: hx(o.titleColour), heading: hx(o.titleColour), text: hx(text), accent: hx(acc), link: hx(acc),
				muted: hx(mix(text, bg, 0.4)), code: hx(text), quote: hx(mix(text, bg, 0.25)), highlight: 'FFF3A3',
				tableHeaderBg: hx(acc), tableHeaderText: hx(contrastText(acc)), tableBorder: hx(mix(text, bg, 0.75)), tableStripe: hx(mix(text, bg, 0.94)),
			};
		}
		this.fonts = { title: this.th.headings[1].font, heading: this.th.headings[2].font, body: this.th.body.font, code: this.th.code.font };
		[this.W, this.H] = o.aspect === '4:3' ? [9144000, 6858000] : [12192000, 6858000];
		this.media = [];
		this.slides = [];
	}

	addRel(type, target, external) {
		const id = 'rId' + this.cur.seq++;
		this.cur.rels.push({ id, type: REL + type, target, external });
		return id;
	}

	nextId() {
		return this.cur.id++;
	}

	runXml(r, size, colour, font, p) {
		const sz = ri(size * 100);
		if (r.br) return `<a:br><a:rPr lang="en-GB" sz="${sz}" dirty="0"/></a:br>`;
		if (!r.text) return '';
		const ext = r.link && isExternalLink(r.link);
		let attrs = `lang="en-GB" sz="${sz}"`;
		if (r.bold || p.bold) attrs += ' b="1"';
		if (r.italic || r.math || p.italic) attrs += ' i="1"';
		if (ext) attrs += ' u="sng"';
		if (r.strike) attrs += ' strike="sngStrike"';
		if (r.sup) attrs += ' baseline="30000"';
		attrs += ' dirty="0"';
		const c = r.colour || (r.code ? this.col.code : r.link || r.internal ? this.col.link : colour);
		let inner = `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`;
		if (r.highlight) inner += `<a:highlight><a:srgbClr val="${this.col.highlight}"/></a:highlight>`;
		inner += `<a:latin typeface="${escXml(r.code || r.math ? this.fonts.code : font)}"/>`;
		if (ext) inner += `<a:hlinkClick r:id="${this.addRel('hyperlink', r.link, true)}"/>`;
		return `<a:r><a:rPr ${attrs}>${inner}</a:rPr><a:t>${escXml(r.text)}</a:t></a:r>`;
	}

	paraXml(p, base) {
		const size = Math.max(8, base * (p.rel || 1));
		const k = p.kind;
		let marL = 0;
		let indent = 0;
		if (p.bullet) { marL = 285750 + 342900 * p.level; indent = -285750; } else if (k === 'quote') marL = 228600 + 342900 * (p.level || 0);
		else if (p.level) marL = 342900 * p.level;
		let pPr = `<a:pPr marL="${marL}" indent="${indent}"${p.align ? ` algn="${p.align}"` : ''}>`;
		pPr += `<a:lnSpc><a:spcPct val="${k === 'code' ? 100000 : 110000}"/></a:lnSpc><a:spcBef><a:spcPts val="${k === 'code' ? 0 : ri(size * 35)}"/></a:spcBef>`;
		if (p.bullet) {
			if (this.design.listMarkerAccent) pPr += `<a:buClr><a:srgbClr val="${this.col.accent}"/></a:buClr>`;
			if (p.bullet === 'bullet') pPr += `<a:buFont typeface="Arial"/><a:buChar char="${['•', '–', '▪'][p.level % 3]}"/>`;
			else pPr += `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod" startAt="${p.bullet.start}"/>`;
		} else {
			pPr += '<a:buNone/>';
		}
		pPr += '</a:pPr>';
		const colour = { text: this.col.text, heading: this.col.heading, quote: this.col.quote, code: this.col.code, title: this.col.title, subtitle: this.col.heading }[k] || this.col.text;
		const font = k === 'code' ? this.fonts.code : k === 'title' ? this.fonts.title : k === 'heading' || k === 'subtitle' ? this.fonts.heading : this.fonts.body;
		const runs = p.runs.map((r) => this.runXml(r, size, colour, font, p)).join('');
		return `<a:p>${pPr}${runs}<a:endParaRPr lang="en-GB" sz="${ri(size * 100)}" dirty="0"/></a:p>`;
	}

	/** Flatten text blocks into slide paragraphs. */
	flatten(blocks, level = 0, quote = false) {
		const out = [];
		for (const b of blocks) {
			switch (b.type) {
				case 'heading':
					out.push({ runs: parseInline(b.text), kind: 'heading', rel: b.level <= 2 ? 1.25 : 1.1, bold: true, level });
					break;
				case 'paragraph':
					out.push({ runs: parseInline(b.text), kind: quote ? 'quote' : 'text', level, italic: quote && this.design.blockquoteStyle !== 'card' && this.design.blockquoteStyle !== 'bar' });
					break;
				case 'list':
					b.items.forEach((item) => {
						const runs = parseInline(item.text);
						if (item.task !== null && item.task !== undefined) runs.unshift({ text: item.task ? '☑ ' : '☐ ', colour: this.design.listMarkerAccent ? this.col.accent : undefined });
						out.push({ runs, kind: quote ? 'quote' : 'text', level, bullet: item.task !== null && item.task !== undefined ? 'bullet' : b.ordered ? { start: b.start } : 'bullet' });
						for (const child of item.children) out.push(...this.flatten([child], level + 1, quote));
					});
					break;
				case 'blockquote':
					out.push(...this.flatten(b.children, level, true));
					break;
				case 'callout':
					out.push({ runs: [{ text: b.title, bold: true }], kind: 'quote', level });
					out.push(...this.flatten(b.children, level, true));
					break;
				case 'code':
					for (const ln of b.text.split('\n')) out.push({ runs: [{ text: ln || ' ', code: true }], kind: 'code', rel: 0.8, level });
					break;
				case 'math':
					out.push({ runs: [{ text: b.text, math: true }], kind: 'code', rel: 0.9, level, align: 'ctr' });
					break;
				default:
					break;
			}
		}
		return out;
	}

	estimateHeight(paras, w, size) {
		let total = 0;
		for (const p of paras) {
			const sz = size * (p.rel || 1);
			const indent = p.bullet ? 285750 + 342900 * p.level : 342900 * (p.level || 0);
			const avail = Math.max(w - indent, w * 0.3);
			const perLine = Math.max(8, Math.floor(avail / (sz * (p.kind === 'code' ? 0.6 : 0.52) * EMU_PT)));
			const lines = runsText(p.runs).split('\n').reduce((n, seg) => n + Math.max(1, Math.ceil(seg.length / perLine)), 0);
			total += lines * sz * (p.kind === 'code' ? 1.0 : 1.1) * 1.2 * EMU_PT + (p.kind === 'code' ? 0 : sz * 0.35 * EMU_PT);
		}
		return total;
	}

	fitSize(paras, w, h, base, min) {
		for (let size = base; size > min; size -= 1) if (this.estimateHeight(paras, w, size) <= h) return size;
		return min;
	}

	textBox(name, x, y, w, h, parasXml, anchor = 't') {
		const id = this.nextId();
		return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escXml(name)} ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
			+ `<p:spPr><a:xfrm><a:off x="${ri(x)}" y="${ri(y)}"/><a:ext cx="${ri(w)}" cy="${ri(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
			+ `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}"><a:normAutofit/></a:bodyPr><a:lstStyle/>${parasXml}</p:txBody></p:sp>`;
	}

	rect(x, y, w, h, fill) {
		const id = this.nextId();
		return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Accent ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
			+ `<p:spPr><a:xfrm><a:off x="${ri(x)}" y="${ri(y)}"/><a:ext cx="${ri(w)}" cy="${ri(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
			+ `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
	}

	tableXml(tbl, x, y, w, hMax) {
		const cols = Math.max(tbl.header.length, ...tbl.rows.map((r) => r.length));
		const rows = [tbl.header, ...tbl.rows];
		let size = 18;
		while (size > 9 && size * 2.1 * EMU_PT * rows.length > hMax) size--;
		const rowH = ri(size * 2.1 * EMU_PT);
		const colW = Math.floor(w / cols);
		const ts = this.design.tableStyle;
		const ln = (tag) => `<a:${tag} w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="${this.col.tableBorder}"/></a:solidFill><a:prstDash val="solid"/></a:${tag}>`;
		const trs = rows.map((cells, rIdx) => {
			const header = rIdx === 0;
			let fill = '<a:noFill/>';
			if (header && ts !== 'minimal') fill = `<a:solidFill><a:srgbClr val="${this.col.tableHeaderBg}"/></a:solidFill>`;
			else if (!header && ts !== 'minimal' && rIdx % 2 === 0) fill = `<a:solidFill><a:srgbClr val="${this.col.tableStripe}"/></a:solidFill>`;
			const tcs = [];
			for (let c = 0; c < cols; c++) {
				const algn = tbl.align[c] === 'center' ? 'ctr' : tbl.align[c] === 'right' ? 'r' : 'l';
				const colour = header && ts !== 'minimal' ? this.col.tableHeaderText : this.col.text;
				const p = { kind: 'text', bold: header };
				const runs = parseInline(cells[c] || '').map((r) => this.runXml(r, size, colour, this.fonts.body, p)).join('');
				tcs.push(`<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="${algn}"/>${runs}<a:endParaRPr lang="en-GB" sz="${size * 100}" dirty="0"/></a:p></a:txBody>`
					+ `<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="ctr">${ln('lnL')}${ln('lnR')}${ln('lnT')}${ln('lnB')}${fill}</a:tcPr></a:tc>`);
			}
			return `<a:tr h="${rowH}">${tcs.join('')}</a:tr>`;
		}).join('');
		const id = this.nextId();
		return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
			+ `<p:xfrm><a:off x="${ri(x)}" y="${ri(y)}"/><a:ext cx="${colW * cols}" cy="${rowH * rows.length}"/></p:xfrm>`
			+ `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/>`
			+ `<a:tblGrid>${Array.from({ length: cols }, () => `<a:gridCol w="${colW}"/>`).join('')}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
	}

	async placeImages(shapes, imgs, region) {
		const loaded = [];
		for (const b of imgs) {
			const img = await this.plugin.loadImage(b.src, this.file.path);
			if (img) loaded.push({ b, img });
		}
		if (!loaded.length) return;
		const n = loaded.length;
		const cols = region.w / region.h > 1.4 ? Math.min(n, 3) : n > 2 ? 2 : 1;
		const rows = Math.ceil(n / cols);
		const gap = 152400;
		const cw = (region.w - gap * (cols - 1)) / cols;
		const chh = (region.h - gap * (rows - 1)) / rows;
		loaded.forEach(({ b, img }, i) => {
			const r = Math.floor(i / cols);
			const c = i % cols;
			const ratio = img.width / img.height || 1;
			let w = Math.min(cw, img.width * 9525 * 2);
			let h = w / ratio;
			if (h > chh) { h = chh; w = h * ratio; }
			const x = region.x + c * (cw + gap) + (cw - w) / 2;
			const y = region.y + r * (chh + gap) + (chh - h) / 2;
			const name = `image${this.media.length + 1}.${img.ext}`;
			this.media.push({ name, bytes: img.bytes });
			const rid = this.addRel('image', `../media/${name}`);
			const id = this.nextId();
			shapes.push(`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}" descr="${escXml(b.alt)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
				+ `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
				+ `<p:spPr><a:xfrm><a:off x="${ri(x)}" y="${ri(y)}"/><a:ext cx="${ri(w)}" cy="${ri(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
		});
	}

	async slideXml(blocks, idx, total) {
		this.cur = { rels: [{ id: 'rId1', type: REL + 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }], seq: 2, id: 2 };
		const W = this.W;
		const H = this.H;
		const o = this.opts;
		const col = this.col;
		const deco = o.decoration;
		const shapes = [];
		blocks = blocks.filter((b) => b.type !== 'hr');
		let title = null;
		if (blocks.length && blocks[0].type === 'heading') title = blocks.shift();
		const images = blocks.filter((b) => b.type === 'image');
		const tables = blocks.filter((b) => b.type === 'table');
		const text = blocks.filter((b) => b.type !== 'image' && b.type !== 'table');
		const mx = ri(W * 0.06);
		const footerH = o.slideNumbers || o.footer ? ri(H * 0.07) : ri(H * 0.035);
		const titleBold = this.th.headings[1].bold;

		if (deco === 'topbar') shapes.push(this.rect(0, 0, W, ri(H * 0.02), col.accent));

		const simple = text.length <= 3 && text.every((b) => b.type === 'heading' || (b.type === 'paragraph' && b.text.length <= 200));
		const isTitle = o.titleSlides && title && !images.length && !tables.length && simple && (idx === 0 || !text.length || (title.level === 1 && text.length === 1 && text[0].type === 'paragraph'));

		if (isTitle) {
			const tPara = { runs: parseInline(title.text), kind: 'title', align: 'ctr', bold: titleBold };
			const tSize = this.fitSize([tPara], W - 2 * mx, H * 0.3, Math.min(60, ri(o.titleSize * 1.4)), 24);
			shapes.push(this.textBox('Title', mx, H * 0.2, W - 2 * mx, H * 0.32, this.paraXml(tPara, tSize), 'b'));
			if (deco !== 'none') shapes.push(this.rect((W - W * 0.1) / 2, H * 0.555, W * 0.1, H * 0.009, col.accent));
			if (text.length) {
				const paras = text.map((b) => ({ runs: parseInline(b.text), kind: 'subtitle', align: 'ctr' }));
				shapes.push(this.textBox('Subtitle', mx, H * 0.6, W - 2 * mx, H * 0.25, paras.map((p) => this.paraXml(p, ri(o.bodySize * 1.15))).join('')));
			}
		} else {
			let top = ri(H * 0.06);
			if (title) {
				const tH = ri(H * 0.14);
				const tx = deco === 'leftbar' ? mx + ri(W * 0.02) : mx;
				if (deco === 'leftbar') shapes.push(this.rect(mx, top + tH * 0.15, ri(W * 0.006), tH * 0.85, col.accent));
				const tPara = { runs: parseInline(title.text), kind: 'title', bold: titleBold };
				const tSize = this.fitSize([tPara], W - mx - tx, tH, o.titleSize, 18);
				shapes.push(this.textBox('Title', tx, top, W - mx - tx, tH, this.paraXml(tPara, tSize), 'b'));
				top += tH;
				if (deco === 'underline') shapes.push(this.rect(mx, top + H * 0.015, W * 0.08, H * 0.008, col.accent));
				top += ri(H * 0.055);
			}
			const bottom = H - footerH - ri(H * 0.02);
			const region = { x: mx, y: top, w: W - 2 * mx, h: bottom - top };
			let textRegion = region;
			let imgRegion = null;
			if (images.length && (text.length || tables.length)) {
				const gap = W * 0.03;
				const tw2 = (region.w - gap) * 0.56;
				textRegion = Object.assign({}, region, { w: tw2 });
				imgRegion = { x: region.x + tw2 + gap, y: region.y, w: region.w - tw2 - gap, h: region.h };
			} else if (images.length) {
				imgRegion = region;
				textRegion = null;
			}
			if (textRegion) {
				const paras = this.flatten(text);
				let tableTop = textRegion.y;
				let tableSpace = textRegion.h;
				if (paras.length) {
					let textH = textRegion.h;
					if (tables.length) {
						const want = this.estimateHeight(paras, textRegion.w, o.bodySize);
						textH = clamp(want, textRegion.h * 0.25, textRegion.h * 0.55);
					}
					const size = this.fitSize(paras, textRegion.w, textH, o.bodySize, 11);
					shapes.push(this.textBox('Content', textRegion.x, textRegion.y, textRegion.w, textH, paras.map((p) => this.paraXml(p, size)).join('')));
					tableTop = textRegion.y + textH + H * 0.03;
					tableSpace = textRegion.y + textRegion.h - tableTop;
				}
				if (tables.length) {
					const each = (tableSpace - H * 0.02 * (tables.length - 1)) / tables.length;
					tables.forEach((t, k) => shapes.push(this.tableXml(t, textRegion.x, tableTop + k * (each + H * 0.02), textRegion.w, each)));
				}
			}
			if (imgRegion) await this.placeImages(shapes, images, imgRegion);
		}

		const fy = H - footerH + ri(footerH * 0.15);
		const fh = ri(footerH * 0.6);
		if (o.footer) {
			const fp = { runs: [{ text: expandTokens(o.footer, this.ctx, { page: idx + 1, pages: total }), colour: col.muted }], kind: 'text' };
			shapes.push(this.textBox('Footer', mx, fy, W * 0.65, fh, this.paraXml(fp, 11), 'ctr'));
		}
		if (o.slideNumbers) {
			const id = this.nextId();
			shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Slide Number ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
				+ `<p:spPr><a:xfrm><a:off x="${ri(W - mx - W * 0.15)}" y="${ri(fy)}"/><a:ext cx="${ri(W * 0.15)}" cy="${fh}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
				+ `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="r"/>`
				+ `<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-GB" sz="1100" dirty="0"><a:solidFill><a:srgbClr val="${col.muted}"/></a:solidFill><a:latin typeface="${escXml(this.fonts.body)}"/></a:rPr><a:t>${idx + 1}</a:t></a:fld>`
				+ '<a:endParaRPr lang="en-GB" sz="1100" dirty="0"/></a:p></p:txBody></p:sp>');
		}
		const xml = XML_HEAD + `<p:sld ${P_NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${col.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`
			+ `<p:spTree>${SPTREE_HEAD}${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
		return { xml, rels: this.cur.rels };
	}

	themeXml() {
		const c = this.col;
		const clr = (tag, v) => `<a:${tag}><a:srgbClr val="${v}"/></a:${tag}>`;
		const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
		return XML_HEAD + `<a:theme xmlns:a="${NS_A}" name="PDF Customiser"><a:themeElements><a:clrScheme name="PDF Customiser">`
			+ clr('dk1', c.text) + clr('lt1', c.bg) + clr('dk2', c.title) + clr('lt2', 'EEEEEE')
			+ clr('accent1', c.accent) + clr('accent2', c.heading) + clr('accent3', c.link) + clr('accent4', c.tableHeaderBg) + clr('accent5', c.quote) + clr('accent6', c.muted)
			+ clr('hlink', c.link) + clr('folHlink', c.link) + '</a:clrScheme>'
			+ `<a:fontScheme name="PDF Customiser"><a:majorFont><a:latin typeface="${escXml(this.fonts.title)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>`
			+ `<a:minorFont><a:latin typeface="${escXml(this.fonts.body)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`
			+ `<a:fmtScheme name="PDF Customiser"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst>`
			+ `<a:lnStyleLst>${[6350, 12700, 19050].map((w) => `<a:ln w="${w}">${solid}</a:ln>`).join('')}</a:lnStyleLst>`
			+ '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
			+ `<a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
	}

	async build(slideTexts) {
		for (let i = 0; i < slideTexts.length; i++) this.slides.push(await this.slideXml(parseMarkdown(slideTexts[i]), i, slideTexts.length));
		if (!this.slides.length) this.slides.push(await this.slideXml([{ type: 'heading', level: 1, text: this.ctx.tokens.title }], 0, 1));

		const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';
		const n = this.slides.length;
		const contentTypes = XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			+ '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
			+ '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>'
			+ `<Override PartName="/ppt/presentation.xml" ContentType="${PML}presentation.main+xml"/>`
			+ `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${PML}slideMaster+xml"/>`
			+ `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${PML}slideLayout+xml"/>`
			+ Array.from({ length: n }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${PML}slide+xml"/>`).join('')
			+ '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
			+ `<Override PartName="/ppt/presProps.xml" ContentType="${PML}presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${PML}viewProps+xml"/>`
			+ `<Override PartName="/ppt/tableStyles.xml" ContentType="${PML}tableStyles+xml"/>`
			+ '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
			+ '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';

		const presentation = XML_HEAD + `<p:presentation ${P_NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>`
			+ `<p:sldIdLst>${Array.from({ length: n }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + i}"/>`).join('')}</p:sldIdLst>`
			+ `<p:sldSz cx="${this.W}" cy="${this.H}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
		const presRels = [
			{ id: 'rId1', type: REL + 'slideMaster', target: 'slideMasters/slideMaster1.xml' },
			{ id: 'rId2', type: REL + 'theme', target: 'theme/theme1.xml' },
			{ id: 'rId3', type: REL + 'presProps', target: 'presProps.xml' },
			{ id: 'rId4', type: REL + 'viewProps', target: 'viewProps.xml' },
			{ id: 'rId5', type: REL + 'tableStyles', target: 'tableStyles.xml' },
			...Array.from({ length: n }, (_, i) => ({ id: `rId${10 + i}`, type: REL + 'slide', target: `slides/slide${i + 1}.xml` })),
		];
		const bg = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${this.col.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
		const lvl = (sz) => `<a:lvl1pPr><a:defRPr sz="${sz}"/></a:lvl1pPr>`;
		const master = XML_HEAD + `<p:sldMaster ${P_NS}><p:cSld>${bg}<p:spTree>${SPTREE_HEAD}</p:spTree></p:cSld>`
			+ '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
			+ '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
			+ `<p:txStyles><p:titleStyle>${lvl(4400)}</p:titleStyle><p:bodyStyle>${lvl(2000)}</p:bodyStyle><p:otherStyle>${lvl(1800)}</p:otherStyle></p:txStyles></p:sldMaster>`;
		const layout = XML_HEAD + `<p:sldLayout ${P_NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${SPTREE_HEAD}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

		const files = [
			{ name: '[Content_Types].xml', data: contentTypes },
			{ name: '_rels/.rels', data: relsXml([
				{ id: 'rId1', type: REL + 'officeDocument', target: 'ppt/presentation.xml' },
				{ id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' },
				{ id: 'rId3', type: REL + 'extended-properties', target: 'docProps/app.xml' },
			]) },
			{ name: 'docProps/core.xml', data: corePropsXml(this.ctx.tokens.title, this.ctx.tokens.author) },
			{ name: 'docProps/app.xml', data: APP_PROPS },
			{ name: 'ppt/presentation.xml', data: presentation },
			{ name: 'ppt/_rels/presentation.xml.rels', data: relsXml(presRels) },
			{ name: 'ppt/presProps.xml', data: XML_HEAD + `<p:presentationPr ${P_NS}/>` },
			{ name: 'ppt/viewProps.xml', data: XML_HEAD + `<p:viewPr ${P_NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>` },
			{ name: 'ppt/tableStyles.xml', data: XML_HEAD + `<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>` },
			{ name: 'ppt/theme/theme1.xml', data: this.themeXml() },
			{ name: 'ppt/slideMasters/slideMaster1.xml', data: master },
			{ name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: relsXml([
				{ id: 'rId1', type: REL + 'slideLayout', target: '../slideLayouts/slideLayout1.xml' },
				{ id: 'rId2', type: REL + 'theme', target: '../theme/theme1.xml' },
			]) },
			{ name: 'ppt/slideLayouts/slideLayout1.xml', data: layout },
			{ name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: relsXml([{ id: 'rId1', type: REL + 'slideMaster', target: '../slideMasters/slideMaster1.xml' }]) },
			...this.slides.flatMap((sl, i) => [
				{ name: `ppt/slides/slide${i + 1}.xml`, data: sl.xml },
				{ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: relsXml(sl.rels) },
			]),
			...this.media.map((m) => ({ name: `ppt/media/${m.name}`, data: m.bytes, store: true })),
		];
		return createZip(files);
	}
}

/* ────────────────────────────────────────────────────────────────────────── *
 * Settings UI
 * ────────────────────────────────────────────────────────────────────────── */

const SAMPLE_HTML = `
<h1>Quarterly Field Notes</h1>
<p>This is a <strong>live preview</strong> of your export. It shows <em>italic text</em>, <a class="external-link" href="https://obsidian.md">a link</a>, <code>inline code</code> and <mark>highlighted words</mark> so you can judge the overall feel before exporting.</p>
<h2>A second-level heading</h2>
<p>Body text flows here with your chosen font, size, line height and paragraph spacing. Good typography makes long documents comfortable to read.</p>
<ul><li>First bullet point</li><li>Second bullet point<ul><li>A nested item</li></ul></li></ul>
<blockquote><p>Design is not just what it looks like and feels like. Design is how it works.</p></blockquote>
<h3>Third-level heading</h3>
<table><thead><tr><th>Task</th><th>Status</th><th>Owner</th></tr></thead><tbody><tr><td>Research</td><td>Done</td><td>Alex</td></tr><tr><td>Draft</td><td>In progress</td><td>Sam</td></tr><tr><td>Review</td><td>Planned</td><td>Jo</td></tr></tbody></table>
<pre><code>const style = "beautiful";
exportNote(style);</code></pre>
<h4>Fourth-level heading</h4>
<p>Smaller headings keep their own fonts and colours.</p>
<h5>Fifth-level heading</h5>
<h6>Sixth-level heading</h6>
<ol><li>Numbered step one</li><li>Numbered step two</li></ol>
<hr>
<p>End of sample.</p>`;

let PREVIEW_SEQ = 0;
const MM_PX = 3.7795;

class LivePreview {
	constructor(plugin, parent, large = false) {
		this.plugin = plugin;
		this.large = large;
		this.id = 'pdfc-preview-' + ++PREVIEW_SEQ;
		const wrap = parent.createDiv({ cls: 'pdfc-preview' + (large ? ' is-large' : '') });
		if (!large) {
			const bar = wrap.createDiv({ cls: 'pdfc-preview-bar' });
			bar.createSpan({ cls: 'pdfc-preview-label', text: 'Live preview' });
			const btn = bar.createEl('button', { cls: 'clickable-icon', attr: { 'aria-label': 'Open larger preview' } });
			setIcon(btn, 'maximize-2');
			btn.addEventListener('click', () => new PreviewModal(plugin.app, plugin).open());
		}
		this.viewport = wrap.createDiv({ cls: 'pdfc-preview-viewport' });
		this.page = this.viewport.createDiv({ cls: 'pdfc-preview-page', attr: { id: this.id } });
		this.styleEl = this.page.createEl('style');
		this.frame = this.page.createDiv({ cls: 'pdfc-pv-frame' });
		this.header = this.page.createDiv({ cls: 'pdfc-pv-band is-header' });
		this.footer = this.page.createDiv({ cls: 'pdfc-pv-band is-footer' });
		this.content = this.page.createDiv({ cls: 'pdfc-preview-content markdown-rendered' });
		const sanitize = obsidian.sanitizeHTMLToDom;
		if (typeof sanitize === 'function') this.content.appendChild(sanitize(SAMPLE_HTML));
		else this.content.innerHTML = SAMPLE_HTML;
		this.watermark = this.page.createDiv({ cls: 'pdfc-pv-watermark' });
		this.ro = new ResizeObserver(() => this.rescale());
		this.ro.observe(this.viewport);
		this.update();
	}

	rescale() {
		const w = this.viewport.clientWidth;
		if (!w) return;
		const scale = w / 794;
		this.page.style.transform = `scale(${scale})`;
		this.viewport.style.height = `${1123 * scale}px`;
	}

	update() {
		const s = this.plugin.settings;
		const P = '#' + this.id;
		const m = s.page.controlMargins ? s.page.margins : { top: 15, right: 15, bottom: 15, left: 15 };
		this.page.style.padding = `${m.top * MM_PX}px ${m.right * MM_PX}px ${m.bottom * MM_PX}px ${m.left * MM_PX}px`;
		for (const el of [this.frame, this.watermark]) {
			el.style.top = `${m.top * MM_PX}px`;
			el.style.bottom = `${m.bottom * MM_PX}px`;
			el.style.left = `${m.left * MM_PX}px`;
			el.style.right = `${m.right * MM_PX}px`;
		}
		let css = buildContentCss(s, `${P} .pdfc-preview-content`) + '\n' + buildPreviewPageCss(s, P);
		if (s.customCss && s.customCss.trim()) css += `\n${P} .pdfc-preview-content { ${s.customCss} }`;
		this.styleEl.textContent = css;

		const slots = computeSlots(s).normal;
		const hf = s.headerFooter;
		const pageVals = { page: 2, pages: 5 };
		const fill = (band, prefix, height, isTop) => {
			band.empty();
			band.style.height = `${height * MM_PX}px`;
			band.style.left = `${m.left * MM_PX}px`;
			band.style.right = `${m.right * MM_PX}px`;
			band.style.fontFamily = cssFont(hf.font || s.typography.body.font, 'sans-serif');
			band.style.fontSize = `${hf.size}pt`;
			band.style.color = s.colours.headerFooter;
			const cells = ['left', 'center', 'right'].map((pos) => expandTokens(slots[`${prefix}-${pos}`], SAMPLE_CTX, pageVals));
			const any = cells.some(Boolean);
			const inner = band.createDiv({ cls: 'pdfc-pv-band-inner' });
			if (any && hf.enabled && hf.separator) inner.style[isTop ? 'borderBottom' : 'borderTop'] = `0.5pt solid ${s.colours.headerFooter}`;
			cells.forEach((txt, i) => inner.createDiv({ cls: 'pdfc-pv-cell is-' + ['left', 'center', 'right'][i], text: txt }));
		};
		fill(this.header, 'top', m.top, true);
		fill(this.footer, 'bottom', m.bottom, false);
		this.rescale();
	}

	destroy() {
		this.ro.disconnect();
	}
}

class PreviewModal extends Modal {
	constructor(app, plugin) {
		super(app);
		this.plugin = plugin;
	}

	onOpen() {
		this.modalEl.addClass('pdfc-preview-modal');
		this.titleEl.setText('PDF preview');
		this.contentEl.createDiv({ cls: 'pdfc-muted', text: 'Sample content rendered with your current settings (page 2 of 5 shown for header and footer).' });
		this.preview = new LivePreview(this.plugin, this.contentEl, true);
	}

	onClose() {
		if (this.preview) this.preview.destroy();
		this.contentEl.empty();
	}
}

/** Mini "Canva-style" thumbnail of a design. */
function drawThumb(el, data) {
	const c = data.colours;
	const t = data.typography;
	const d = data.design;
	const thumb = el.createDiv({ cls: 'pdfc-thumb' });
	thumb.style.backgroundColor = c.background;
	thumb.style.cssText += backgroundPatternCss({ colours: c, design: d }).replace(/ !important/g, '');
	if (d.accentBar === 'top') thumb.style.boxShadow = `inset 0 4px 0 ${c.accent}`;
	if (d.accentBar === 'left') thumb.style.boxShadow = `inset 4px 0 0 ${c.accent}`;
	if (d.pageBorder !== 'none') {
		const fr = thumb.createDiv({ cls: 'pdfc-thumb-frame' });
		fr.style.border = d.pageBorder === 'double' ? `3px double ${c.accent}` : `${d.pageBorder === 'thick' ? 2 : 1}px solid ${c.accent}`;
	}
	const heading = (level, text) => {
		const h = t.headings['h' + level];
		const e = thumb.createDiv({ cls: `pdfc-thumb-h${level}`, text });
		e.style.cssText += `font-family: ${cssFont(h.font || t.body.font)}; color: ${c['h' + level]}; font-weight: ${h.weight}; font-style: ${h.italic ? 'italic' : 'normal'}; text-transform: ${h.uppercase ? 'uppercase' : 'none'}; letter-spacing: ${h.letterSpacing}em;`;
		if (h.align && h.align !== 'inherit') e.style.textAlign = h.align;
		if (level <= d.decorationLevels) e.style.cssText += headingDecoration(d.headingStyle, level, c, d).replace(/ !important/g, '');
	};
	heading(1, 'Heading');
	heading(2, 'Subheading');
	[100, 94, 98, 62].forEach((w) => {
		const ln = thumb.createDiv({ cls: 'pdfc-thumb-line' });
		ln.style.width = w + '%';
		ln.style.background = c.text;
	});
	const q = thumb.createDiv({ cls: 'pdfc-thumb-quote' });
	q.style.borderLeft = `2px solid ${c.blockquoteBorder}`;
	q.style.background = c.blockquoteBg;
	q.createDiv({ cls: 'pdfc-thumb-line' }).style.background = c.blockquoteText;
	const tb = thumb.createDiv({ cls: 'pdfc-thumb-table' });
	tb.style.borderColor = c.tableBorder;
	tb.createDiv({ cls: 'pdfc-thumb-th' }).style.background = d.tableStyle === 'minimal' ? c.text : c.tableHeaderBg;
	for (let i = 0; i < 2; i++) {
		const row = tb.createDiv({ cls: 'pdfc-thumb-tr' });
		row.style.borderColor = c.tableBorder;
		if (i === 1 && (d.tableStyle === 'striped' || d.tableStyle === 'header')) row.style.background = c.tableStripe;
	}
	return thumb;
}

const TABS = [
	{ id: 'general', name: 'General', icon: 'settings' },
	{ id: 'designs', name: 'Designs', icon: 'layout-template', preview: true },
	{ id: 'colours', name: 'Colours', icon: 'palette', preview: true },
	{ id: 'typography', name: 'Typography', icon: 'type', preview: true },
	{ id: 'layout', name: 'Page & layout', icon: 'file-text', preview: true },
	{ id: 'headerFooter', name: 'Header & footer', icon: 'panel-top', preview: true },
	{ id: 'presets', name: 'Presets', icon: 'bookmark' },
	{ id: 'docx', name: 'Word', icon: 'file-type' },
	{ id: 'pptx', name: 'PowerPoint', icon: 'presentation' },
	{ id: 'advanced', name: 'Advanced', icon: 'code' },
];

const WEIGHTS = { 300: 'Light (300)', 400: 'Regular (400)', 500: 'Medium (500)', 600: 'Semibold (600)', 700: 'Bold (700)', 800: 'Extra bold (800)', 900: 'Black (900)' };

class PdfCustomiserSettingTab extends PluginSettingTab {
	constructor(app, plugin) {
		super(app, plugin);
		this.plugin = plugin;
		this.activeTab = 'general';
		this.galleryFilter = 'all';
		this.previews = [];
	}

	get s() {
		return this.plugin.settings;
	}

	display() {
		const el = this.containerEl;
		this.destroyPreviews();
		el.empty();
		el.addClass('pdfc-settings');

		const list = el.createEl('datalist', { attr: { id: 'pdfc-font-list' } });
		for (const f of FONT_SUGGESTIONS) list.createEl('option', { attr: { value: f } });

		const hero = el.createDiv({ cls: 'pdfc-hero' });
		hero.createDiv({ cls: 'pdfc-hero-title', text: 'PDF Customiser' });
		hero.createDiv({ cls: 'pdfc-hero-sub', text: 'Design your PDF, Word and PowerPoint exports. Changes apply to Obsidian’s own “Export to PDF” dialog.' });

		const nav = el.createDiv({ cls: 'pdfc-tabs', attr: { role: 'tablist' } });
		for (const tab of TABS) {
			const btn = nav.createEl('button', { cls: 'pdfc-tab' + (tab.id === this.activeTab ? ' is-active' : ''), attr: { role: 'tab' } });
			setIcon(btn.createSpan({ cls: 'pdfc-tab-icon' }), tab.icon);
			btn.createSpan({ text: tab.name });
			btn.addEventListener('click', () => { this.activeTab = tab.id; this.display(); });
		}

		const body = el.createDiv({ cls: 'pdfc-tab-body' });
		const tab = TABS.find((t) => t.id === this.activeTab) || TABS[0];
		let target = body;
		if (tab.preview) {
			const split = body.createDiv({ cls: 'pdfc-split' });
			target = split.createDiv({ cls: 'pdfc-split-main' });
			const side = split.createDiv({ cls: 'pdfc-split-side' });
			this.previews.push(new LivePreview(this.plugin, side));
		}
		this['render_' + tab.id](target);
	}

	hide() {
		this.destroyPreviews();
	}

	destroyPreviews() {
		this.previews.forEach((p) => p.destroy());
		this.previews = [];
	}

	/** Called after every edit. */
	changed(styleEdit = true) {
		this.plugin.onSettingsChanged(styleEdit);
		this.previews.forEach((p) => p.update());
	}

	/* Setting helpers ------------------------------------------------------ */

	heading(el, name, desc) {
		const st = new Setting(el).setName(name).setHeading();
		if (desc) st.setDesc(desc);
		return st;
	}

	note(el, text) {
		el.createDiv({ cls: 'pdfc-note', text });
	}

	toggle(el, name, path, desc, opts = {}) {
		const st = new Setting(el).setName(name);
		if (desc) st.setDesc(desc);
		st.addToggle((t) => t.setValue(!!getPath(this.s, path)).onChange((v) => {
			setPath(this.s, path, v);
			this.changed(opts.style !== false);
			if (opts.rerender) this.display();
		}));
		return st;
	}

	text(el, name, path, desc, opts = {}) {
		const st = new Setting(el).setName(name);
		if (desc) st.setDesc(desc);
		st.addText((t) => {
			t.setPlaceholder(opts.placeholder || '').setValue(String(getPath(this.s, path) ?? ''));
			if (opts.font) {
				t.inputEl.setAttr('list', 'pdfc-font-list');
				t.inputEl.addClass('pdfc-font-input');
				t.inputEl.style.fontFamily = cssFont(t.getValue(), 'inherit');
			}
			t.onChange((v) => {
				setPath(this.s, path, v);
				if (opts.font) t.inputEl.style.fontFamily = cssFont(v, 'inherit');
				if (opts.onChange) opts.onChange(v);
				this.changed(opts.style !== false);
			});
		});
		return st;
	}

	dropdown(el, name, path, options, desc, opts = {}) {
		const st = new Setting(el).setName(name);
		if (desc) st.setDesc(desc);
		st.addDropdown((d) => {
			d.addOptions(options).setValue(String(getPath(this.s, path))).onChange((v) => {
				setPath(this.s, path, opts.number ? Number(v) : v);
				this.changed(opts.style !== false);
				if (opts.rerender) this.display();
			});
		});
		return st;
	}

	slider(el, name, path, min, max, step, desc, unit = '') {
		const st = new Setting(el).setName(name);
		if (desc) st.setDesc(desc);
		const label = st.controlEl.createSpan({ cls: 'pdfc-slider-value' });
		const fmt = (v) => `${+Number(v).toFixed(3)}${unit}`;
		label.setText(fmt(getPath(this.s, path)));
		st.addSlider((sl) => sl.setLimits(min, max, step).setValue(Number(getPath(this.s, path))).setDynamicTooltip().onChange((v) => {
			setPath(this.s, path, v);
			label.setText(fmt(v));
			this.changed();
		}));
		return st;
	}

	colour(el, name, path, desc) {
		const st = new Setting(el).setName(name);
		if (desc) st.setDesc(desc);
		let picker;
		st.addColorPicker((cp) => {
			picker = cp;
			cp.setValue(normHex(getPath(this.s, path))).onChange((v) => { setPath(this.s, path, v); this.changed(); });
		});
		st.addExtraButton((b) => b.setIcon('rotate-ccw').setTooltip('Reset to the design’s colour').onClick(() => {
			const tpl = getTemplate(this.s.activeTemplate) || TEMPLATES[0];
			const v = getPath(tpl.data, path) ?? getPath(DEFAULT_SETTINGS, path);
			if (!v) return;
			setPath(this.s, path, v);
			picker.setValue(v);
			this.changed();
		}));
		return st;
	}

	/* Tabs ---------------------------------------------------------------- */

	render_general(el) {
		this.heading(el, 'PDF export', 'PDF Customiser styles Obsidian’s built-in “Export to PDF”. A “PDF Customiser” section is also added to that dialog so you can switch designs quickly.');
		this.toggle(el, 'Apply custom style to PDF exports', 'enabled', 'Turn off to export PDFs with your normal Obsidian theme.', { style: false });

		this.heading(el, 'Note menu', 'Choose which export buttons appear in the ⋯ menu at the top of a note.');
		this.toggle(el, 'Export to PDF', 'menu.pdf', null, { style: false });
		this.toggle(el, 'Export to Word (.docx)', 'menu.docx', null, { style: false });
		this.toggle(el, 'Export to PowerPoint (.pptx)', 'menu.pptx', null, { style: false });
		this.toggle(el, 'Also show in the file explorer’s right-click menu', 'menu.fileExplorer', null, { style: false });

		this.heading(el, 'Saving Word & PowerPoint files');
		this.dropdown(el, 'Save location', 'output.location', { same: 'Same folder as the note', folder: 'A specific vault folder', ask: 'Ask every time (desktop)' }, null, { rerender: true, style: false });
		if (this.s.output.location === 'folder') this.text(el, 'Export folder', 'output.folder', 'Vault-relative path. Created if it doesn’t exist.', { placeholder: 'Exports', style: false });
		this.text(el, 'File name', 'output.filename', 'Tokens: {title}, {filename}, {date}, {time}, {folder}, {fm.key}.', { placeholder: '{title}', style: false });
		this.toggle(el, 'Overwrite existing files', 'output.overwrite', 'When off, a number is added (e.g. “Note 2.docx”).', { style: false });
		this.toggle(el, 'Open the file after exporting', 'output.openAfter', 'Opens it in your default app (desktop).', { style: false });

		this.heading(el, 'Document details');
		this.text(el, 'Default author', 'info.author', 'Used for {author} when the note has no “author” in its frontmatter.', { placeholder: 'Your name' });
		this.text(el, 'Date format', 'info.dateFormat', 'Moment.js format used for {date}, e.g. D MMMM YYYY or YYYY-MM-DD.', { placeholder: 'D MMMM YYYY' });
		this.toggle(el, 'Include embedded notes', 'content.resolveEmbeds', 'Word & PowerPoint: replace ![[Note]] embeds with the embedded note’s content.', { style: false });
	}

	render_designs(el) {
		this.heading(el, 'Design gallery', 'Pick a starting design, then fine-tune anything in the other tabs. Your header, footer and page settings are kept.');
		const presets = this.s.presets;
		const chips = el.createDiv({ cls: 'pdfc-chips' });
		for (const [id, label] of TEMPLATE_CATEGORIES) {
			if (id === 'mine' && !presets.length) continue;
			const chip = chips.createEl('button', { cls: 'pdfc-chip' + (this.galleryFilter === id ? ' is-active' : ''), text: label });
			chip.addEventListener('click', () => { this.galleryFilter = id; this.display(); });
		}
		const grid = el.createDiv({ cls: 'pdfc-gallery' });
		const f = this.galleryFilter;
		const card = (name, desc, data, active, onPick, badge) => {
			const c = grid.createDiv({ cls: 'pdfc-card' + (active ? ' is-active' : ''), attr: { tabindex: '0', role: 'button', 'aria-label': name } });
			drawThumb(c, data);
			const meta = c.createDiv({ cls: 'pdfc-card-meta' });
			const row = meta.createDiv({ cls: 'pdfc-card-name', text: name });
			if (badge) row.createSpan({ cls: 'pdfc-badge', text: badge });
			meta.createDiv({ cls: 'pdfc-card-desc', text: desc });
			c.addEventListener('click', onPick);
			c.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onPick(); } });
		};
		if (f !== 'mine') {
			for (const t of TEMPLATES) {
				if (f !== 'all' && t.category !== f) continue;
				const active = t.id === this.s.activeTemplate;
				card(t.name, t.description, t.data, active, () => {
					this.plugin.applyTemplate(t.id);
					new Notice(`Design applied: ${t.name}`);
					this.display();
				}, active && this.s.templateModified ? 'Edited' : active ? 'Active' : '');
			}
		}
		if (f === 'all' || f === 'mine') {
			presets.forEach((p, i) => {
				const data = deepMerge(clone(TEMPLATES[0].data), p.data);
				card(`★ ${p.name}`, 'Your preset', data, false, () => {
					this.plugin.applyPreset(i);
					new Notice(`Preset applied: ${p.name}`);
					this.display();
				});
			});
		}
		if (this.plugin.undoSnapshot) {
			new Setting(el).setName('Changed your mind?').setDesc('Restore the style you had before the last design was applied.')
				.addButton((b) => b.setButtonText('Undo').onClick(() => { this.plugin.undoStyle(); this.display(); }));
		}

		this.heading(el, 'Fine-tune the design');
		this.dropdown(el, 'Heading decoration', 'design.headingStyle', {
			none: 'None', underline: 'Underline', 'double-underline': 'Double underline', overline: 'Overline', leftbar: 'Side bar',
			band: 'Tinted band', boxed: 'Outlined box', pill: 'Solid pill', centered: 'Centred with rule',
		}, 'Uses the accent colour.');
		this.slider(el, 'Decorate headings up to', 'design.decorationLevels', 1, 6, 1, 'Apply the decoration to H1 down to this level.', '');
		this.dropdown(el, 'Block quotes', 'design.blockquoteStyle', { bar: 'Side bar', card: 'Card', quote: 'Large quotation mark', minimal: 'Minimal italic' });
		this.dropdown(el, 'Tables', 'design.tableStyle', { grid: 'Full grid', striped: 'Striped rows', minimal: 'Minimal (academic)', header: 'Bold header + grid' });
		this.dropdown(el, 'Images', 'design.imageStyle', { plain: 'Plain', rounded: 'Rounded corners', shadow: 'Soft shadow', framed: 'Photo frame' });
		this.dropdown(el, 'Page accent bar', 'design.accentBar', { none: 'None', top: 'Top edge', left: 'Left edge' }, 'A strip of accent colour on every page.');
		this.dropdown(el, 'Page border', 'design.pageBorder', { none: 'None', thin: 'Thin', thick: 'Thick', double: 'Double' });
		this.dropdown(el, 'Background pattern', 'design.backgroundPattern', { none: 'None', dots: 'Dot grid', grid: 'Squared grid', lines: 'Ruled lines' });
		this.toggle(el, 'Accent-coloured list markers', 'design.listMarkerAccent');
		this.toggle(el, 'Drop cap after the first H1', 'design.dropCap', 'A large decorative first letter, magazine style.');
		this.toggle(el, 'Underline links', 'design.linkUnderline');
		this.slider(el, 'Corner roundness', 'design.radius', 0, 20, 1, 'For code blocks, quotes, bands and images.', 'px');
	}

	render_colours(el) {
		this.heading(el, 'Palette tools');
		this.colour(el, 'Accent', 'colours.accent', 'Used for decorations, bars, borders and markers.');
		new Setting(el).setName('Generate a palette from the accent').setDesc('Creates matching heading, link, quote and table colours from your accent.')
			.addButton((b) => b.setButtonText('Generate').onClick(() => { this.plugin.paletteFromAccent(); this.changed(); this.display(); }));
		new Setting(el).setName('Use the H1 colour for every heading')
			.addButton((b) => b.setButtonText('Apply').onClick(() => {
				for (let n = 2; n <= 6; n++) this.s.colours['h' + n] = this.s.colours.h1;
				this.changed();
				this.display();
			}));

		this.heading(el, 'Headings', 'Each heading level has its own colour.');
		for (let n = 1; n <= 6; n++) this.colour(el, `Heading ${n}`, `colours.h${n}`);

		this.heading(el, 'Body & page');
		this.colour(el, 'Body text', 'colours.text');
		this.colour(el, 'Page background', 'colours.background');
		this.colour(el, 'Links', 'colours.link');
		this.colour(el, 'Highlight (==text==)', 'colours.highlight');
		this.colour(el, 'Horizontal rules', 'colours.hr');

		this.heading(el, 'Blocks');
		this.colour(el, 'Quote text', 'colours.blockquoteText');
		this.colour(el, 'Quote border', 'colours.blockquoteBorder');
		this.colour(el, 'Quote background', 'colours.blockquoteBg');
		this.colour(el, 'Code text', 'colours.codeText');
		this.colour(el, 'Code background', 'colours.codeBg');
		this.colour(el, 'Table header background', 'colours.tableHeaderBg');
		this.colour(el, 'Table header text', 'colours.tableHeaderText');
		this.colour(el, 'Table borders', 'colours.tableBorder');
		this.colour(el, 'Table stripes', 'colours.tableStripe');

		this.heading(el, 'Header & footer');
		this.colour(el, 'Header, footer & page numbers', 'colours.headerFooter');
	}

	render_typography(el) {
		this.note(el, 'Fonts must be installed on your computer. Type any font name or pick a suggestion; you can also enter a full CSS stack like “Lora, Georgia, serif”.');
		this.heading(el, 'Body text');
		this.text(el, 'Font', 'typography.body.font', null, { font: true, placeholder: 'Inter' });
		this.slider(el, 'Size', 'typography.body.size', 8, 16, 0.5, null, 'pt');
		this.slider(el, 'Line height', 'typography.body.lineHeight', 1, 2.4, 0.05, null, '×');
		this.dropdown(el, 'Weight', 'typography.body.weight', WEIGHTS, null, { number: true });
		this.slider(el, 'Paragraph spacing', 'design.paragraphSpacing', 0, 2, 0.05, 'Space after each paragraph.', 'em');
		this.dropdown(el, 'Alignment', 'design.textAlign', { left: 'Left', justify: 'Justified' });

		this.heading(el, 'Headings', 'Every heading level has its own font, size, weight and colour. Leave a font blank to use the body font.');
		const quick = new Setting(el).setName('Quick: one font for all headings');
		let quickFont = '';
		quick.addText((t) => {
			t.inputEl.setAttr('list', 'pdfc-font-list');
			t.setPlaceholder('e.g. Playfair Display').onChange((v) => { quickFont = v; });
		});
		quick.addButton((b) => b.setButtonText('Apply').onClick(() => {
			for (let n = 1; n <= 6; n++) this.s.typography.headings['h' + n].font = quickFont;
			this.changed();
			this.display();
		}));
		new Setting(el).setName('Quick: size scale').setDesc('Sets heading sizes as a typographic scale of the body size.')
			.addDropdown((d) => d.addOptions({ '': 'Choose a scale…', 1.125: 'Major second (1.125)', 1.2: 'Minor third (1.2)', 1.25: 'Major third (1.25)', 1.333: 'Perfect fourth (1.333)', 1.414: 'Augmented fourth (1.414)', 1.5: 'Perfect fifth (1.5)' })
				.onChange((v) => {
					if (!v) return;
					const r = Number(v);
					const b = this.s.typography.body.size;
					const exps = [4, 3, 2, 1, 0.5, 0];
					for (let n = 1; n <= 6; n++) this.s.typography.headings['h' + n].size = Math.round(b * Math.pow(r, exps[n - 1]) * 2) / 2;
					this.changed();
					this.display();
				}));

		for (let n = 1; n <= 6; n++) {
			const h = this.s.typography.headings['h' + n];
			const det = el.createEl('details', { cls: 'pdfc-collapsible' });
			if (n === 1) det.open = true;
			const sum = det.createEl('summary');
			const label = sum.createSpan({ cls: 'pdfc-summary-label', text: `Heading ${n}` });
			label.style.fontFamily = cssFont(h.font || this.s.typography.body.font);
			label.style.color = this.s.colours['h' + n];
			label.style.fontWeight = String(h.weight);
			sum.createSpan({ cls: 'pdfc-summary-meta', text: `${primaryFont(h.font) || 'Body font'} · ${h.size}pt` });
			const inner = det.createDiv({ cls: 'pdfc-collapsible-body' });
			const p = `typography.headings.h${n}`;
			this.text(inner, 'Font', `${p}.font`, null, { font: true, placeholder: 'Same as body' });
			this.slider(inner, 'Size', `${p}.size`, 8, 48, 0.5, null, 'pt');
			this.dropdown(inner, 'Weight', `${p}.weight`, WEIGHTS, null, { number: true });
			this.colour(inner, 'Colour', `colours.h${n}`);
			this.toggle(inner, 'Italic', `${p}.italic`);
			this.toggle(inner, 'Uppercase', `${p}.uppercase`);
			this.slider(inner, 'Letter spacing', `${p}.letterSpacing`, -0.05, 0.3, 0.01, null, 'em');
			this.dropdown(inner, 'Alignment', `${p}.align`, { inherit: 'Same as body', left: 'Left', center: 'Centre', right: 'Right' });
		}

		this.heading(el, 'Code');
		this.text(el, 'Font', 'typography.code.font', null, { font: true, placeholder: 'JetBrains Mono' });
		this.slider(el, 'Size', 'typography.code.size', 7, 14, 0.5, null, 'pt');
	}

	render_layout(el) {
		this.heading(el, 'Margins');
		this.toggle(el, 'Let PDF Customiser set the margins', 'page.controlMargins', 'Needed for header/footer space. In Obsidian’s export dialog keep “Margin” on Default so these apply.', { rerender: true });
		if (this.s.page.controlMargins) {
			for (const side of ['top', 'bottom', 'left', 'right']) this.slider(el, side.charAt(0).toUpperCase() + side.slice(1), `page.margins.${side}`, 0, 40, 1, null, 'mm');
		}

		this.heading(el, 'Flow');
		this.toggle(el, 'Start each H1 on a new page', 'page.h1PageBreak');
		this.toggle(el, 'Keep blocks together', 'page.avoidBreaks', 'Avoid splitting code blocks, tables, quotes and images across pages.');
		this.slider(el, 'Maximum image width', 'page.imageMaxWidth', 20, 100, 5, null, '%');
		this.toggle(el, 'Print link addresses', 'page.showLinkUrls', 'Shows the URL in brackets after external links — handy for printed copies.');

		this.heading(el, 'Watermark');
		this.toggle(el, 'Show a watermark', 'watermark.enabled', null, { rerender: true });
		if (this.s.watermark.enabled) {
			this.text(el, 'Text', 'watermark.text', null, { placeholder: 'DRAFT' });
			this.colour(el, 'Colour', 'watermark.colour');
			this.slider(el, 'Opacity', 'watermark.opacity', 0.02, 0.5, 0.01, null, '');
			this.slider(el, 'Angle', 'watermark.angle', -90, 90, 1, null, '°');
			this.slider(el, 'Size', 'watermark.size', 24, 200, 2, null, 'pt');
		}
	}

	render_headerFooter(el) {
		this.heading(el, 'Header & footer', 'Three slots each — left, centre and right.');
		this.toggle(el, 'Show header & footer', 'headerFooter.enabled', null, { rerender: true });
		if (this.s.headerFooter.enabled) {
			for (const [part, label] of [['header', 'Header'], ['footer', 'Footer']]) {
				for (const pos of ['left', 'center', 'right']) {
					this.text(el, `${label} — ${pos === 'center' ? 'centre' : pos}`, `headerFooter.${part}.${pos}`, null, { placeholder: 'Empty' });
				}
			}
			this.text(el, 'Font', 'headerFooter.font', 'Leave blank to use the body font.', { font: true, placeholder: 'Same as body' });
			this.slider(el, 'Size', 'headerFooter.size', 6, 14, 0.5, null, 'pt');
			this.colour(el, 'Colour', 'colours.headerFooter');
			this.toggle(el, 'Separator line', 'headerFooter.separator');
			this.toggle(el, 'Show on the first page', 'headerFooter.showOnFirstPage', 'Turn off for a clean cover page.');
		}

		this.heading(el, 'Page numbers');
		this.toggle(el, 'Show page numbers', 'pageNumbers.enabled', null, { rerender: true });
		if (this.s.pageNumbers.enabled) {
			this.dropdown(el, 'Position', 'pageNumbers.position', {
				'top-left': 'Top left', 'top-center': 'Top centre', 'top-right': 'Top right',
				'bottom-left': 'Bottom left', 'bottom-center': 'Bottom centre', 'bottom-right': 'Bottom right',
			}, 'Takes over that header/footer slot.');
			this.text(el, 'Format', 'pageNumbers.format', 'Use {page} and {pages}, e.g. “Page {page} of {pages}”.', { placeholder: '{page} / {pages}' });
			this.dropdown(el, 'Numbering style', 'pageNumbers.style', {
				decimal: '1, 2, 3', 'decimal-leading-zero': '01, 02, 03', 'lower-roman': 'i, ii, iii', 'upper-roman': 'I, II, III', 'lower-alpha': 'a, b, c', 'upper-alpha': 'A, B, C',
			});
			this.toggle(el, 'Number the first page', 'pageNumbers.showOnFirstPage');
		}

		this.heading(el, 'Tokens', 'Click a token to copy it. Tokens work in header, footer, file names and PowerPoint footers.');
		const box = el.createDiv({ cls: 'pdfc-tokens' });
		for (const [tok, desc] of TOKENS) {
			const t = box.createEl('button', { cls: 'pdfc-token', text: tok, attr: { 'aria-label': desc } });
			t.addEventListener('click', () => { navigator.clipboard.writeText(tok); new Notice(`Copied ${tok}`); });
		}
	}

	render_presets(el) {
		this.heading(el, 'Save the current style', 'A preset stores colours, fonts, design, page, header & footer, watermark and custom CSS.');
		let name = '';
		new Setting(el).setName('Preset name')
			.addText((t) => t.setPlaceholder('e.g. Client report').onChange((v) => { name = v; }))
			.addButton((b) => b.setButtonText('Save preset').setCta().onClick(() => {
				if (!name.trim()) { new Notice('Give your preset a name first.'); return; }
				this.plugin.savePreset(name.trim());
				new Notice(`Saved preset “${name.trim()}”`);
				this.display();
			}));

		this.heading(el, 'Your presets');
		if (!this.s.presets.length) this.note(el, 'No presets yet. Presets also appear in the Designs gallery and in the PDF export dialog.');
		this.s.presets.forEach((p, i) => {
			new Setting(el).setName(p.name).setDesc(`Saved ${new Date(p.created).toLocaleString()}`)
				.addButton((b) => b.setButtonText('Apply').setCta().onClick(() => { this.plugin.applyPreset(i); new Notice(`Applied “${p.name}”`); this.display(); }))
				.addExtraButton((b) => b.setIcon('save').setTooltip('Overwrite with current style').onClick(() => {
					p.data = this.plugin.snapshot(PRESET_KEYS);
					p.created = Date.now();
					this.changed(false);
					new Notice(`Updated “${p.name}”`);
					this.display();
				}))
				.addExtraButton((b) => b.setIcon('copy').setTooltip('Copy as JSON').onClick(() => {
					navigator.clipboard.writeText(JSON.stringify({ name: p.name, data: p.data }, null, 2));
					new Notice('Preset copied to clipboard');
				}))
				.addExtraButton((b) => b.setIcon('trash-2').setTooltip('Delete').onClick(() => {
					if (!window.confirm(`Delete preset “${p.name}”?`)) return;
					this.s.presets.splice(i, 1);
					this.changed(false);
					this.display();
				}));
		});

		this.heading(el, 'Import a preset', 'Paste JSON copied from another vault.');
		let json = '';
		const imp = new Setting(el).addTextArea((t) => { t.setPlaceholder('{ "name": "…", "data": { … } }').onChange((v) => { json = v; }); t.inputEl.rows = 4; t.inputEl.addClass('pdfc-code-input'); });
		imp.addButton((b) => b.setButtonText('Import').onClick(() => {
			try {
				const obj = JSON.parse(json);
				if (!obj || !isObj(obj.data)) throw new Error('Missing "data"');
				this.s.presets.push({ name: String(obj.name || 'Imported preset'), created: Date.now(), data: obj.data });
				this.changed(false);
				new Notice('Preset imported');
				this.display();
			} catch (e) {
				new Notice('That doesn’t look like a PDF Customiser preset: ' + e.message);
			}
		}));

		this.heading(el, 'Reset');
		new Setting(el).setName('Reset all settings').setDesc('Restores every option to its default. Your presets are kept.')
			.addButton((b) => b.setButtonText('Reset').setWarning().onClick(() => {
				if (!window.confirm('Reset all PDF Customiser settings to their defaults?')) return;
				this.plugin.resetSettings();
				this.display();
			}));
	}

	render_docx(el) {
		this.heading(el, 'Word export', 'Notes are converted to native Word styles (Heading 1–6, lists, tables, images, links), so the document stays editable.');
		this.toggle(el, 'Use the PDF style', 'docx.useTheme', 'Carry over your fonts, colours and heading decorations. Off = plain Office look.', { style: false });
		this.dropdown(el, 'Page size', 'docx.pageSize', { A4: 'A4', Letter: 'US Letter', Legal: 'US Legal', A5: 'A5', A3: 'A3' }, 'Margins follow Page & layout.', { style: false });
		this.toggle(el, 'Add the note title at the top', 'docx.includeTitle', null, { style: false });
		this.toggle(el, 'Add a table of contents', 'docx.toc', 'Word builds it when the document is opened (you may be asked to update fields).', { style: false });
		this.toggle(el, 'Include header & footer', 'docx.headerFooter', 'Uses the Header & footer tab.', { style: false });
		this.toggle(el, 'Include page numbers', 'docx.pageNumbers', null, { style: false });
		this.toggle(el, 'Embed images', 'docx.embedImages', 'SVG and WebP images are converted to PNG.', { style: false });
		this.toggle(el, 'Apply the page background colour', 'docx.pageBackground', 'Word only prints page colour when “Print background colours” is on.', { style: false });
	}

	render_pptx(el) {
		this.heading(el, 'PowerPoint export', 'Each slide is separated by a horizontal rule (---) on its own line. The first heading on a slide becomes its title.');
		this.dropdown(el, 'Slide size', 'pptx.aspect', { '16:9': 'Widescreen 16:9', '4:3': 'Standard 4:3' }, null, { style: false });
		this.dropdown(el, 'Also start a new slide at', 'pptx.splitOn', { none: 'Only at ---', h1: 'Every H1', h2: 'Every H1 and H2' }, null, { style: false });
		this.toggle(el, 'Title slides', 'pptx.titleSlides', 'The first slide, and any slide with only a heading, gets a centred title layout.', { style: false });
		this.slider(el, 'Title size', 'pptx.titleSize', 20, 60, 1, null, 'pt');
		this.slider(el, 'Body text size', 'pptx.bodySize', 12, 32, 1, 'Shrinks automatically when a slide is full.', 'pt');
		this.dropdown(el, 'Title decoration', 'pptx.decoration', { none: 'None', underline: 'Accent underline', leftbar: 'Accent side bar', topbar: 'Accent bar across the top' }, null, { style: false });
		this.toggle(el, 'Slide numbers', 'pptx.slideNumbers', null, { style: false });
		this.text(el, 'Footer text', 'pptx.footer', 'Shown on every slide. Tokens allowed, e.g. {title} · {date}.', { placeholder: 'Empty', style: false });

		this.heading(el, 'Colours & fonts');
		this.toggle(el, 'Use the PDF style', 'pptx.useTheme', 'Use your PDF fonts and colours (H1 for titles, accent for decorations).', { rerender: true, style: false });
		if (!this.s.pptx.useTheme) {
			this.colour(el, 'Slide background', 'pptx.background');
			this.colour(el, 'Titles', 'pptx.titleColour');
			this.colour(el, 'Text', 'pptx.textColour');
			this.colour(el, 'Accent', 'pptx.accent');
		}
	}

	render_advanced(el) {
		this.heading(el, 'Custom CSS', 'Extra CSS for the PDF, scoped to the exported note. Write plain selectors, e.g. “h2 { font-variant: small-caps; }”.');
		const ta = new Setting(el).setClass('pdfc-css-setting');
		ta.addTextArea((t) => {
			t.setPlaceholder('h2 {\n  font-variant: small-caps;\n}').setValue(this.s.customCss).onChange((v) => { this.s.customCss = v; this.changed(); });
			t.inputEl.rows = 12;
			t.inputEl.addClass('pdfc-code-input');
		});
		this.heading(el, 'Troubleshooting');
		new Setting(el).setName('Copy the generated print CSS').setDesc('Useful for debugging or turning your design into a CSS snippet.')
			.addButton((b) => b.setButtonText('Copy').onClick(() => {
				navigator.clipboard.writeText(buildPrintCss(this.s, this.plugin.buildContext(this.app.workspace.getActiveFile())));
				new Notice('Print CSS copied');
			}));
	}
}

/* ────────────────────────────────────────────────────────────────────────── *
 * Quick design switcher
 * ────────────────────────────────────────────────────────────────────────── */

class DesignSuggestModal extends SuggestModal {
	constructor(app, plugin) {
		super(app);
		this.plugin = plugin;
		this.setPlaceholder('Choose a design or preset for your exports…');
	}

	getSuggestions(query) {
		const q = query.toLowerCase();
		const items = [
			...TEMPLATES.map((t) => ({ kind: 'template', id: t.id, name: t.name, desc: t.description })),
			...this.plugin.settings.presets.map((p, i) => ({ kind: 'preset', idx: i, name: p.name, desc: 'Your preset' })),
		];
		return items.filter((it) => it.name.toLowerCase().includes(q));
	}

	renderSuggestion(it, el) {
		el.createDiv({ text: (it.kind === 'preset' ? '★ ' : '') + it.name });
		el.createEl('small', { cls: 'pdfc-muted', text: it.desc });
	}

	onChooseSuggestion(it) {
		if (it.kind === 'template') this.plugin.applyTemplate(it.id);
		else this.plugin.applyPreset(it.idx);
		new Notice(`Export design: ${it.name}`);
	}
}

/* ────────────────────────────────────────────────────────────────────────── *
 * Plugin
 * ────────────────────────────────────────────────────────────────────────── */

class PdfCustomiserPlugin extends Plugin {
	async onload() {
		await this.loadSettings();
		this.saveSettingsDebounced = debounce(() => this.saveSettings(), 400, true);
		this.undoSnapshot = null;

		this.styleEl = document.head.createEl('style', { attr: { id: 'pdf-customiser-print-styles' } });
		this.refreshPrintStyles();

		this.settingTab = new PdfCustomiserSettingTab(this.app, this);
		this.addSettingTab(this.settingTab);

		this.registerEvent(this.app.workspace.on('file-menu', (menu, file, source, leaf) => {
			if (!(file instanceof TFile) || file.extension !== 'md') return;
			const noteMenu = /more-options/.test(source || '');
			if (!noteMenu && !this.settings.menu.fileExplorer) return;
			const m = this.settings.menu;
			const add = (title, icon, fn) => menu.addItem((item) => item.setTitle(title).setIcon(icon).setSection('pdf-customiser').onClick(fn));
			if (m.pdf) add('Export to PDF', 'file-output', () => this.exportPdf(file, noteMenu ? leaf : null));
			if (m.docx) add('Export to Word (.docx)', 'file-text', () => this.exportDocx(file));
			if (m.pptx) add('Export to PowerPoint (.pptx)', 'presentation', () => this.exportPptx(file));
		}));

		this.registerEvent(this.app.workspace.on('file-open', () => this.refreshPrintStyles()));
		this.registerEvent(this.app.metadataCache.on('changed', (f) => {
			const active = this.app.workspace.getActiveFile();
			if (active && f.path === active.path) this.refreshPrintStyles();
		}));

		const withActive = (fn) => (checking) => {
			const f = this.app.workspace.getActiveFile();
			if (!f || f.extension !== 'md') return false;
			if (!checking) fn(f);
			return true;
		};
		this.addCommand({ id: 'export-pdf', name: 'Export current note to PDF', checkCallback: withActive((f) => this.exportPdf(f)) });
		this.addCommand({ id: 'export-docx', name: 'Export current note to Word (.docx)', checkCallback: withActive((f) => this.exportDocx(f)) });
		this.addCommand({ id: 'export-pptx', name: 'Export current note to PowerPoint (.pptx)', checkCallback: withActive((f) => this.exportPptx(f)) });
		this.addCommand({ id: 'choose-design', name: 'Choose export design…', callback: () => new DesignSuggestModal(this.app, this).open() });
		this.addCommand({ id: 'toggle-style', name: 'Toggle custom PDF style on/off', callback: () => {
			this.settings.enabled = !this.settings.enabled;
			this.onSettingsChanged(false);
			new Notice(`PDF Customiser style ${this.settings.enabled ? 'on' : 'off'}`);
		} });

		// Add our section to Obsidian's own Export to PDF dialog.
		this.modalObserver = new MutationObserver((muts) => {
			for (const mu of muts) {
				for (const node of mu.addedNodes) {
					if (!(node instanceof HTMLElement) || !node.classList.contains('modal-container')) continue;
					[20, 150, 500].forEach((ms) => window.setTimeout(() => this.injectIntoExportModal(node), ms));
				}
			}
		});
		this.modalObserver.observe(document.body, { childList: true });
	}

	onunload() {
		if (this.modalObserver) this.modalObserver.disconnect();
		if (this.styleEl) this.styleEl.remove();
		document.querySelectorAll('.pdfc-native').forEach((e) => e.remove());
	}

	/* Settings ------------------------------------------------------------ */

	async loadSettings() {
		this.settings = deepMerge(clone(DEFAULT_SETTINGS), (await this.loadData()) || {});
		if (!Array.isArray(this.settings.presets)) this.settings.presets = [];
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	onSettingsChanged(styleEdit = true) {
		if (styleEdit) this.settings.templateModified = true;
		this.saveSettingsDebounced();
		this.refreshPrintStyles();
	}

	snapshot(keys) {
		const out = {};
		for (const k of keys) out[k] = clone(this.settings[k]);
		return out;
	}

	applyStyleData(data, keys) {
		this.undoSnapshot = this.snapshot(PRESET_KEYS.concat(['activeTemplate', 'templateModified']));
		for (const k of keys) {
			if (data[k] === undefined) continue;
			this.settings[k] = isObj(DEFAULT_SETTINGS[k]) ? deepMerge(clone(DEFAULT_SETTINGS[k]), data[k]) : data[k];
		}
	}

	applyTemplate(id) {
		const t = getTemplate(id);
		if (!t) return;
		this.applyStyleData(t.data, TEMPLATE_KEYS);
		this.settings.activeTemplate = id;
		this.settings.templateModified = false;
		this.onSettingsChanged(false);
	}

	applyPreset(idx) {
		const p = this.settings.presets[idx];
		if (!p) return;
		this.applyStyleData(p.data, PRESET_KEYS);
		this.settings.templateModified = true;
		this.onSettingsChanged(false);
	}

	undoStyle() {
		if (!this.undoSnapshot) return;
		Object.assign(this.settings, clone(this.undoSnapshot));
		this.undoSnapshot = null;
		this.onSettingsChanged(false);
	}

	savePreset(name) {
		const existing = this.settings.presets.find((p) => p.name === name);
		const data = this.snapshot(PRESET_KEYS);
		if (existing) { existing.data = data; existing.created = Date.now(); } else this.settings.presets.push({ name, created: Date.now(), data });
		this.onSettingsChanged(false);
	}

	resetSettings() {
		const presets = this.settings.presets;
		this.settings = clone(DEFAULT_SETTINGS);
		this.settings.presets = presets;
		this.onSettingsChanged(false);
	}

	paletteFromAccent() {
		const c = this.settings.colours;
		const a = c.accent;
		const dark = luminance(c.background) < 0.3;
		const ink = dark ? '#ffffff' : '#111111';
		c.h1 = a;
		c.h2 = mix(a, ink, 0.2);
		c.h3 = mix(a, ink, 0.35);
		c.h4 = mix(a, ink, 0.5);
		c.h5 = mix(a, ink, 0.6);
		c.h6 = mix(a, ink, 0.65);
		c.link = mix(a, ink, 0.1);
		c.blockquoteBorder = a;
		c.blockquoteBg = mix(a, c.background, 0.93);
		c.tableHeaderBg = mix(a, c.background, 0.82);
		c.tableHeaderText = mix(a, ink, 0.55);
		c.tableStripe = mix(a, c.background, 0.95);
		c.tableBorder = mix(a, c.background, 0.7);
		c.hr = mix(a, c.background, 0.7);
		c.codeBg = mix(a, c.background, 0.93);
		c.codeText = mix(a, ink, 0.3);
		c.highlight = mix(a, c.background, 0.75);
	}

	/* Context & styles ------------------------------------------------------- */

	buildContext(file) {
		const fm = file ? (this.app.metadataCache.getFileCache(file) || {}).frontmatter || {} : {};
		const moment = window.moment;
		const fmt = this.settings.info.dateFormat || 'YYYY-MM-DD';
		const now = new Date();
		const parent = file && file.parent && file.parent.path !== '/' ? file.parent.name : '';
		return {
			frontmatter: fm,
			tokens: {
				title: String(fm.title || (file ? file.basename : 'Untitled')),
				filename: file ? file.basename : 'Untitled',
				folder: parent,
				date: moment ? moment().format(fmt) : now.toLocaleDateString(),
				time: moment ? moment().format('HH:mm') : now.toLocaleTimeString(),
				author: String(fm.author || this.settings.info.author || ''),
				vault: this.app.vault.getName(),
			},
		};
	}

	refreshPrintStyles(file) {
		if (!this.styleEl) return;
		let f = file || this.app.workspace.getActiveFile();
		if (f && f.extension !== 'md') f = null;
		this.styleEl.textContent = this.settings.enabled ? buildPrintCss(this.settings, this.buildContext(f)) : '';
	}

	/* Native export dialog ---------------------------------------------------- */

	injectIntoExportModal(container) {
		if (!container.isConnected || container.querySelector('.pdfc-native')) return;
		const selects = Array.from(container.querySelectorAll('select'));
		const isPdf = selects.some((sel) => Array.from(sel.options).some((o) => /^(a3|a4|a5|letter|legal|tabloid)$/i.test(o.value) || /^(a4|letter)$/i.test(o.text.trim())));
		if (!isPdf) return;
		const content = container.querySelector('.modal-content') || container.querySelector('.modal');
		if (!content) return;
		this.refreshPrintStyles();
		const section = document.createElement('div');
		section.className = 'pdfc-native';
		const cta = container.querySelector('button.mod-cta');
		const anchor = cta ? cta.closest('.setting-item') || cta.closest('.modal-button-container') : null;
		if (anchor && content.contains(anchor)) anchor.parentElement.insertBefore(section, anchor);
		else content.appendChild(section);
		this.renderNativeSection(section);
	}

	renderNativeSection(el) {
		el.empty();
		const s = this.settings;
		const rerender = () => this.renderNativeSection(el);
		new Setting(el).setName('PDF Customiser').setHeading();
		new Setting(el).setName('Apply custom style').addToggle((t) => t.setValue(s.enabled).onChange((v) => { s.enabled = v; this.onSettingsChanged(false); rerender(); }));
		if (s.enabled) {
			const tpl = getTemplate(s.activeTemplate);
			new Setting(el).setName('Design').setDesc(tpl ? `Current: ${tpl.name}${s.templateModified ? ' (edited)' : ''}` : 'Current: custom')
				.addDropdown((d) => {
					d.addOption('', 'Keep current style');
					for (const t of TEMPLATES) d.addOption('t:' + t.id, t.name);
					s.presets.forEach((p, i) => d.addOption('p:' + i, '★ ' + p.name));
					d.setValue('').onChange((v) => {
						if (v.startsWith('t:')) this.applyTemplate(v.slice(2));
						else if (v.startsWith('p:')) this.applyPreset(Number(v.slice(2)));
						this.refreshPrintStyles();
						rerender();
					});
				});
			new Setting(el).setName('Header & footer').addToggle((t) => t.setValue(s.headerFooter.enabled).onChange((v) => { s.headerFooter.enabled = v; this.onSettingsChanged(false); }));
			new Setting(el).setName('Page numbers').addToggle((t) => t.setValue(s.pageNumbers.enabled).onChange((v) => { s.pageNumbers.enabled = v; this.onSettingsChanged(false); }));
			new Setting(el).setName('Watermark').setDesc(s.watermark.text ? `“${s.watermark.text}”` : '')
				.addToggle((t) => t.setValue(s.watermark.enabled).onChange((v) => { s.watermark.enabled = v; this.onSettingsChanged(false); }));
			const tip = el.createDiv({ cls: 'pdfc-native-tip' });
			tip.setText(s.page.controlMargins ? 'Tip: keep Margin on “Default” so PDF Customiser’s margins and header/footer space apply.' : 'Margins are controlled by this dialog.');
		}
		new Setting(el).addButton((b) => b.setButtonText('More style options…').onClick(() => this.openSettings()));
	}

	openSettings() {
		const setting = this.app.setting;
		if (!setting) return;
		setting.open();
		setting.openTabById(this.manifest.id);
	}

	/* Exports ----------------------------------------------------------------- */

	async exportPdf(file, leaf) {
		try {
			if (leaf) this.app.workspace.setActiveLeaf(leaf, { focus: true });
			const active = this.app.workspace.getActiveFile();
			if (!active || active.path !== file.path) await this.app.workspace.getLeaf(false).openFile(file);
			this.refreshPrintStyles(file);
			const commands = this.app.commands;
			const ok = commands && commands.executeCommandById('workspace:export-pdf');
			if (!ok) new Notice('Obsidian’s PDF export isn’t available here.');
		} catch (e) {
			console.error('PDF Customiser:', e);
			new Notice('Could not start the PDF export: ' + e.message);
		}
	}

	async exportDocx(file) {
		await this.runExport(file, 'docx', 'Word', async (md, ctx) => {
			const blocks = parseMarkdown(md);
			return new DocxBuilder(this, file, ctx).build(blocks);
		});
	}

	async exportPptx(file) {
		await this.runExport(file, 'pptx', 'PowerPoint', async (md, ctx) => {
			const slides = splitSlides(md, this.settings.pptx.splitOn);
			return new PptxBuilder(this, file, ctx).build(slides);
		});
	}

	async runExport(file, ext, label, build) {
		const progress = new Notice(`Exporting “${file.basename}” to ${label}…`, 0);
		try {
			const ctx = this.buildContext(file);
			const md = await this.preprocess(await this.app.vault.cachedRead(file), file);
			const bytes = await build(md, ctx);
			const saved = await this.saveOutput(file, ext, bytes, ctx);
			progress.hide();
			if (!saved) return;
			new Notice(`Exported to ${saved.display}`);
			if (this.settings.output.openAfter) this.openExternally(saved);
		} catch (e) {
			progress.hide();
			console.error('PDF Customiser:', e);
			new Notice(`${label} export failed: ${e.message}`);
		}
	}

	async preprocess(md, file) {
		let text = stripFrontmatter(md).replace(/%%[\s\S]*?%%/g, '').replace(/<!--[\s\S]*?-->/g, '');
		if (this.settings.content.resolveEmbeds) text = await this.resolveEmbeds(text, file, 0, new Set([file.path]));
		return text.replace(/[ \t]\^[A-Za-z0-9-]+[ \t]*$/gm, '');
	}

	async resolveEmbeds(text, file, depth, seen) {
		if (depth > 3) return text;
		const re = /!\[\[([^\]]+)\]\]/g;
		const found = [];
		let m;
		while ((m = re.exec(text))) {
			const target = m[1].split('|')[0];
			const hash = target.indexOf('#');
			const path = hash >= 0 ? target.slice(0, hash) : target;
			const sub = hash >= 0 ? target.slice(hash + 1) : '';
			if (IMAGE_EXT.test(path)) continue;
			const dest = path ? this.app.metadataCache.getFirstLinkpathDest(path, file.path) : file;
			if (!dest || dest.extension !== 'md' || (path && seen.has(dest.path))) continue;
			found.push({ full: m[0], dest, sub });
		}
		for (const f of found) {
			let content = stripFrontmatter(await this.app.vault.cachedRead(f.dest));
			if (f.sub) content = extractSection(content, f.sub);
			const next = new Set(seen);
			next.add(f.dest.path);
			content = await this.resolveEmbeds(content, f.dest, depth + 1, next);
			text = text.split(f.full).join('\n' + content.trim() + '\n');
		}
		return text;
	}

	async loadImage(src, sourcePath) {
		try {
			let bytes;
			if (/^https?:\/\//i.test(src)) {
				const res = await obsidian.requestUrl({ url: src });
				bytes = new Uint8Array(res.arrayBuffer);
			} else if (/^data:image\//i.test(src)) {
				const b64 = src.split(',')[1] || '';
				bytes = Uint8Array.from(atob(b64), (ch) => ch.charCodeAt(0));
			} else {
				let p = src;
				try { p = decodeURIComponent(src); } catch (e) { /* keep raw */ }
				let f = this.app.metadataCache.getFirstLinkpathDest(p, sourcePath);
				if (!f) {
					const abs = this.app.vault.getAbstractFileByPath(normalizePath(p));
					if (abs instanceof TFile) f = abs;
				}
				if (!f) return null;
				bytes = new Uint8Array(await this.app.vault.readBinary(f));
			}
			const info = imageInfo(bytes);
			if (info) return Object.assign({ bytes }, info);
			return await rasterise(bytes);
		} catch (e) {
			console.warn('PDF Customiser: could not load image', src, e);
			return null;
		}
	}

	async saveOutput(file, ext, bytes, ctx) {
		const o = this.settings.output;
		const base = (expandTokens(o.filename || '{title}', ctx, null).replace(/[\\/:*?"<>|#^[\]]/g, '-').trim() || file.basename);
		const name = `${base}.${ext}`;
		const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

		if (o.location === 'ask' && Platform.isDesktopApp) {
			try {
				const electron = window.require('electron');
				const remote = electron.remote || window.require('@electron/remote');
				const fs = window.require('fs');
				const pathMod = window.require('path');
				const adapter = this.app.vault.adapter;
				const dir = adapter.getBasePath ? pathMod.join(adapter.getBasePath(), file.parent ? file.parent.path : '') : '';
				const res = await remote.dialog.showSaveDialog({
					defaultPath: pathMod.join(dir, name),
					filters: [{ name: ext === 'docx' ? 'Word document' : 'PowerPoint presentation', extensions: [ext] }],
				});
				if (res.canceled || !res.filePath) return null;
				await fs.promises.writeFile(res.filePath, Buffer.from(buffer));
				return { external: res.filePath, display: res.filePath };
			} catch (e) {
				console.warn('PDF Customiser: save dialog unavailable, saving next to the note', e);
			}
		}

		let folder = o.location === 'folder' ? normalizePath(o.folder || 'Exports') : file.parent ? file.parent.path : '';
		if (folder === '/' || folder === '.') folder = '';
		if (folder && !this.app.vault.getAbstractFileByPath(folder)) {
			try { await this.app.vault.createFolder(folder); } catch (e) { /* already exists */ }
		}
		const join = (n) => normalizePath(folder ? `${folder}/${n}` : n);
		let path = join(name);
		if (!o.overwrite) {
			let i = 2;
			while (this.app.vault.getAbstractFileByPath(path)) path = join(`${base} ${i++}.${ext}`);
		}
		const existing = this.app.vault.getAbstractFileByPath(path);
		if (existing instanceof TFile) await this.app.vault.modifyBinary(existing, buffer);
		else await this.app.vault.createBinary(path, buffer);
		return { vault: path, display: path };
	}

	openExternally(saved) {
		try {
			if (saved.vault && typeof this.app.openWithDefaultApp === 'function') { this.app.openWithDefaultApp(saved.vault); return; }
			if (!Platform.isDesktopApp) return;
			const { shell } = window.require('electron');
			let p = saved.external;
			if (!p && saved.vault && this.app.vault.adapter.getFullPath) p = this.app.vault.adapter.getFullPath(saved.vault);
			if (p) shell.openPath(p);
		} catch (e) {
			console.warn('PDF Customiser: could not open file', e);
		}
	}
}

module.exports = PdfCustomiserPlugin;
module.exports.default = PdfCustomiserPlugin;
// Exposed for testing; not used by Obsidian.
module.exports.internals = {
	DEFAULT_SETTINGS, TEMPLATES, parseMarkdown, parseInline, splitSlides, buildPrintCss, buildContentCss,
	computeSlots, expandTokens, DocxBuilder, PptxBuilder, createZip, imageInfo, deepMerge, clone, SAMPLE_HTML, SAMPLE_CTX,
};
