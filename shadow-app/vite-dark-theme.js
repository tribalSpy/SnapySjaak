// Night mode for the whole app, made from src/styles.css at build time: every
// rule that sets a colour gets a twin under :root[data-theme="dark"] with
// that colour turned around (light <-> dark, the same hue), so a new rule in
// styles.css is in the night mode too without writing it twice. Shadows stay
// dark. The lamp button (main.jsx) sets data-theme on <html>.
import postcss from "postcss";

const DARK = ':root[data-theme="dark"]';
const COLOR_RE = /#[0-9a-fA-F]{8}\b|#[0-9a-fA-F]{6}\b|#[0-9a-fA-F]{3,4}\b|rgba?\([^)]*\)|\b(?:white|black)\b/g;

function parseColor(text) {
  const value = text.toLowerCase();
  if (value === "white") return [255, 255, 255, 1];
  if (value === "black") return [0, 0, 0, 1];
  if (value.startsWith("#")) {
    let hex = value.slice(1);
    if (hex.length <= 4) hex = hex.split("").map((char) => char + char).join("");
    const alpha = hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1;
    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16), alpha];
  }
  const parts = value.replace(/rgba?\(|\)/g, "").split(/[\s,/]+/).filter(Boolean).map(Number);
  if (parts.length < 3 || parts.some((part) => !Number.isFinite(part))) return null;
  return [parts[0], parts[1], parts[2], parts[3] ?? 1];
}

function toHsl([red, green, blue]) {
  const r = red / 255;
  const g = green / 255;
  const b = blue / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const light = (max + min) / 2;
  if (max === min) return [0, 0, light];
  const delta = max - min;
  const sat = light > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  const hue = max === r ? (g - b) / delta + (g < b ? 6 : 0) : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4;
  return [hue / 6, sat, light];
}

function fromHsl([hue, sat, light]) {
  if (!sat) return [light, light, light].map((value) => Math.round(value * 255));
  const q = light < 0.5 ? light * (1 + sat) : light + sat - light * sat;
  const p = 2 * light - q;
  const channel = (t) => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return [channel(hue + 1 / 3), channel(hue), channel(hue - 1 / 3)].map((value) => Math.round(value * 255));
}

// The pale end is not simply turned around: a white card (1.00) stays a bit
// lighter than the page (#f4f6f8, 0.96), pale tints / hovers (0.93) a bit
// lighter again, borders (0.87) lighter still -- as cards on a dark page.
const PALE_POINTS = [[0.85, 0.287], [0.87, 0.3], [0.93, 0.2], [0.96, 0.13], [1, 0.17]];

function nightLightness(light) {
  if (light < PALE_POINTS[0][0]) return 0.95 - 0.78 * light;
  for (let index = 1; index < PALE_POINTS.length; index += 1) {
    const [x1, y1] = PALE_POINTS[index - 1];
    const [x2, y2] = PALE_POINTS[index];
    if (light <= x2) return y1 + ((light - x1) / (x2 - x1)) * (y2 - y1);
  }
  return PALE_POINTS[PALE_POINTS.length - 1][1];
}

// Light <-> dark with the same hue: white surfaces become #1b2026-ish, dark
// ink becomes light ink, a green / red accent stays green / red but lighter
// so it reads on the dark surface. Pale tints lose some colour.
function nightColor(text, { shadow = false } = {}) {
  const color = parseColor(text);
  if (!color) return text;
  const alpha = color[3];
  if (shadow) return `rgba(0, 0, 0, ${Math.min(1, Math.round(alpha * 2.2 * 100) / 100)})`;
  const [hue, sat, light] = toHsl(color);
  const nextLight = nightLightness(light);
  const nextSat = light > 0.85 ? sat * 0.55 : sat * 0.9;
  const [r, g, b] = fromHsl([hue, nextSat, nextLight]);
  if (alpha < 1) return `rgba(${r}, ${g}, ${b}, ${alpha})`;
  return `#${[r, g, b].map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

function darkSelector(selector) {
  return selector.split(",").map((part) => {
    const trimmed = part.trim();
    if (trimmed.startsWith(":root")) return `${DARK}${trimmed.slice(5)}`;
    if (/^(html|body)\b/.test(trimmed)) return `${DARK} ${trimmed.replace(/^html\b\s*/, "")}`.trim();
    return `${DARK} ${trimmed}`;
  }).join(", ");
}

export function buildNightCss(css) {
  const source = postcss.parse(css);
  const out = postcss.root();
  out.append(postcss.rule({ selector: DARK, nodes: [postcss.decl({ prop: "color-scheme", value: "dark" })] }));
  out.append(postcss.rule({ selector: `${DARK} body`, nodes: [postcss.decl({ prop: "background", value: nightColor("#f4f6f8") })] }));
  source.walkRules((rule) => {
    if (rule.parent?.type === "atrule" && /keyframes/i.test(rule.parent.name)) return;
    const decls = [];
    rule.each((node) => {
      if (node.type !== "decl") return;
      COLOR_RE.lastIndex = 0;
      if (!COLOR_RE.test(node.value)) return;
      const shadow = /shadow/i.test(node.prop);
      decls.push(postcss.decl({ prop: node.prop, value: node.value.replace(COLOR_RE, (match) => nightColor(match, { shadow })), important: node.important }));
    });
    if (!decls.length) return;
    const twin = postcss.rule({ selector: darkSelector(rule.selector), nodes: decls });
    // Keep the rule inside its @media / @supports.
    let target = twin;
    for (let parent = rule.parent; parent && parent.type === "atrule"; parent = parent.parent) {
      target = postcss.atRule({ name: parent.name, params: parent.params, nodes: [target] });
    }
    out.append(target);
  });
  return out.toString();
}

export default function darkThemePlugin() {
  return {
    name: "snappysjaak-night-mode",
    enforce: "pre",
    transform(code, id) {
      if (!/[\\/]src[\\/]styles\.css$/.test(id.split("?")[0])) return null;
      return { code: `${code}\n\n/* ---- Night mode (made from the rules above, see vite-dark-theme.js) ---- */\n${buildNightCss(code)}\n`, map: null };
    },
  };
}
