// Gera dist/control-hub.html: a página em arquivo único (CSS e JS embutidos) usada para publicar o artefato.
// Uso: node scripts/build-artifact.mjs
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => readFileSync(join(root, p), 'utf8');

const html = read('index.html');
const css = read('css/styles.css');
const js = read('js/app.js');

const title = html.match(/<title>[\s\S]*?<\/title>/)[0];
let body = html.match(/<body>([\s\S]*)<\/body>/)[1];

const scriptTag = '<script src="js/app.js"></script>';
if (!body.includes(scriptTag)) throw new Error(`index.html precisa conter ${scriptTag}`);
// Função no replace: evita que "$&", "$1" etc. do código sejam interpretados como padrões.
body = body.replace(scriptTag, () => `<script>\n${js}</script>`);

// O artefato envolve a página no próprio <html>/<head>/<body>: aqui vai só o conteúdo.
const out = `${title}\n<style>\n${css}</style>\n${body.trim()}\n`;

mkdirSync(join(root, 'dist'), { recursive: true });
writeFileSync(join(root, 'dist/control-hub.html'), out);
console.log(`dist/control-hub.html (${(out.length / 1024).toFixed(1)} KB)`);
