const fs = require('fs');
const path = require('path');
const root = '/Users/kartikeykandpal/Documents/Codex/2026-10-22/mak';
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
let passed = 0;
function check(condition, message) { if (!condition) throw new Error(message); passed++; }

const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]);
const links = [...html.matchAll(/href="#([^"]+)"/g)].map(m => m[1]);
check(new Set(ids).size === ids.length, 'Section IDs are unique');
check(links.every(id => ids.includes(id)), 'Every in-page navigation target exists');
check(/<meta name="viewport"/.test(html), 'Responsive viewport metadata exists');
check(/<title>[^<]+<\/title>/.test(html), 'Document title exists');

const handlers = new Map();
const attrs = {};
const nav = { open: false, classList: {
  toggle(name) { if (name !== 'open') throw new Error('Unexpected class'); this.owner.open = !this.owner.open; return this.owner.open; },
  remove(name) { if (name === 'open') this.owner.open = false; }, owner: null
} };
nav.classList.owner = nav;
const button = { textContent: '☰', setAttribute(k,v) { attrs[k] = v; }, addEventListener(k,fn) { handlers.set('button:'+k,fn); } };
const linksMock = [{addEventListener(k,fn){handlers.set('link:'+k,fn);}}];
nav.querySelectorAll = () => linksMock;
global.document = { querySelector(sel) { return sel === '.menu-toggle' ? button : nav; } };
global.document.querySelectorAll = () => linksMock;
require(path.join(root,'script.js'));
handlers.get('button:click')();
check(nav.open && attrs['aria-expanded'] === 'true', 'Menu opens and announces expanded state');
check(button.textContent === '×' && attrs['aria-label'] === 'Close navigation', 'Open state updates icon and label');
handlers.get('link:click')();
check(!nav.open && attrs['aria-expanded'] === 'false', 'Selecting a link closes the menu');
check(button.textContent === '☰' && attrs['aria-label'] === 'Open navigation', 'Closed state restores icon and label');
console.log(`${passed} checks passed; 0 failed.`);
