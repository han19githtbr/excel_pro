// ============================================
// MOTOR DE FÓRMULAS - Tokenizer + Parser + Avaliador
// Suporta: operadores com precedência, funções aninhadas,
// referências relativas/absolutas ($A$1), intervalos (A1:B10),
// referências entre planilhas (Planilha2!A1) e 65+ funções.
// ============================================

class FormulaError extends Error {
    constructor(code) {
        super(code);
        this.code = code;
    }
}

// ---------- TOKENIZER ----------
function tokenize(expr) {
    const tokens = [];
    let i = 0;
    const n = expr.length;
    while (i < n) {
        const c = expr[i];
        if (c === ' ' || c === '\t') { i++; continue; }
        if (c === '"') {
            let j = i + 1, str = '';
            while (j < n) {
                if (expr[j] === '"') {
                    if (expr[j + 1] === '"') { str += '"'; j += 2; continue; }
                    break;
                }
                str += expr[j];
                j++;
            }
            tokens.push({ type: 'STRING', value: str });
            i = j + 1;
            continue;
        }
        if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(expr[i + 1] || ''))) {
            let j = i;
            while (j < n && /[0-9.]/.test(expr[j])) j++;
            if (expr[j] === 'E' || expr[j] === 'e') {
                j++;
                if (expr[j] === '+' || expr[j] === '-') j++;
                while (j < n && /[0-9]/.test(expr[j])) j++;
            }
            tokens.push({ type: 'NUMBER', value: parseFloat(expr.substring(i, j)) });
            i = j;
            continue;
        }
        // sheet name / identifier / cell ref (supports 'Sheet Name'! too)
        if (/[A-Za-z_$']/.test(c)) {
            let j = i;
            if (c === "'") {
                j++;
                while (j < n && expr[j] !== "'") j++;
                j++; // skip closing quote
            } else {
                while (j < n && /[A-Za-z0-9_$.]/.test(expr[j])) j++;
            }
            let word = expr.substring(i, j);
            // sheet reference?
            if (expr[j] === '!') {
                let sheetName = word.replace(/^'|'$/g, '');
                j++;
                let k = j;
                while (k < n && /[A-Za-z0-9_$]/.test(expr[k])) k++;
                const cellPart = expr.substring(j, k);
                tokens.push({ type: 'REF', value: cellPart, sheet: sheetName });
                i = k;
                continue;
            }
            if (expr[j] === '(') {
                tokens.push({ type: 'FUNC', value: word.toUpperCase() });
                i = j;
                continue;
            }
            if (/^\$?[A-Za-z]{1,3}\$?[0-9]+$/.test(word)) {
                tokens.push({ type: 'REF', value: word.toUpperCase(), sheet: null });
                i = j;
                continue;
            }
            if (word.toUpperCase() === 'TRUE' || word.toUpperCase() === 'FALSE') {
                tokens.push({ type: 'BOOL', value: word.toUpperCase() === 'TRUE' });
                i = j;
                continue;
            }
            tokens.push({ type: 'IDENT', value: word });
            i = j;
            continue;
        }
        if (c === ':') { tokens.push({ type: 'COLON' }); i++; continue; }
        if (c === ',' || c === ';') { tokens.push({ type: 'COMMA' }); i++; continue; }
        if (c === '(') { tokens.push({ type: 'LPAREN' }); i++; continue; }
        if (c === ')') { tokens.push({ type: 'RPAREN' }); i++; continue; }
        if (c === '%') { tokens.push({ type: 'PERCENT' }); i++; continue; }
        if (c === '&') { tokens.push({ type: 'OP', value: '&' }); i++; continue; }
        if (c === '^') { tokens.push({ type: 'OP', value: '^' }); i++; continue; }
        if (c === '+' || c === '-' || c === '*' || c === '/') { tokens.push({ type: 'OP', value: c }); i++; continue; }
        if (c === '>' || c === '<') {
            if (expr[i + 1] === '=') { tokens.push({ type: 'OP', value: c + '=' }); i += 2; continue; }
            if (c === '<' && expr[i + 1] === '>') { tokens.push({ type: 'OP', value: '<>' }); i += 2; continue; }
            tokens.push({ type: 'OP', value: c }); i++; continue;
        }
        if (c === '=') { tokens.push({ type: 'OP', value: '=' }); i++; continue; }
        // unknown char, skip
        i++;
    }
    return tokens;
}

// ---------- PARSER (recursive descent -> AST) ----------
class Parser {
    constructor(tokens) {
        this.tokens = tokens;
        this.pos = 0;
    }
    peek() { return this.tokens[this.pos]; }
    next() { return this.tokens[this.pos++]; }
    expect(type) {
        const t = this.next();
        if (!t || t.type !== type) throw new FormulaError('#ERROR!');
        return t;
    }
    parse() {
        if (this.tokens.length === 0) return { type: 'Literal', value: '' };
        const node = this.parseComparison();
        return node;
    }
    parseComparison() {
        let left = this.parseConcat();
        while (this.peek() && this.peek().type === 'OP' && ['=', '<>', '<', '>', '<=', '>='].includes(this.peek().value)) {
            const op = this.next().value;
            const right = this.parseConcat();
            left = { type: 'BinOp', op, left, right };
        }
        return left;
    }
    parseConcat() {
        let left = this.parseAdditive();
        while (this.peek() && this.peek().type === 'OP' && this.peek().value === '&') {
            this.next();
            const right = this.parseAdditive();
            left = { type: 'BinOp', op: '&', left, right };
        }
        return left;
    }
    parseAdditive() {
        let left = this.parseTerm();
        while (this.peek() && this.peek().type === 'OP' && (this.peek().value === '+' || this.peek().value === '-')) {
            const op = this.next().value;
            const right = this.parseTerm();
            left = { type: 'BinOp', op, left, right };
        }
        return left;
    }
    parseTerm() {
        let left = this.parseFactor();
        while (this.peek() && this.peek().type === 'OP' && (this.peek().value === '*' || this.peek().value === '/')) {
            const op = this.next().value;
            const right = this.parseFactor();
            left = { type: 'BinOp', op, left, right };
        }
        return left;
    }
    parseFactor() {
        let left = this.parseUnary();
        while (this.peek() && this.peek().type === 'OP' && this.peek().value === '^') {
            this.next();
            const right = this.parseUnary();
            left = { type: 'BinOp', op: '^', left, right };
        }
        return left;
    }
    parseUnary() {
        if (this.peek() && this.peek().type === 'OP' && (this.peek().value === '-' || this.peek().value === '+')) {
            const op = this.next().value;
            const operand = this.parseUnary();
            return { type: 'UnaryOp', op, operand };
        }
        return this.parsePostfix();
    }
    parsePostfix() {
        let node = this.parseRange();
        while (this.peek() && this.peek().type === 'PERCENT') {
            this.next();
            node = { type: 'Percent', operand: node };
        }
        return node;
    }
    parseRange() {
        let left = this.parsePrimary();
        if (this.peek() && this.peek().type === 'COLON') {
            this.next();
            const right = this.parsePrimary();
            if (left.type === 'Ref' && right.type === 'Ref') {
                return { type: 'Range', from: left, to: right };
            }
            throw new FormulaError('#REF!');
        }
        return left;
    }
    parsePrimary() {
        const t = this.peek();
        if (!t) throw new FormulaError('#ERROR!');
        if (t.type === 'NUMBER') { this.next(); return { type: 'Literal', value: t.value }; }
        if (t.type === 'STRING') { this.next(); return { type: 'Literal', value: t.value }; }
        if (t.type === 'BOOL') { this.next(); return { type: 'Literal', value: t.value }; }
        if (t.type === 'REF') { this.next(); return { type: 'Ref', ref: t.value, sheet: t.sheet }; }
        if (t.type === 'LPAREN') {
            this.next();
            const node = this.parseComparison();
            this.expect('RPAREN');
            return node;
        }
        if (t.type === 'FUNC') {
            this.next();
            this.expect('LPAREN');
            const args = [];
            if (this.peek() && this.peek().type !== 'RPAREN') {
                args.push(this.parseComparison());
                while (this.peek() && this.peek().type === 'COMMA') {
                    this.next();
                    args.push(this.parseComparison());
                }
            }
            this.expect('RPAREN');
            return { type: 'Call', name: t.value, args };
        }
        if (t.type === 'IDENT') {
            // unrecognized bareword -> treat as text/error
            this.next();
            return { type: 'Literal', value: '#NAME?' , isError: true};
        }
        throw new FormulaError('#ERROR!');
    }
}

// ---------- REFERENCE HELPERS ----------
function colLettersToIndex(letters) {
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
        col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return col - 1;
}

function parseRefString(ref) {
    const m = ref.match(/^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/);
    if (!m) return null;
    return {
        col: colLettersToIndex(m[2].toUpperCase()),
        row: parseInt(m[4], 10) - 1,
        colAbs: m[1] === '$',
        rowAbs: m[3] === '$'
    };
}

// ---------- EVALUATOR ----------
// context: {
//   currentSheet: string,
//   getCellRaw(sheet,row,col) -> raw string/value stored in cell,
//   getCellComputed(sheet,row,col) -> evaluated display value (number/string), handles its own memo+circular guard,
//   sheetExists(name) -> bool
// }
class FormulaEvaluator {
    constructor(context) {
        this.ctx = context;
    }

    evaluate(formulaText) {
        try {
            const expr = formulaText.substring(1); // remove '='
            const tokens = tokenize(expr);
            const parser = new Parser(tokens);
            const ast = parser.parse();
            const result = this.evalNode(ast);
            return this.unwrap(result);
        } catch (e) {
            if (e instanceof FormulaError) return e.code;
            if (e && e.code) return e.code;
            return '#ERROR!';
        }
    }

    unwrap(v) {
        if (v && v.__range) {
            return v.values.length ? v.values[0] : '';
        }
        return v;
    }

    resolveCellValue(sheet, row, col) {
        const s = sheet || this.ctx.currentSheet;
        return this.ctx.getCellComputed(s, row, col);
    }

    evalNode(node) {
        switch (node.type) {
            case 'Literal':
                if (node.isError) throw new FormulaError(node.value);
                return node.value;
            case 'Ref': {
                const parsed = parseRefString(node.ref);
                if (!parsed) throw new FormulaError('#REF!');
                if (node.sheet && !this.ctx.sheetExists(node.sheet)) throw new FormulaError('#REF!');
                return this.resolveCellValue(node.sheet, parsed.row, parsed.col);
            }
            case 'Range': {
                const from = parseRefString(node.from.ref);
                const to = parseRefString(node.to.ref);
                if (!from || !to) throw new FormulaError('#REF!');
                const sheet = node.from.sheet || node.to.sheet;
                if (sheet && !this.ctx.sheetExists(sheet)) throw new FormulaError('#REF!');
                const r1 = Math.min(from.row, to.row), r2 = Math.max(from.row, to.row);
                const c1 = Math.min(from.col, to.col), c2 = Math.max(from.col, to.col);
                const values = [];
                const cells = [];
                for (let r = r1; r <= r2; r++) {
                    for (let c = c1; c <= c2; c++) {
                        values.push(this.resolveCellValue(sheet, r, c));
                        cells.push({ row: r, col: c, sheet: sheet || this.ctx.currentSheet });
                    }
                }
                return { __range: true, values, cells, rows: r2 - r1 + 1, cols: c2 - c1 + 1, r1, c1, r2, c2 };
            }
            case 'UnaryOp': {
                const v = toNumber(this.unwrap(this.evalNode(node.operand)));
                return node.op === '-' ? -v : +v;
            }
            case 'Percent': {
                const v = toNumber(this.unwrap(this.evalNode(node.operand)));
                return v / 100;
            }
            case 'BinOp':
                return this.evalBinOp(node);
            case 'Call':
                return this.evalCall(node);
            default:
                throw new FormulaError('#ERROR!');
        }
    }

    evalBinOp(node) {
        const op = node.op;
        if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
            const l = this.unwrap(this.evalNode(node.left));
            const r = this.unwrap(this.evalNode(node.right));
            return compareValues(l, r, op);
        }
        if (op === '&') {
            const l = this.unwrap(this.evalNode(node.left));
            const r = this.unwrap(this.evalNode(node.right));
            return toDisplayString(l) + toDisplayString(r);
        }
        const l = toNumber(this.unwrap(this.evalNode(node.left)));
        const r = toNumber(this.unwrap(this.evalNode(node.right)));
        switch (op) {
            case '+': return l + r;
            case '-': return l - r;
            case '*': return l * r;
            case '/': if (r === 0) throw new FormulaError('#DIV/0!'); return l / r;
            case '^': return Math.pow(l, r);
        }
        throw new FormulaError('#ERROR!');
    }

    evalCall(node) {
        const fn = FUNCTIONS[node.name];
        if (!fn) throw new FormulaError('#NAME?');
        return fn(node.args, this);
    }

    // helpers exposed to functions
    val(argNode) { return this.unwrap(this.evalNode(argNode)); }
    flat(argNode) {
        const v = this.evalNode(argNode);
        if (v && v.__range) return v.values;
        return [this.unwrap(v)];
    }
    range(argNode) {
        const v = this.evalNode(argNode);
        if (v && v.__range) return v;
        // single cell as 1x1 range
        const val = this.unwrap(v);
        return { __range: true, values: [val], cells: [], rows: 1, cols: 1 };
    }
}

// ---------- TYPE COERCION ----------
function toNumber(v) {
    if (v === '' || v === null || v === undefined) return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v).replace(',', '.'));
    return isNaN(n) ? NaN : n;
}
function isNumeric(v) {
    if (typeof v === 'number') return !isNaN(v);
    if (v === '' || v === null || v === undefined) return false;
    if (typeof v === 'string' && /^-?[\d.,]+$/.test(v.trim())) return !isNaN(parseFloat(v.replace(',', '.')));
    return false;
}
function toDisplayString(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'VERDADEIRO' : 'FALSO';
    return String(v);
}
function compareValues(l, r, op) {
    const bothNumeric = isNumeric(l) && isNumeric(r);
    let cmp;
    if (bothNumeric) {
        const ln = toNumber(l), rn = toNumber(r);
        cmp = ln < rn ? -1 : ln > rn ? 1 : 0;
    } else {
        const ls = toDisplayString(l).toLowerCase(), rs = toDisplayString(r).toLowerCase();
        cmp = ls < rs ? -1 : ls > rs ? 1 : 0;
    }
    switch (op) {
        case '=': return cmp === 0;
        case '<>': return cmp !== 0;
        case '<': return cmp < 0;
        case '>': return cmp > 0;
        case '<=': return cmp <= 0;
        case '>=': return cmp >= 0;
    }
}

function matchesCriteria(value, criteria) {
    if (typeof criteria === 'string' && /^(<=|>=|<>|<|>|=)/.test(criteria)) {
        const m = criteria.match(/^(<=|>=|<>|<|>|=)(.*)$/);
        const op = m[1];
        let critVal = m[2];
        const num = parseFloat(critVal);
        const critIsNum = !isNaN(num) && critVal.trim() !== '';
        return compareValues(value, critIsNum ? num : critVal, op);
    }
    if (typeof criteria === 'string' && (criteria.includes('*') || criteria.includes('?'))) {
        const pattern = '^' + criteria.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$';
        return new RegExp(pattern, 'i').test(toDisplayString(value));
    }
    if (isNumeric(criteria) && isNumeric(value)) return toNumber(value) === toNumber(criteria);
    return toDisplayString(value).toLowerCase() === toDisplayString(criteria).toLowerCase();
}

function numArray(vals) {
    return vals.filter(v => isNumeric(v) && v !== '').map(toNumber);
}

// ---------- FUNCTION LIBRARY ----------
const FUNCTIONS = {};

function reg(name, fn) { FUNCTIONS[name] = fn; }

// -- Math --
reg('SUM', (args, ev) => args.reduce((sum, a) => sum + numArray(ev.flat(a)).reduce((x, y) => x + y, 0), 0));
reg('PRODUCT', (args, ev) => {
    let nums = [];
    args.forEach(a => nums = nums.concat(numArray(ev.flat(a))));
    return nums.length ? nums.reduce((x, y) => x * y, 1) : 0;
});
reg('ROUND', (args, ev) => {
    const n = toNumber(ev.val(args[0])), d = args[1] ? toNumber(ev.val(args[1])) : 0;
    const f = Math.pow(10, d);
    return Math.round(n * f) / f;
});
reg('ROUNDUP', (args, ev) => {
    const n = toNumber(ev.val(args[0])), d = args[1] ? toNumber(ev.val(args[1])) : 0;
    const f = Math.pow(10, d);
    return (n >= 0 ? Math.ceil(n * f) : Math.floor(n * f)) / f;
});
reg('ROUNDDOWN', (args, ev) => {
    const n = toNumber(ev.val(args[0])), d = args[1] ? toNumber(ev.val(args[1])) : 0;
    const f = Math.pow(10, d);
    return (n >= 0 ? Math.floor(n * f) : Math.ceil(n * f)) / f;
});
reg('ABS', (args, ev) => Math.abs(toNumber(ev.val(args[0]))));
reg('SQRT', (args, ev) => { const n = toNumber(ev.val(args[0])); if (n < 0) throw new FormulaError('#NUM!'); return Math.sqrt(n); });
reg('POWER', (args, ev) => Math.pow(toNumber(ev.val(args[0])), toNumber(ev.val(args[1]))));
reg('MOD', (args, ev) => { const a = toNumber(ev.val(args[0])), b = toNumber(ev.val(args[1])); if (b === 0) throw new FormulaError('#DIV/0!'); return a - b * Math.floor(a / b); });
reg('INT', (args, ev) => Math.floor(toNumber(ev.val(args[0]))));
reg('TRUNC', (args, ev) => Math.trunc(toNumber(ev.val(args[0]))));
reg('PI', () => Math.PI);
reg('RAND', () => Math.random());
reg('RANDBETWEEN', (args, ev) => {
    const a = Math.ceil(toNumber(ev.val(args[0]))), b = Math.floor(toNumber(ev.val(args[1])));
    return Math.floor(Math.random() * (b - a + 1)) + a;
});
reg('CEILING', (args, ev) => { const n = toNumber(ev.val(args[0])), s = args[1] ? toNumber(ev.val(args[1])) : 1; return s === 0 ? 0 : Math.ceil(n / s) * s; });
reg('FLOOR', (args, ev) => { const n = toNumber(ev.val(args[0])), s = args[1] ? toNumber(ev.val(args[1])) : 1; return s === 0 ? 0 : Math.floor(n / s) * s; });
reg('SIGN', (args, ev) => Math.sign(toNumber(ev.val(args[0]))));
reg('EXP', (args, ev) => Math.exp(toNumber(ev.val(args[0]))));
reg('LN', (args, ev) => Math.log(toNumber(ev.val(args[0]))));
reg('LOG', (args, ev) => { const n = toNumber(ev.val(args[0])), b = args[1] ? toNumber(ev.val(args[1])) : 10; return Math.log(n) / Math.log(b); });
reg('LOG10', (args, ev) => Math.log10(toNumber(ev.val(args[0]))));
reg('SIN', (args, ev) => Math.sin(toNumber(ev.val(args[0]))));
reg('COS', (args, ev) => Math.cos(toNumber(ev.val(args[0]))));
reg('TAN', (args, ev) => Math.tan(toNumber(ev.val(args[0]))));

// -- Statistics --
reg('AVERAGE', (args, ev) => {
    let nums = [];
    args.forEach(a => nums = nums.concat(numArray(ev.flat(a))));
    if (!nums.length) throw new FormulaError('#DIV/0!');
    return nums.reduce((x, y) => x + y, 0) / nums.length;
});
reg('MEDIAN', (args, ev) => {
    let nums = [];
    args.forEach(a => nums = nums.concat(numArray(ev.flat(a))));
    if (!nums.length) throw new FormulaError('#NUM!');
    nums.sort((a, b) => a - b);
    const mid = Math.floor(nums.length / 2);
    return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
});
reg('MAX', (args, ev) => { let nums = []; args.forEach(a => nums = nums.concat(numArray(ev.flat(a)))); return nums.length ? Math.max(...nums) : 0; });
reg('MIN', (args, ev) => { let nums = []; args.forEach(a => nums = nums.concat(numArray(ev.flat(a)))); return nums.length ? Math.min(...nums) : 0; });
reg('COUNT', (args, ev) => { let nums = []; args.forEach(a => nums = nums.concat(numArray(ev.flat(a)))); return nums.length; });
reg('COUNTA', (args, ev) => { let all = []; args.forEach(a => all = all.concat(ev.flat(a))); return all.filter(v => v !== '' && v !== null && v !== undefined).length; });
reg('COUNTBLANK', (args, ev) => { let all = ev.flat(args[0]); return all.filter(v => v === '' || v === null || v === undefined).length; });
reg('STDEV', (args, ev) => {
    let nums = []; args.forEach(a => nums = nums.concat(numArray(ev.flat(a))));
    if (nums.length < 2) throw new FormulaError('#DIV/0!');
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    const variance = nums.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (nums.length - 1);
    return Math.sqrt(variance);
});
reg('VAR', (args, ev) => {
    let nums = []; args.forEach(a => nums = nums.concat(numArray(ev.flat(a))));
    if (nums.length < 2) throw new FormulaError('#DIV/0!');
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return nums.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (nums.length - 1);
});
reg('RANK', (args, ev) => {
    const val = toNumber(ev.val(args[0]));
    const nums = numArray(ev.flat(args[1]));
    const desc = args[2] ? toNumber(ev.val(args[2])) === 0 : true;
    const sorted = [...nums].sort((a, b) => desc ? b - a : a - b);
    const idx = sorted.indexOf(val);
    if (idx === -1) throw new FormulaError('#N/A');
    return idx + 1;
});
reg('LARGE', (args, ev) => { const nums = numArray(ev.flat(args[0])).sort((a, b) => b - a); const k = toNumber(ev.val(args[1])); if (k < 1 || k > nums.length) throw new FormulaError('#NUM!'); return nums[k - 1]; });
reg('SMALL', (args, ev) => { const nums = numArray(ev.flat(args[0])).sort((a, b) => a - b); const k = toNumber(ev.val(args[1])); if (k < 1 || k > nums.length) throw new FormulaError('#NUM!'); return nums[k - 1]; });
reg('MODE', (args, ev) => {
    const nums = numArray(ev.flat(args[0]));
    const freq = {};
    let best = null, bestCount = 0;
    nums.forEach(n => { freq[n] = (freq[n] || 0) + 1; if (freq[n] > bestCount) { bestCount = freq[n]; best = n; } });
    if (best === null) throw new FormulaError('#N/A');
    return best;
});

// -- Logic --
reg('IF', (args, ev) => {
    const cond = ev.val(args[0]);
    const truthy = toBool(cond);
    if (truthy) return args[1] !== undefined ? ev.val(args[1]) : true;
    return args[2] !== undefined ? ev.val(args[2]) : false;
});
reg('AND', (args, ev) => args.every(a => toBool(ev.val(a))));
reg('OR', (args, ev) => args.some(a => toBool(ev.val(a))));
reg('NOT', (args, ev) => !toBool(ev.val(args[0])));
reg('IFERROR', (args, ev) => {
    try { const v = ev.val(args[0]); if (typeof v === 'string' && /^#[A-Z/!?0-9]+$/.test(v)) return ev.val(args[1]); return v; }
    catch (e) { return ev.val(args[1]); }
});
reg('IFNA', (args, ev) => {
    try { const v = ev.val(args[0]); if (v === '#N/A') return ev.val(args[1]); return v; }
    catch (e) { return ev.val(args[1]); }
});
reg('ISERROR', (args, ev) => {
    try { const v = ev.val(args[0]); return typeof v === 'string' && /^#[A-Z/!?0-9]+$/.test(v); }
    catch (e) { return true; }
});
reg('ISBLANK', (args, ev) => { const v = ev.val(args[0]); return v === '' || v === null || v === undefined; });
reg('ISNUMBER', (args, ev) => isNumeric(ev.val(args[0])));
reg('ISTEXT', (args, ev) => { const v = ev.val(args[0]); return typeof v === 'string' && !isNumeric(v); });
reg('TRUE', () => true);
reg('FALSE', () => false);
function toBool(v) {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') return v.toUpperCase() === 'VERDADEIRO' || v.toUpperCase() === 'TRUE' || (isNumeric(v) && toNumber(v) !== 0);
    return !!v;
}

// -- Text --
reg('CONCATENATE', (args, ev) => args.map(a => toDisplayString(ev.val(a))).join(''));
reg('CONCAT', (args, ev) => { let out = ''; args.forEach(a => ev.flat(a).forEach(v => out += toDisplayString(v))); return out; });
reg('LEFT', (args, ev) => toDisplayString(ev.val(args[0])).substring(0, args[1] ? toNumber(ev.val(args[1])) : 1));
reg('RIGHT', (args, ev) => { const s = toDisplayString(ev.val(args[0])); const n = args[1] ? toNumber(ev.val(args[1])) : 1; return s.substring(Math.max(0, s.length - n)); });
reg('MID', (args, ev) => { const s = toDisplayString(ev.val(args[0])); const start = toNumber(ev.val(args[1])); const len = toNumber(ev.val(args[2])); return s.substring(start - 1, start - 1 + len); });
reg('LEN', (args, ev) => toDisplayString(ev.val(args[0])).length);
reg('UPPER', (args, ev) => toDisplayString(ev.val(args[0])).toUpperCase());
reg('LOWER', (args, ev) => toDisplayString(ev.val(args[0])).toLowerCase());
reg('PROPER', (args, ev) => toDisplayString(ev.val(args[0])).replace(/\w\S*/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()));
reg('TRIM', (args, ev) => toDisplayString(ev.val(args[0])).trim().replace(/\s+/g, ' '));
reg('SUBSTITUTE', (args, ev) => {
    const s = toDisplayString(ev.val(args[0]));
    const oldT = toDisplayString(ev.val(args[1]));
    const newT = toDisplayString(ev.val(args[2]));
    if (args[3]) {
        const occ = toNumber(ev.val(args[3]));
        let count = 0;
        return s.replace(new RegExp(oldT.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), m => { count++; return count === occ ? newT : m; });
    }
    return s.split(oldT).join(newT);
});
reg('REPLACE', (args, ev) => {
    const s = toDisplayString(ev.val(args[0]));
    const start = toNumber(ev.val(args[1]));
    const len = toNumber(ev.val(args[2]));
    const newT = toDisplayString(ev.val(args[3]));
    return s.substring(0, start - 1) + newT + s.substring(start - 1 + len);
});
reg('FIND', (args, ev) => {
    const find = toDisplayString(ev.val(args[0]));
    const within = toDisplayString(ev.val(args[1]));
    const start = args[2] ? toNumber(ev.val(args[2])) : 1;
    const idx = within.indexOf(find, start - 1);
    if (idx === -1) throw new FormulaError('#VALUE!');
    return idx + 1;
});
reg('SEARCH', (args, ev) => {
    const find = toDisplayString(ev.val(args[0])).toLowerCase();
    const within = toDisplayString(ev.val(args[1])).toLowerCase();
    const start = args[2] ? toNumber(ev.val(args[2])) : 1;
    const idx = within.indexOf(find, start - 1);
    if (idx === -1) throw new FormulaError('#VALUE!');
    return idx + 1;
});
reg('REPT', (args, ev) => toDisplayString(ev.val(args[0])).repeat(Math.max(0, toNumber(ev.val(args[1])))));
reg('EXACT', (args, ev) => toDisplayString(ev.val(args[0])) === toDisplayString(ev.val(args[1])));
reg('TEXT', (args, ev) => {
    const v = ev.val(args[0]);
    const fmt = toDisplayString(ev.val(args[1]));
    const n = toNumber(v);
    if (isNaN(n)) return toDisplayString(v);
    if (/0\.0+/.test(fmt)) { const dec = (fmt.split('.')[1] || '').length; return n.toFixed(dec); }
    if (fmt.includes('%')) return (n * 100).toFixed(2) + '%';
    if (fmt.includes('#,##0') || fmt.includes('0,0')) return new Intl.NumberFormat('pt-BR').format(n);
    return String(n);
});
reg('VALUE', (args, ev) => { const n = toNumber(ev.val(args[0])); if (isNaN(n)) throw new FormulaError('#VALUE!'); return n; });

// -- Lookup --
reg('VLOOKUP', (args, ev) => {
    const lookupVal = ev.val(args[0]);
    const range = ev.range(args[1]);
    const colIndex = toNumber(ev.val(args[2]));
    const exact = args[3] ? !toBool(ev.val(args[3])) : false;
    if (colIndex < 1 || colIndex > range.cols) throw new FormulaError('#REF!');
    for (let r = 0; r < range.rows; r++) {
        const cellVal = range.values[r * range.cols];
        const isMatch = exact ? matchesCriteria(cellVal, lookupVal) : (String(cellVal).toLowerCase() === String(lookupVal).toLowerCase() || toNumber(cellVal) === toNumber(lookupVal));
        if (isMatch) return range.values[r * range.cols + (colIndex - 1)];
    }
    throw new FormulaError('#N/A');
});
reg('HLOOKUP', (args, ev) => {
    const lookupVal = ev.val(args[0]);
    const range = ev.range(args[1]);
    const rowIndex = toNumber(ev.val(args[2]));
    if (rowIndex < 1 || rowIndex > range.rows) throw new FormulaError('#REF!');
    for (let c = 0; c < range.cols; c++) {
        const cellVal = range.values[c];
        if (String(cellVal).toLowerCase() === String(lookupVal).toLowerCase() || toNumber(cellVal) === toNumber(lookupVal)) {
            return range.values[(rowIndex - 1) * range.cols + c];
        }
    }
    throw new FormulaError('#N/A');
});
reg('INDEX', (args, ev) => {
    const range = ev.range(args[0]);
    const rowIndex = args[1] ? toNumber(ev.val(args[1])) : 1;
    const colIndex = args[2] ? toNumber(ev.val(args[2])) : 1;
    if (range.cols === 1 && !args[2]) {
        if (rowIndex < 1 || rowIndex > range.rows) throw new FormulaError('#REF!');
        return range.values[rowIndex - 1];
    }
    if (rowIndex < 1 || rowIndex > range.rows || colIndex < 1 || colIndex > range.cols) throw new FormulaError('#REF!');
    return range.values[(rowIndex - 1) * range.cols + (colIndex - 1)];
});
reg('MATCH', (args, ev) => {
    const lookupVal = ev.val(args[0]);
    const arr = ev.flat(args[1]);
    const matchType = args[2] ? toNumber(ev.val(args[2])) : 1;
    if (matchType === 0) {
        const idx = arr.findIndex(v => matchesCriteria(v, lookupVal));
        if (idx === -1) throw new FormulaError('#N/A');
        return idx + 1;
    }
    // approximate match (assume sorted ascending for 1, descending for -1)
    let best = -1;
    for (let i = 0; i < arr.length; i++) {
        if (matchType === 1 && toNumber(arr[i]) <= toNumber(lookupVal)) best = i;
        if (matchType === -1 && toNumber(arr[i]) >= toNumber(lookupVal)) best = i;
    }
    if (best === -1) throw new FormulaError('#N/A');
    return best + 1;
});
reg('CHOOSE', (args, ev) => {
    const idx = toNumber(ev.val(args[0]));
    if (idx < 1 || idx >= args.length) throw new FormulaError('#VALUE!');
    return ev.val(args[idx]);
});

// -- Conditional aggregates --
reg('SUMIF', (args, ev) => {
    const range = ev.range(args[0]);
    const criteria = ev.val(args[1]);
    const sumRange = args[2] ? ev.range(args[2]) : range;
    let sum = 0;
    for (let i = 0; i < range.values.length; i++) {
        if (matchesCriteria(range.values[i], criteria)) {
            const sv = sumRange.values[i];
            if (isNumeric(sv)) sum += toNumber(sv);
        }
    }
    return sum;
});
reg('COUNTIF', (args, ev) => {
    const range = ev.range(args[0]);
    const criteria = ev.val(args[1]);
    return range.values.filter(v => matchesCriteria(v, criteria)).length;
});
reg('AVERAGEIF', (args, ev) => {
    const range = ev.range(args[0]);
    const criteria = ev.val(args[1]);
    const avgRange = args[2] ? ev.range(args[2]) : range;
    let sum = 0, count = 0;
    for (let i = 0; i < range.values.length; i++) {
        if (matchesCriteria(range.values[i], criteria)) {
            const sv = avgRange.values[i];
            if (isNumeric(sv)) { sum += toNumber(sv); count++; }
        }
    }
    if (!count) throw new FormulaError('#DIV/0!');
    return sum / count;
});
reg('SUMIFS', (args, ev) => {
    const sumRange = ev.range(args[0]);
    const critPairs = [];
    for (let i = 1; i < args.length; i += 2) critPairs.push([ev.range(args[i]), ev.val(args[i + 1])]);
    let sum = 0;
    for (let i = 0; i < sumRange.values.length; i++) {
        if (critPairs.every(([r, c]) => matchesCriteria(r.values[i], c))) {
            if (isNumeric(sumRange.values[i])) sum += toNumber(sumRange.values[i]);
        }
    }
    return sum;
});
reg('COUNTIFS', (args, ev) => {
    const critPairs = [];
    for (let i = 0; i < args.length; i += 2) critPairs.push([ev.range(args[i]), ev.val(args[i + 1])]);
    const len = critPairs[0][0].values.length;
    let count = 0;
    for (let i = 0; i < len; i++) {
        if (critPairs.every(([r, c]) => matchesCriteria(r.values[i], c))) count++;
    }
    return count;
});

// -- Date/Time --
reg('TODAY', () => { const d = new Date(); return d.toLocaleDateString('pt-BR'); });
reg('NOW', () => { const d = new Date(); return d.toLocaleString('pt-BR'); });
reg('DATE', (args, ev) => {
    const y = toNumber(ev.val(args[0])), m = toNumber(ev.val(args[1])), d = toNumber(ev.val(args[2]));
    return new Date(y, m - 1, d).toLocaleDateString('pt-BR');
});
function parseBRDate(v) {
    if (v instanceof Date) return v;
    const s = toDisplayString(v);
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) return new Date(parseInt(m[3].length === 2 ? '20' + m[3] : m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
}
reg('YEAR', (args, ev) => { const d = parseBRDate(ev.val(args[0])); if (!d) throw new FormulaError('#VALUE!'); return d.getFullYear(); });
reg('MONTH', (args, ev) => { const d = parseBRDate(ev.val(args[0])); if (!d) throw new FormulaError('#VALUE!'); return d.getMonth() + 1; });
reg('DAY', (args, ev) => { const d = parseBRDate(ev.val(args[0])); if (!d) throw new FormulaError('#VALUE!'); return d.getDate(); });
reg('WEEKDAY', (args, ev) => { const d = parseBRDate(ev.val(args[0])); if (!d) throw new FormulaError('#VALUE!'); return d.getDay() + 1; });

if (typeof window !== 'undefined') {
    window.FormulaEvaluator = FormulaEvaluator;
    window.FormulaError = FormulaError;
}
