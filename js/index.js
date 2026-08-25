// ============================================
// CLASSE PRINCIPAL - EXCEL SIMULATOR PRO
// ============================================
class ExcelApplication {
    constructor() {
        this.state = {
            sheets: [],           // [{ name, data, styles, mergedCells, rowCount, colCount, colWidths, rowHeights, comments, validations, conditionalFormats, charts, freeze:{row:false,col:false}, showGridlines:true }]
            activeSheet: 0,
            selection: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
            selectedCell: { row: 0, col: 0 },
            clipboard: null,      // { mode:'copy'|'cut', sheetIndex, cells:[{row,col,value,style}], anchorRow, anchorCol }
            undoStack: [],
            redoStack: [],
            maxUndo: 150,
            isEditing: false,
            isSelecting: false,
            isFilling: false,
            zoom: 100,
            showFormulas: false,
            formatPainter: null,
            orientation: 'portrait'
        };

        this.DOM = {};
        this.recalcVisiting = new Set();
        this.initDOM();
        this.initWorkbook();
        this.buildFunctionList();
        this.render();
        this.bindEvents();
        this.updateUI();
        this.setupPWA();
        this.restoreAutosave();
    }

    // ===================================================
    // INICIALIZAÇÃO
    // ===================================================
    initDOM() {
        this.DOM = {
            spreadsheet: document.getElementById('spreadsheet'),
            formulaField: document.getElementById('formulaField'),
            cellReference: document.getElementById('cell-reference'),
            statusReady: document.getElementById('status-ready'),
            statusCell: document.getElementById('status-cell'),
            statusSelection: document.getElementById('status-selection'),
            statusAgg: document.getElementById('status-agg'),
            zoomLevel: document.getElementById('zoomLevel'),
            sheetTabs: document.getElementById('sheet-tabs'),
            sheetScroll: document.getElementById('sheet-scroll'),
            chartLayer: document.getElementById('chart-layer'),
            workbookName: document.getElementById('workbook-name'),
            fileStatus: document.getElementById('file-status')
        };
    }

    makeEmptySheet(name, rowCount = 100, colCount = 26) {
        const sheet = {
            name,
            data: {},
            styles: {},
            mergedCells: {},
            rowCount, colCount,
            colWidths: {},
            rowHeights: {},
            comments: {},
            validations: {},
            conditionalFormats: [],
            charts: [],
            freeze: { row: false, col: false },
            showGridlines: true
        };
        for (let r = 0; r < rowCount; r++) {
            for (let c = 0; c < colCount; c++) {
                const key = r + ',' + c;
                sheet.data[key] = '';
                sheet.styles[key] = this.getDefaultStyle();
            }
        }
        return sheet;
    }

    initWorkbook() {
        const sheet = this.makeEmptySheet('Planilha1');
        const sampleData = [
            ['Produto', 'Jan', 'Fev', 'Mar', 'Total'],
            ['Produto A', 150, 200, 180, '=SUM(B2:D2)'],
            ['Produto B', 120, 140, 160, '=SUM(B3:D3)'],
            ['Produto C', 90, 110, 130, '=SUM(B4:D4)'],
            ['Total', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)', '=SUM(B5:D5)']
        ];
        sampleData.forEach((row, r) => {
            row.forEach((value, c) => {
                sheet.data[r + ',' + c] = value;
                if (r === 0 || c === 0 || r === sampleData.length - 1) {
                    sheet.styles[r + ',' + c] = { ...this.getDefaultStyle(), bold: true, backgroundColor: r === 0 ? '#107C41' : '#f0f0f0', color: r === 0 ? '#ffffff' : '#000000' };
                }
            });
        });
        this.state.sheets = [sheet];
        this.state.activeSheet = 0;
    }

    getDefaultStyle() {
        return {
            bold: false, italic: false, underline: false, strike: false,
            fontFamily: 'Calibri', fontSize: '12', color: '#000000',
            backgroundColor: '#ffffff', align: 'left', valign: 'middle',
            wrapText: false, numberFormat: 'general'
        };
    }

    get sheet() { return this.state.sheets[this.state.activeSheet]; }

    // ===================================================
    // REFERÊNCIAS DE CÉLULA
    // ===================================================
    getCellKey(row, col) { return `${row},${col}`; }
    getColumnLetter(col) {
        let letter = '';
        col++;
        while (col > 0) {
            col--;
            letter = String.fromCharCode(65 + (col % 26)) + letter;
            col = Math.floor(col / 26);
        }
        return letter;
    }
    getCellReference(row, col) { return this.getColumnLetter(col) + (row + 1); }
    parseCellReference(ref) {
        const match = ref.match(/^\$?([A-Z]+)\$?(\d+)$/i);
        if (!match) return null;
        const col = match[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        const row = parseInt(match[2]) - 1;
        return { row, col };
    }
    colLettersToIndex(letters) {
        return letters.toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    }

    // ===================================================
    // MOTOR DE FÓRMULAS - integração
    // ===================================================
    sheetByName(name) { return this.state.sheets.find(s => s.name === name); }

    getCellRaw(row, col, sheetIndex = this.state.activeSheet) {
        const s = this.state.sheets[sheetIndex];
        if (!s) return '';
        return s.data[this.getCellKey(row, col)] ?? '';
    }

    setCellRaw(row, col, value, sheetIndex = this.state.activeSheet) {
        const s = this.state.sheets[sheetIndex];
        if (!s) return;
        this.ensureCellCapacity(s, row, col);
        s.data[this.getCellKey(row, col)] = value;
        if (!s.styles[this.getCellKey(row, col)]) s.styles[this.getCellKey(row, col)] = this.getDefaultStyle();
    }

    ensureCellCapacity(sheet, row, col) {
        if (row >= sheet.rowCount) sheet.rowCount = row + 1;
        if (col >= sheet.colCount) sheet.colCount = col + 1;
    }

    getDisplayValue(row, col, sheetIndex = this.state.activeSheet) {
        const raw = this.getCellRaw(row, col, sheetIndex);
        if (this.state.showFormulas && typeof raw === 'string' && raw.startsWith('=')) return raw;
        if (typeof raw === 'string' && raw.startsWith('=')) {
            const key = sheetIndex + '!' + row + ',' + col;
            if (this.recalcVisiting.has(key)) return '#CIRCULAR!';
            this.recalcVisiting.add(key);
            try {
                const sheetName = this.state.sheets[sheetIndex].name;
                const evaluator = new FormulaEvaluator({
                    currentSheet: sheetName,
                    sheetExists: (name) => !!this.sheetByName(name),
                    getCellComputed: (sh, r, c) => {
                        const idx = this.state.sheets.findIndex(x => x.name === sh);
                        return this.getDisplayValue(r, c, idx === -1 ? sheetIndex : idx);
                    }
                });
                const result = evaluator.evaluate(raw);
                this.recalcVisiting.delete(key);
                return result;
            } catch (e) {
                this.recalcVisiting.delete(key);
                return '#ERROR!';
            }
        }
        return raw;
    }

    // ===================================================
    // RENDERIZAÇÃO DA PLANILHA
    // ===================================================
    render() {
        const s = this.sheet;
        const table = this.DOM.spreadsheet;
        table.innerHTML = '';
        table.classList.toggle('no-grid', !s.showGridlines);

        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        const cornerTh = document.createElement('th');
        cornerTh.className = 'corner-header';
        headerRow.appendChild(cornerTh);

        // Ajuste: definir a largura da tabela e garantir que os cabeçalhos fiquem fixos
        for (let c = 0; c < s.colCount; c++) {
            const th = document.createElement('th');
            th.textContent = this.getColumnLetter(c);
            th.dataset.col = c;
            const width = s.colWidths[c] || 90;
            th.style.width = width + 'px';
            th.style.minWidth = width + 'px';
            th.style.position = 'sticky';
            th.style.top = '0';
            th.style.zIndex = '5';
            th.addEventListener('click', (e) => this.selectColumn(c, e.shiftKey));
            const handle = document.createElement('span');
            handle.className = 'col-resize-handle';
            handle.addEventListener('mousedown', (e) => this.startColResize(e, c));
            th.appendChild(handle);
            headerRow.appendChild(th);
        }
        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        for (let r = 0; r < s.rowCount; r++) {
            const tr = document.createElement('tr');
            tr.style.height = (s.rowHeights[r] || 28) + 'px';

            const tdRow = document.createElement('td');
            tdRow.className = 'row-header';
            tdRow.textContent = r + 1;
            tdRow.dataset.row = r;
            tdRow.style.position = 'sticky';
            tdRow.style.left = '0';
            tdRow.style.zIndex = '2';
            tdRow.addEventListener('click', (e) => this.selectRow(r, e.shiftKey));
            const rHandle = document.createElement('span');
            rHandle.className = 'row-resize-handle';
            rHandle.addEventListener('mousedown', (e) => this.startRowResize(e, r));
            tdRow.appendChild(rHandle);
            tr.appendChild(tdRow);

            for (let c = 0; c < s.colCount; c++) {
                const td = document.createElement('td');
                td.dataset.row = r;
                td.dataset.col = c;

                const mergeKey = this.getMergeKey(r, c);
                if (mergeKey && mergeKey !== this.getCellKey(r, c)) {
                    td.style.display = 'none';
                    tr.appendChild(td);
                    continue;
                }

                td.addEventListener('mousedown', (e) => this.onCellMouseDown(r, c, e));
                td.addEventListener('mouseenter', () => this.onCellMouseEnter(r, c));
                td.addEventListener('dblclick', () => this.startEditing(r, c));

                this.applyCellStyle(td, r, c);
                this.renderCellContent(td, r, c);

                if (s.mergedCells[this.getCellKey(r, c)]) {
                    const merge = s.mergedCells[this.getCellKey(r, c)];
                    if (merge.startRow === r && merge.startCol === c) {
                        td.colSpan = merge.endCol - merge.startCol + 1;
                        td.rowSpan = merge.endRow - merge.startRow + 1;
                    }
                }

                if (s.comments[this.getCellKey(r, c)]) {
                    td.classList.add('has-comment');
                    td.title = s.comments[this.getCellKey(r, c)];
                }

                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
        table.appendChild(tbody);

        this.applyConditionalFormats();
        this.applySelectionClasses();
        this.renderFillHandle();
        this.renderCharts();
    }

    renderCellContent(td, row, col) {
        const value = this.getDisplayValue(row, col);
        const key = this.getCellKey(row, col);
        const style = this.sheet.styles[key] || this.getDefaultStyle();
        td.textContent = this.formatForDisplay(value, style.numberFormat);
    }

    formatForDisplay(value, numberFormat) {
        if (typeof value === 'boolean') return value ? 'VERDADEIRO' : 'FALSO';
        if (value === null || value === undefined) return '';
        if (typeof value === 'string' && /^#[A-Z/!?0-9]+$/.test(value)) return value;
        const isNum = String(value).trim() !== '' && !isNaN(parseFloat(value)) && isFinite(parseFloat(value));
        if (isNum && numberFormat !== 'general' && numberFormat !== 'text') {
            const num = parseFloat(value);
            switch (numberFormat) {
                case 'number': return num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                case 'currency': return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(num);
                case 'accounting': return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', currencySign: 'accounting' }).format(num);
                case 'percent': return num.toLocaleString('pt-BR', { style: 'percent', minimumFractionDigits: 2 });
                case 'fraction': return this.toFraction(num);
                case 'scientific': return num.toExponential(2);
            }
        }
        return String(value);
    }

    toFraction(value) {
        const frac = Math.abs(value) % 1;
        if (frac === 0) return String(value);
        const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
        const denominator = 64;
        const numerator = Math.round(frac * denominator);
        if (numerator === 0) return String(Math.trunc(value));
        const divisor = gcd(numerator, denominator);
        return `${Math.trunc(value)} ${numerator / divisor}/${denominator / divisor}`;
    }

    applyCellStyle(td, row, col) {
        const key = this.getCellKey(row, col);
        const style = this.sheet.styles[key] || this.getDefaultStyle();
        td.style.fontWeight = style.bold ? 'bold' : 'normal';
        td.style.fontStyle = style.italic ? 'italic' : 'normal';
        let decoration = '';
        if (style.underline) decoration += 'underline ';
        if (style.strike) decoration += 'line-through ';
        td.style.textDecoration = decoration.trim() || 'none';
        td.style.fontFamily = style.fontFamily;
        td.style.fontSize = style.fontSize + 'px';
        td.style.color = style.color;
        td.style.backgroundColor = style.backgroundColor;
        td.style.textAlign = style.align;
        td.style.verticalAlign = style.valign || 'middle';
        td.style.whiteSpace = style.wrapText ? 'normal' : 'nowrap';
        if (style.border) td.style.border = style.border;
    }

    getCellElement(row, col) {
        return this.DOM.spreadsheet.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
    }

    getMergeKey(row, col) {
        for (const [mergeKey, merge] of Object.entries(this.sheet.mergedCells)) {
            if (row >= merge.startRow && row <= merge.endRow && col >= merge.startCol && col <= merge.endCol) {
                return mergeKey;
            }
        }
        return null;
    }

    // ===================================================
    // SELEÇÃO
    // ===================================================
    normSelection() {
        const { start, end } = this.state.selection;
        return {
            r1: Math.min(start.row, end.row), r2: Math.max(start.row, end.row),
            c1: Math.min(start.col, end.col), c2: Math.max(start.col, end.col)
        };
    }

    selectCell(row, col, extend = false) {
        if (this.state.isEditing) this.finishEditing();
        row = Math.max(0, Math.min(row, this.sheet.rowCount - 1));
        col = Math.max(0, Math.min(col, this.sheet.colCount - 1));
        if (extend) {
            this.state.selection.end = { row, col };
        } else {
            this.state.selection = { start: { row, col }, end: { row, col } };
        }
        this.state.selectedCell = { row, col };
        this.DOM.cellReference.textContent = this.getCellReference(row, col);
        this.DOM.formulaField.value = this.getCellRaw(row, col);
        this.applySelectionClasses();
        this.updateStatus();
    }

    onCellMouseDown(row, col, e) {
        if (e.button !== 0) return;
        if (e.shiftKey) {
            this.selectCell(row, col, true);
        } else {
            this.selectCell(row, col, false);
        }
        this.state.isSelecting = true;
    }

    onCellMouseEnter(row, col) {
        if (this.state.isSelecting) {
            this.state.selection.end = { row, col };
            this.applySelectionClasses();
            this.updateStatus();
        }
    }

    selectColumn(col, extend = false) {
        if (extend) {
            this.state.selection.end = { row: this.sheet.rowCount - 1, col };
        } else {
            this.state.selection = { start: { row: 0, col }, end: { row: this.sheet.rowCount - 1, col } };
            this.state.selectedCell = { row: 0, col };
        }
        this.applySelectionClasses();
        this.updateStatus();
    }

    selectRow(row, extend = false) {
        if (extend) {
            this.state.selection.end = { row, col: this.sheet.colCount - 1 };
        } else {
            this.state.selection = { start: { row, col: 0 }, end: { row, col: this.sheet.colCount - 1 } };
            this.state.selectedCell = { row, col: 0 };
        }
        this.applySelectionClasses();
        this.updateStatus();
    }

    applySelectionClasses() {
        this.DOM.spreadsheet.querySelectorAll('td.selected,td.range-selected').forEach(el => el.classList.remove('selected', 'range-selected'));
        const { r1, r2, c1, c2 } = this.normSelection();
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                const cell = this.getCellElement(r, c);
                if (!cell) continue;
                if (r === this.state.selectedCell.row && c === this.state.selectedCell.col) {
                    cell.classList.add('selected');
                } else {
                    cell.classList.add('range-selected');
                }
            }
        }
    }

    renderFillHandle() {
        let handle = document.getElementById('fillHandle');
        if (handle) handle.remove();
        const { r2, c2 } = this.normSelection();
        const cell = this.getCellElement(r2, c2);
        if (!cell) return;
        handle = document.createElement('div');
        handle.id = 'fillHandle';
        handle.className = 'fill-handle';
        cell.style.position = 'relative';
        handle.addEventListener('mousedown', (e) => { e.stopPropagation(); this.state.isFilling = true; });
        cell.appendChild(handle);
    }

    // ===================================================
    // EDIÇÃO
    // ===================================================
    startEditing(row, col, initialChar = null) {
        this.state.isEditing = true;
        this.state.selectedCell = { row, col };
        this.state.selection = { start: { row, col }, end: { row, col } };
        const cell = this.getCellElement(row, col);
        if (cell) cell.classList.add('editing');
        const value = initialChar !== null ? initialChar : this.getCellRaw(row, col);
        this.DOM.formulaField.value = value;
        this.DOM.formulaField.focus();
        if (initialChar !== null) {
            this.DOM.formulaField.setSelectionRange(value.length, value.length);
        } else {
            this.DOM.formulaField.select();
        }
        this.DOM.cellReference.textContent = this.getCellReference(row, col);
    }

    finishEditing(moveDir = null) {
        if (!this.state.isEditing) { if (moveDir) this.moveSelection(moveDir); return; }
        const { row, col } = this.state.selectedCell;
        const value = this.DOM.formulaField.value;
        this.setCellValue(row, col, value);
        this.state.isEditing = false;
        document.querySelectorAll('#spreadsheet td.editing').forEach(el => el.classList.remove('editing'));
        this.render();
        this.selectCell(row, col);
        if (moveDir) this.moveSelection(moveDir);
        this.autosave();
    }

    cancelEditing() {
        if (!this.state.selectedCell) return;
        const { row, col } = this.state.selectedCell;
        this.DOM.formulaField.value = this.getCellRaw(row, col);
        this.state.isEditing = false;
        document.querySelectorAll('#spreadsheet td.editing').forEach(el => el.classList.remove('editing'));
    }

    moveSelection(dir) {
        let { row, col } = this.state.selectedCell;
        if (dir === 'down') row = Math.min(row + 1, this.sheet.rowCount - 1);
        if (dir === 'right') col = Math.min(col + 1, this.sheet.colCount - 1);
        if (dir === 'up') row = Math.max(row - 1, 0);
        if (dir === 'left') col = Math.max(col - 1, 0);
        this.selectCell(row, col);
    }

    // ===================================================
    // VALORES DE CÉLULA + UNDO
    // ===================================================
    setCellValue(row, col, value) {
        const key = this.getCellKey(row, col);
        const sheetIndex = this.state.activeSheet;
        const oldValue = this.getCellRaw(row, col);
        if (oldValue === value) return;
        this.setCellRaw(row, col, value, sheetIndex);
        this.pushUndo({ type: 'cellChange', sheetIndex, changes: [{ row, col, oldValue, newValue: value }] });
        this.updateUI();
    }

    setCellValueSilent(row, col, value, sheetIndex = this.state.activeSheet) {
        this.setCellRaw(row, col, value, sheetIndex);
    }

    pushUndo(action) {
        this.state.undoStack.push(action);
        if (this.state.undoStack.length > this.state.maxUndo) this.state.undoStack.shift();
        this.state.redoStack = [];
    }

    undo() {
        if (!this.state.undoStack.length) return;
        const action = this.state.undoStack.pop();
        this.state.redoStack.push(action);
        this.applyUndoAction(action, true);
        this.render();
        this.updateUI();
    }

    redo() {
        if (!this.state.redoStack.length) return;
        const action = this.state.redoStack.pop();
        this.state.undoStack.push(action);
        this.applyUndoAction(action, false);
        this.render();
        this.updateUI();
    }

    applyUndoAction(action, isUndo) {
        const sheet = this.state.sheets[action.sheetIndex];
        if (!sheet) return;
        if (action.type === 'cellChange') {
            action.changes.forEach(ch => {
                sheet.data[this.getCellKey(ch.row, ch.col)] = isUndo ? ch.oldValue : ch.newValue;
            });
        } else if (action.type === 'styleChange') {
            action.changes.forEach(ch => {
                sheet.styles[this.getCellKey(ch.row, ch.col)] = isUndo ? { ...ch.oldStyle } : { ...ch.newStyle };
            });
        } else if (action.type === 'bulk') {
            action.changes.forEach(ch => {
                const key = this.getCellKey(ch.row, ch.col);
                sheet.data[key] = isUndo ? ch.oldValue : ch.newValue;
                if (ch.oldStyle) sheet.styles[key] = isUndo ? { ...ch.oldStyle } : { ...ch.newStyle };
            });
        } else if (action.type === 'snapshot') {
            const snap = isUndo ? action.before : action.after;
            this.state.sheets[action.sheetIndex] = JSON.parse(JSON.stringify(snap));
        }
        if (this.state.activeSheet === action.sheetIndex) this.selectCell(this.state.selectedCell.row, this.state.selectedCell.col);
    }

    snapshotSheet(sheetIndex = this.state.activeSheet) {
        return JSON.parse(JSON.stringify(this.state.sheets[sheetIndex]));
    }

    // ===================================================
    // ESTILOS (aplicados ao intervalo selecionado)
    // ===================================================
    forEachSelectedCell(cb) {
        const { r1, r2, c1, c2 } = this.normSelection();
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) cb(r, c);
    }

    toggleStyle(prop) {
        const { r1, c1 } = this.normSelection();
        const key0 = this.getCellKey(r1, c1);
        const newVal = !(this.sheet.styles[key0] && this.sheet.styles[key0][prop]);
        this.applyStyleToSelection(prop, newVal);
    }

    applyStyleToSelection(prop, value) {
        const changes = [];
        this.forEachSelectedCell((r, c) => {
            const key = this.getCellKey(r, c);
            if (!this.sheet.styles[key]) this.sheet.styles[key] = this.getDefaultStyle();
            const oldStyle = { ...this.sheet.styles[key] };
            this.sheet.styles[key][prop] = value;
            changes.push({ row: r, col: c, oldStyle, newStyle: { ...this.sheet.styles[key] } });
        });
        this.pushUndo({ type: 'styleChange', sheetIndex: this.state.activeSheet, changes });
        this.render();
        this.selectCell(this.state.selectedCell.row, this.state.selectedCell.col);
        this.autosave();
    }

    applyNumberFormat(format) { this.applyStyleToSelection('numberFormat', format); }

    adjustDecimals(delta) {
        this.forEachSelectedCell((r, c) => {
            const key = this.getCellKey(r, c);
            const style = this.sheet.styles[key];
            if (style.numberFormat === 'general') style.numberFormat = 'number';
        });
        this.render();
    }

    // ===================================================
    // MESCLAR
    // ===================================================
    mergeCells() {
        const { r1, r2, c1, c2 } = this.normSelection();
        if (r1 === r2 && c1 === c2) return;
        const key = this.getCellKey(r1, c1);
        this.sheet.mergedCells[key] = { startRow: r1, startCol: c1, endRow: r2, endCol: c2 };
        this.render();
        this.selectCell(r1, c1);
        this.autosave();
    }

    unmergeCells() {
        const { r1, c1 } = this.normSelection();
        const removeKeys = [];
        for (const [mk, merge] of Object.entries(this.sheet.mergedCells)) {
            if (r1 >= merge.startRow && r1 <= merge.endRow && c1 >= merge.startCol && c1 <= merge.endCol) removeKeys.push(mk);
        }
        removeKeys.forEach(k => delete this.sheet.mergedCells[k]);
        this.render();
        this.selectCell(this.state.selectedCell.row, this.state.selectedCell.col);
        this.autosave();
    }

    // ===================================================
    // AJUSTE AUTOMÁTICO DE REFERÊNCIAS AO INSERIR/EXCLUIR
    // ===================================================
    remapFormulaRefs(formula, transform) {
        if (typeof formula !== 'string' || !formula.startsWith('=')) return formula;
        const s = formula.slice(1);
        let out = '=';
        let i = 0;
        while (i < s.length) {
            const ch = s[i];
            if (ch === '"') {
                let j = i + 1;
                while (j < s.length && s[j] !== '"') j++;
                out += s.slice(i, j + 1);
                i = j + 1;
                continue;
            }
            const remaining = s.slice(i);
            const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/.exec(remaining);
            const prevChar = s[i - 1] || '';
            if (m && !/[A-Za-z0-9_]/.test(prevChar)) {
                const [full, dollarCol, colLetters, dollarRow, rowNum] = m;
                const col = this.colLettersToIndex(colLetters);
                const row = parseInt(rowNum, 10) - 1;
                const mapped = transform(row, col);
                if (!mapped) out += '#REF!';
                else out += dollarCol + this.getColumnLetter(mapped.col) + dollarRow + (mapped.row + 1);
                i += full.length;
                continue;
            }
            out += ch;
            i++;
        }
        return out;
    }

    shiftFormulaRefsInSheet(sheetIndex, transform) {
        const s = this.state.sheets[sheetIndex];
        if (!s) return;
        for (let r = 0; r < s.rowCount; r++) {
            for (let c = 0; c < s.colCount; c++) {
                const key = this.getCellKey(r, c);
                const raw = s.data[key];
                if (typeof raw === 'string' && raw.startsWith('=')) {
                    s.data[key] = this.remapFormulaRefs(raw, transform);
                }
            }
        }
    }

    // ===================================================
    // INSERIR/EXCLUIR LINHAS E COLUNAS
    // ===================================================
    insertRow() {
        const row = this.state.selectedCell.row;
        const s = this.sheet;
        s.rowCount++;
        for (let r = s.rowCount - 1; r > row; r--) {
            for (let c = 0; c < s.colCount; c++) {
                const key = this.getCellKey(r, c), prevKey = this.getCellKey(r - 1, c);
                s.data[key] = s.data[prevKey] || '';
                s.styles[key] = { ...(s.styles[prevKey] || this.getDefaultStyle()) };
            }
        }
        for (let c = 0; c < s.colCount; c++) {
            const key = this.getCellKey(row, c);
            s.data[key] = '';
            s.styles[key] = this.getDefaultStyle();
        }
        this.shiftFormulaRefsInSheet(this.state.activeSheet, (r, c) => ({ row: r >= row ? r + 1 : r, col: c }));
        this.render();
        this.selectCell(row, 0);
        this.autosave();
    }

    deleteRow() {
        const s = this.sheet;
        const { r1, r2 } = this.normSelection();
        if (s.rowCount <= (r2 - r1 + 1)) return;
        const count = r2 - r1 + 1;
        for (let r = r1; r < s.rowCount - count; r++) {
            for (let c = 0; c < s.colCount; c++) {
                const key = this.getCellKey(r, c), nextKey = this.getCellKey(r + count, c);
                s.data[key] = s.data[nextKey] || '';
                s.styles[key] = { ...(s.styles[nextKey] || this.getDefaultStyle()) };
            }
        }
        s.rowCount -= count;
        this.shiftFormulaRefsInSheet(this.state.activeSheet, (r, c) => {
            if (r >= r1 && r <= r2) return null;
            return { row: r > r2 ? r - count : r, col: c };
        });
        this.render();
        this.selectCell(Math.min(r1, s.rowCount - 1), 0);
        this.autosave();
    }

    insertColumn() {
        const col = this.state.selectedCell.col;
        const s = this.sheet;
        s.colCount++;
        for (let c = s.colCount - 1; c > col; c--) {
            for (let r = 0; r < s.rowCount; r++) {
                const key = this.getCellKey(r, c), prevKey = this.getCellKey(r, c - 1);
                s.data[key] = s.data[prevKey] || '';
                s.styles[key] = { ...(s.styles[prevKey] || this.getDefaultStyle()) };
            }
        }
        for (let r = 0; r < s.rowCount; r++) {
            const key = this.getCellKey(r, col);
            s.data[key] = '';
            s.styles[key] = this.getDefaultStyle();
        }
        this.shiftFormulaRefsInSheet(this.state.activeSheet, (r, c) => ({ row: r, col: c >= col ? c + 1 : c }));
        this.render();
        this.selectCell(0, col);
        this.autosave();
    }

    deleteColumn() {
        const s = this.sheet;
        const { c1, c2 } = this.normSelection();
        const count = c2 - c1 + 1;
        if (s.colCount <= count) return;
        for (let c = c1; c < s.colCount - count; c++) {
            for (let r = 0; r < s.rowCount; r++) {
                const key = this.getCellKey(r, c), nextKey = this.getCellKey(r, c + count);
                s.data[key] = s.data[nextKey] || '';
                s.styles[key] = { ...(s.styles[nextKey] || this.getDefaultStyle()) };
            }
        }
        s.colCount -= count;
        this.shiftFormulaRefsInSheet(this.state.activeSheet, (r, c) => {
            if (c >= c1 && c <= c2) return null;
            return { row: r, col: c > c2 ? c - count : c };
        });
        this.render();
        this.selectCell(0, Math.min(c1, s.colCount - 1));
        this.autosave();
    }

    clearCells() {
        const changes = [];
        this.forEachSelectedCell((r, c) => {
            const oldValue = this.getCellRaw(r, c);
            if (oldValue !== '') {
                changes.push({ row: r, col: c, oldValue, newValue: '' });
                this.setCellRaw(r, c, '');
            }
        });
        if (changes.length) this.pushUndo({ type: 'cellChange', sheetIndex: this.state.activeSheet, changes });
        this.render();
        this.selectCell(this.state.selectedCell.row, this.state.selectedCell.col);
        this.autosave();
    }

    // ===================================================
    // REDIMENSIONAR COLUNAS/LINHAS
    // ===================================================
    startColResize(e, col) {
        e.stopPropagation(); e.preventDefault();
        const startX = e.clientX;
        const startWidth = this.sheet.colWidths[col] || 90;
        const onMove = (ev) => {
            const w = Math.max(30, startWidth + (ev.clientX - startX));
            this.sheet.colWidths[col] = w;
            const th = this.DOM.spreadsheet.querySelector(`th[data-col="${col}"]`);
            if (th) { th.style.width = w + 'px'; th.style.minWidth = w + 'px'; }
        };
        const onUp = () => {
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            this.render();
            this.autosave();
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
    }

    startRowResize(e, row) {
        e.stopPropagation(); e.preventDefault();
        const startY = e.clientY;
        const startHeight = this.sheet.rowHeights[row] || 28;
        const onMove = (ev) => {
            const h = Math.max(16, startHeight + (ev.clientY - startY));
            this.sheet.rowHeights[row] = h;
        };
        const onUp = () => {
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            this.render();
            this.autosave();
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
    }

    // ===================================================
    // ÁREA DE TRANSFERÊNCIA (com ajuste de referências relativas)
    // ===================================================
    copy(cut = false) {
        const { r1, r2, c1, c2 } = this.normSelection();
        const cells = [];
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                cells.push({ row: r, col: c, value: this.getCellRaw(r, c), style: { ...(this.sheet.styles[this.getCellKey(r, c)] || this.getDefaultStyle()) } });
            }
        }
        this.state.clipboard = { mode: cut ? 'cut' : 'copy', sheetIndex: this.state.activeSheet, cells, anchorRow: r1, anchorCol: c1 };
        this.markClipboardVisual();
        this.DOM.statusReady.textContent = cut ? 'Recortado' : 'Copiado';
    }

    cut() { this.copy(true); }

    markClipboardVisual() {
        this.DOM.spreadsheet.querySelectorAll('.marching-ants').forEach(el => el.classList.remove('marching-ants'));
        if (!this.state.clipboard) return;
        this.state.clipboard.cells.forEach(cell => {
            const el = this.getCellElement(cell.row, cell.col);
            if (el) el.classList.add('marching-ants');
        });
    }

    adjustFormulaRefs(formula, dRow, dCol) {
        if (typeof formula !== 'string' || !formula.startsWith('=')) return formula;
        const s = formula.slice(1);
        let out = '=';
        let i = 0;
        while (i < s.length) {
            const ch = s[i];
            if (ch === '"') {
                let j = i + 1;
                while (j < s.length && s[j] !== '"') j++;
                out += s.slice(i, j + 1);
                i = j + 1;
                continue;
            }
            const remaining = s.slice(i);
            const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/.exec(remaining);
            const prevChar = s[i - 1] || '';
            if (m && !/[A-Za-z0-9_]/.test(prevChar)) {
                const [full, dollarCol, colLetters, dollarRow, rowNum] = m;
                let newCol = this.colLettersToIndex(colLetters);
                let newRow = parseInt(rowNum, 10) - 1;
                if (!dollarCol) newCol += dCol;
                if (!dollarRow) newRow += dRow;
                if (newCol < 0 || newRow < 0) {
                    out += '#REF!';
                } else {
                    out += dollarCol + this.getColumnLetter(newCol) + dollarRow + (newRow + 1);
                }
                i += full.length;
                continue;
            }
            out += ch;
            i++;
        }
        return out;
    }

    paste() {
        if (!this.state.clipboard) return;
        const clip = this.state.clipboard;
        const { row: destRow, col: destCol } = this.state.selectedCell;
        const dRow = destRow - clip.anchorRow;
        const dCol = destCol - clip.anchorCol;
        const changes = [];
        clip.cells.forEach(cell => {
            const newRow = cell.row + dRow, newCol = cell.col + dCol;
            if (newRow < 0 || newCol < 0) return;
            this.ensureCellCapacity(this.sheet, newRow, newCol);
            const oldValue = this.getCellRaw(newRow, newCol);
            const oldStyle = { ...(this.sheet.styles[this.getCellKey(newRow, newCol)] || this.getDefaultStyle()) };
            const newValue = this.adjustFormulaRefs(cell.value, dRow, dCol);
            this.setCellRaw(newRow, newCol, newValue);
            this.sheet.styles[this.getCellKey(newRow, newCol)] = { ...cell.style };
            changes.push({ row: newRow, col: newCol, oldValue, newValue, oldStyle, newStyle: { ...cell.style } });
            if (clip.mode === 'cut') {
                this.setCellRaw(cell.row, cell.col, '', clip.sheetIndex);
            }
        });
        if (clip.mode === 'cut') {
            clip.cells.forEach(cell => changes.push({ row: cell.row, col: cell.col, oldValue: cell.value, newValue: '' }));
            this.state.clipboard = null;
        }
        this.pushUndo({ type: 'bulk', sheetIndex: this.state.activeSheet, changes });
        this.render();
        this.selectCell(destRow, destCol);
        this.autosave();
    }

    // ===================================================
    // ALÇA DE PREENCHIMENTO (arrastar para preencher) - CORRIGIDO
    // ===================================================
    finishFill(targetRow, targetCol) {
        const { r1, r2, c1, c2 } = this.normSelection();
        const source = [];
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                source.push({ row: r, col: c, value: this.getCellRaw(r, c) });
            }
        }

        const changes = [];
        let isVertical = false;

        // Preenchimento vertical (para baixo)
        if (targetRow > r2) {
            isVertical = true;
            // Coletar valores numéricos da primeira coluna do intervalo fonte
            const nums = [];
            for (let r = r1; r <= r2; r++) {
                const val = this.getCellRaw(r, c1);
                const num = parseFloat(val);
                if (!isNaN(num) && val !== '') {
                    nums.push(num);
                }
            }
            
            // Calcular o passo (step) - diferença entre o último e o primeiro valor
            let step = 1;
            if (nums.length >= 2) {
                // Usar a diferença entre os dois últimos valores para o passo
                step = nums[nums.length - 1] - nums[nums.length - 2];
                // Se o passo for 0, usar a diferença entre o primeiro e o último
                if (step === 0 && nums.length >= 2) {
                    step = nums[nums.length - 1] - nums[0];
                }
            }
            
            // Se só tem um número, usar 1 como passo
            if (nums.length === 1) {
                step = 1;
            }

            let rowIdx = 0;
            const srcHeight = r2 - r1 + 1;
            
            for (let r = r2 + 1; r <= targetRow; r++) {
                const srcRow = r1 + (rowIdx % srcHeight);
                for (let c = c1; c <= c2; c++) {
                    const srcCell = source[(rowIdx % srcHeight) * (c2 - c1 + 1) + (c - c1)];
                    let val = srcCell.value;
                    
                    if (typeof val === 'string' && val.startsWith('=')) {
                        // Ajustar fórmulas
                        val = this.adjustFormulaRefs(val, r - srcCell.row, 0);
                    } else if (c === c1 && nums.length >= 1) {
                        // Preencher com incremento baseado no passo
                        const baseVal = nums[nums.length - 1] || nums[0];
                        const increment = (r - r2) * step;
                        val = baseVal + increment;
                    } else {
                        // Para outras colunas, copiar o valor
                        val = srcCell.value;
                    }
                    
                    const oldValue = this.getCellRaw(r, c);
                    this.setCellRaw(r, c, String(val));
                    changes.push({ row: r, col: c, oldValue, newValue: String(val) });
                }
                rowIdx++;
            }
            this.state.selection.end = { row: targetRow, col: c2 };
        } 
        // Preenchimento horizontal (para a direita)
        else if (targetCol > c2) {
            // Coletar valores numéricos da primeira linha do intervalo fonte
            const nums = [];
            for (let c = c1; c <= c2; c++) {
                const val = this.getCellRaw(r1, c);
                const num = parseFloat(val);
                if (!isNaN(num) && val !== '') {
                    nums.push(num);
                }
            }
            
            let step = 1;
            if (nums.length >= 2) {
                step = nums[nums.length - 1] - nums[nums.length - 2];
                if (step === 0 && nums.length >= 2) {
                    step = nums[nums.length - 1] - nums[0];
                }
            }
            if (nums.length === 1) {
                step = 1;
            }

            let colIdx = 0;
            const srcWidth = c2 - c1 + 1;
            
            for (let c = c2 + 1; c <= targetCol; c++) {
                const srcCol = c1 + (colIdx % srcWidth);
                for (let r = r1; r <= r2; r++) {
                    const srcCell = source[(r - r1) * srcWidth + (colIdx % srcWidth)];
                    let val = srcCell.value;
                    
                    if (typeof val === 'string' && val.startsWith('=')) {
                        val = this.adjustFormulaRefs(val, 0, c - srcCell.col);
                    } else if (r === r1 && nums.length >= 1) {
                        const baseVal = nums[nums.length - 1] || nums[0];
                        const increment = (c - c2) * step;
                        val = baseVal + increment;
                    } else {
                        val = srcCell.value;
                    }
                    
                    const oldValue = this.getCellRaw(r, c);
                    this.setCellRaw(r, c, String(val));
                    changes.push({ row: r, col: c, oldValue, newValue: String(val) });
                }
                colIdx++;
            }
            this.state.selection.end = { row: r2, col: targetCol };
        }

        if (changes.length) {
            this.pushUndo({ type: 'cellChange', sheetIndex: this.state.activeSheet, changes });
        }
        this.render();
        this.applySelectionClasses();
        this.autosave();
    }

    // ===================================================
    // AUTOSSOMA (insere fórmula real)
    // ===================================================
    autoSum() {
        const { row, col } = this.state.selectedCell;
        let r = row - 1;
        while (r >= 0 && this.getDisplayValue(r, col) !== '' && !isNaN(parseFloat(this.getDisplayValue(r, col)))) r--;
        const startRow = r + 1;
        if (startRow >= row) { this.DOM.statusReady.textContent = 'Nada para somar'; return; }
        const formula = `=SUM(${this.getCellReference(startRow, col)}:${this.getCellReference(row - 1, col)})`;
        this.startEditing(row, col, formula);
        this.finishEditing();
    }

    // ===================================================
    // ORDENAÇÃO
    // ===================================================
    sortSelection(ascending = true) {
        const { r1, r2, c1, c2 } = this.normSelection();
        if (r2 <= r1) { this.DOM.statusReady.textContent = 'Selecione mais de uma linha para ordenar'; return; }
        const sortCol = this.state.selectedCell.col >= c1 && this.state.selectedCell.col <= c2 ? this.state.selectedCell.col : c1;
        const rows = [];
        for (let r = r1; r <= r2; r++) {
            const rowVals = [];
            for (let c = c1; c <= c2; c++) rowVals.push(this.getCellRaw(r, c));
            rows.push(rowVals);
        }
        const sortIdx = sortCol - c1;
        rows.sort((a, b) => {
            const av = a[sortIdx], bv = b[sortIdx];
            const an = parseFloat(av), bn = parseFloat(bv);
            let cmp;
            if (!isNaN(an) && !isNaN(bn) && av !== '' && bv !== '') cmp = an - bn;
            else cmp = String(av).localeCompare(String(bv), 'pt-BR');
            return ascending ? cmp : -cmp;
        });
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                this.setCellRaw(r, c, rows[r - r1][c - c1]);
            }
        }
        this.render();
        this.autosave();
    }

    // ===================================================
    // LOCALIZAR E SUBSTITUIR
    // ===================================================
    findReplace(term, replacement, replaceAll) {
        const s = this.sheet;
        let found = 0;
        for (let r = 0; r < s.rowCount; r++) {
            for (let c = 0; c < s.colCount; c++) {
                const val = String(this.getCellRaw(r, c));
                if (val.toLowerCase().includes(term.toLowerCase())) {
                    found++;
                    if (replacement !== null) {
                        const newVal = val.split(new RegExp(term, 'gi')).join(replacement);
                        this.setCellRaw(r, c, newVal);
                        if (!replaceAll) { this.render(); this.selectCell(r, c); return found; }
                    } else {
                        this.selectCell(r, c);
                        this.render();
                        return found;
                    }
                }
            }
        }
        this.render();
        this.autosave();
        return found;
    }

    // ===================================================
    // FORMATAÇÃO CONDICIONAL
    // ===================================================
    addConditionalFormat(op, value, color) {
        const { r1, r2, c1, c2 } = this.normSelection();
        this.sheet.conditionalFormats.push({ r1, r2, c1, c2, op, value: parseFloat(value), color });
        this.render();
        this.autosave();
    }

    applyConditionalFormats() {
        (this.sheet.conditionalFormats || []).forEach(rule => {
            for (let r = rule.r1; r <= rule.r2; r++) {
                for (let c = rule.c1; c <= rule.c2; c++) {
                    const val = parseFloat(this.getDisplayValue(r, c));
                    if (isNaN(val)) continue;
                    let match = false;
                    switch (rule.op) {
                        case '>': match = val > rule.value; break;
                        case '<': match = val < rule.value; break;
                        case '=': match = val === rule.value; break;
                        case '>=': match = val >= rule.value; break;
                        case '<=': match = val <= rule.value; break;
                    }
                    if (match) {
                        const el = this.getCellElement(r, c);
                        if (el) el.style.backgroundColor = rule.color;
                    }
                }
            }
        });
    }

    // ===================================================
    // PLANILHAS (ABAS)
    // ===================================================
    addSheet() {
        const name = this.uniqueSheetName('Planilha');
        this.state.sheets.push(this.makeEmptySheet(name));
        this.switchSheet(this.state.sheets.length - 1);
        this.autosave();
    }

    uniqueSheetName(base) {
        let n = this.state.sheets.length + 1;
        let name = `${base}${n}`;
        while (this.sheetByName(name)) { n++; name = `${base}${n}`; }
        return name;
    }

    switchSheet(index) {
        if (this.state.isEditing) this.finishEditing();
        this.state.activeSheet = index;
        this.state.selection = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
        this.state.selectedCell = { row: 0, col: 0 };
        this.render();
        this.renderSheetTabs();
        this.updateUI();
        this.DOM.statusReady.textContent = `Planilha ${this.state.sheets[index].name} ativa`;
    }

    deleteSheet(index) {
        if (this.state.sheets.length <= 1) { alert('A pasta de trabalho deve ter ao menos uma planilha.'); return; }
        if (!confirm(`Excluir "${this.state.sheets[index].name}"?`)) return;
        this.state.sheets.splice(index, 1);
        if (this.state.activeSheet >= this.state.sheets.length) this.state.activeSheet = this.state.sheets.length - 1;
        this.render();
        this.renderSheetTabs();
        this.autosave();
    }

    renameSheet(index, name) {
        name = (name || '').trim();
        if (!name || this.sheetByName(name)) return;
        this.state.sheets[index].name = name;
        this.renderSheetTabs();
        this.autosave();
    }

    renderSheetTabs() {
        const container = this.DOM.sheetTabs;
        container.innerHTML = '';
        this.state.sheets.forEach((s, index) => {
            const tab = document.createElement('span');
            tab.className = 'sheet-tab' + (index === this.state.activeSheet ? ' active' : '');
            tab.textContent = s.name;
            tab.dataset.sheet = index;
            tab.addEventListener('click', () => this.switchSheet(index));
            tab.addEventListener('dblclick', () => {
                const newName = prompt('Novo nome da planilha:', s.name);
                if (newName) this.renameSheet(index, newName);
            });
            tab.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                if (confirm(`Excluir a planilha "${s.name}"?`)) this.deleteSheet(index);
            });
            container.appendChild(tab);
        });
        const addBtn = document.createElement('button');
        addBtn.id = 'btnAddSheet';
        addBtn.className = 'sheet-add-btn';
        addBtn.innerHTML = '<i class="fas fa-plus"></i>';
        addBtn.title = 'Nova planilha';
        addBtn.addEventListener('click', () => this.addSheet());
        container.appendChild(addBtn);
    }

    // ===================================================
    // GRÁFICOS (Chart.js)
    // ===================================================
    createChart(type) {
        const { r1, r2, c1, c2 } = this.normSelection();
        if (r1 === r2 && c1 === c2) { alert('Selecione um intervalo de células com dados para criar o gráfico.'); return; }
        const labels = [];
        const datasets = [];
        const hasHeaderRow = isNaN(parseFloat(this.getDisplayValue(r1, c1 + 1 <= c2 ? c1 + 1 : c1)));
        let dataStartRow = r1;
        let labelCol = c1;
        // Assume first column = labels, first row = series names (if text)
        for (let r = r1; r <= r2; r++) labels.push(String(this.getDisplayValue(r, labelCol)));
        for (let c = c1 + 1; c <= c2; c++) {
            const seriesName = String(this.getDisplayValue(r1, c)) || this.getColumnLetter(c);
            const values = [];
            for (let r = r1; r <= r2; r++) {
                const v = parseFloat(this.getDisplayValue(r, c));
                values.push(isNaN(v) ? 0 : v);
            }
            datasets.push({ label: seriesName, data: values });
        }
        if (!datasets.length) {
            // single column of numbers, use row index as label
            const values = [];
            for (let r = r1; r <= r2; r++) values.push(parseFloat(this.getDisplayValue(r, c1)) || 0);
            datasets.push({ label: 'Série 1', data: values });
        }
        const chart = {
            id: 'chart_' + Date.now(),
            type, labels, datasets,
            x: 60, y: 60, w: 420, h: 280,
            title: `Gráfico ${this.getColumnLetter(c1)}${r1 + 1}:${this.getColumnLetter(c2)}${r2 + 1}`
        };
        this.sheet.charts.push(chart);
        this.renderCharts();
        this.autosave();
    }

    renderCharts() {
        const layer = this.DOM.chartLayer;
        layer.innerHTML = '';
        (this.sheet.charts || []).forEach(chart => {
            const box = document.createElement('div');
            box.className = 'chart-box';
            box.style.left = chart.x + 'px';
            box.style.top = chart.y + 'px';
            box.style.width = chart.w + 'px';
            box.style.height = chart.h + 'px';

            const header = document.createElement('div');
            header.className = 'chart-header';
            header.innerHTML = `<span>${chart.title}</span>`;
            const closeBtn = document.createElement('i');
            closeBtn.className = 'fas fa-times';
            closeBtn.addEventListener('click', () => {
                this.sheet.charts = this.sheet.charts.filter(c => c.id !== chart.id);
                this.renderCharts();
                this.autosave();
            });
            header.appendChild(closeBtn);
            box.appendChild(header);

            const canvasWrap = document.createElement('div');
            canvasWrap.className = 'chart-canvas-wrap';
            const canvas = document.createElement('canvas');
            canvasWrap.appendChild(canvas);
            box.appendChild(canvasWrap);

            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'chart-resize-handle';
            box.appendChild(resizeHandle);

            this.makeDraggable(box, header, chart);
            this.makeResizable(box, resizeHandle, chart);

            layer.appendChild(box);

            if (typeof Chart !== 'undefined') {
                const palette = ['#107C41', '#2b7de9', '#e8a33d', '#d1495b', '#7b61ff', '#00b8a9'];
                const chartType = ({ area: 'line', doughnut: 'doughnut' }[chart.type]) || chart.type;
                const cfg = {
                    type: chartType === 'bar' ? 'bar' : chartType,
                    data: {
                        labels: chart.labels,
                        datasets: chart.datasets.map((ds, i) => ({
                            label: ds.label,
                            data: ds.data,
                            backgroundColor: chart.type === 'pie' || chart.type === 'doughnut'
                                ? chart.labels.map((_, j) => palette[j % palette.length])
                                : palette[i % palette.length] + (chart.type === 'area' ? '55' : ''),
                            borderColor: palette[i % palette.length],
                            fill: chart.type === 'area',
                            borderWidth: 2
                        }))
                    },
                    options: {
                        responsive: true, maintainAspectRatio: false,
                        plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 10 } } } },
                        scales: (chart.type === 'pie' || chart.type === 'doughnut' || chart.type === 'radar') ? {} : { y: { beginAtZero: true } }
                    }
                };
                new Chart(canvas.getContext('2d'), cfg);
            } else {
                canvasWrap.textContent = 'Chart.js indisponível (verifique a conexão).';
            }
        });
    }

    makeDraggable(box, handle, chart) {
        handle.addEventListener('mousedown', (e) => {
            e.stopPropagation();
            const startX = e.clientX, startY = e.clientY;
            const startLeft = chart.x, startTop = chart.y;
            const onMove = (ev) => {
                chart.x = startLeft + (ev.clientX - startX);
                chart.y = startTop + (ev.clientY - startY);
                box.style.left = chart.x + 'px';
                box.style.top = chart.y + 'px';
            };
            const onUp = () => {
                document.removeEventListener('mousemove', onMove);
                document.removeEventListener('mouseup', onUp);
                this.autosave();
            };
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
        });
    }

    makeResizable(box, handle, chart) {
        handle.addEventListener('mousedown', (e) => {
            e.stopPropagation();
            const startX = e.clientX, startY = e.clientY;
            const startW = chart.w, startH = chart.h;
            const onMove = (ev) => {
                chart.w = Math.max(200, startW + (ev.clientX - startX));
                chart.h = Math.max(150, startH + (ev.clientY - startY));
                box.style.width = chart.w + 'px';
                box.style.height = chart.h + 'px';
            };
            const onUp = () => {
                document.removeEventListener('mousemove', onMove);
                document.removeEventListener('mouseup', onUp);
                this.renderCharts();
                this.autosave();
            };
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
        });
    }

    // ===================================================
    // IMPORTAR / EXPORTAR
    // ===================================================
    exportCSV() {
        const s = this.sheet;
        let csv = '';
        for (let r = 0; r < s.rowCount; r++) {
            const row = [];
            for (let c = 0; c < s.colCount; c++) {
                const value = this.getDisplayValue(r, c);
                row.push(`"${String(value).replace(/"/g, '""')}"`);
            }
            csv += row.join(',') + '\r\n';
        }
        this.downloadBlob(new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' }), `${s.name}.csv`);
        this.DOM.statusReady.textContent = 'CSV exportado!';
    }

    importCSV() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv,.xlsx,.xls,.json';
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const ext = file.name.split('.').pop().toLowerCase();
            if (ext === 'xlsx' || ext === 'xls') { this.importXLSXFile(file); return; }
            if (ext === 'json') { this.importJSONFile(file); return; }
            const reader = new FileReader();
            reader.onload = (event) => {
                const csv = event.target.result;
                const rows = this.parseCSV(csv);
                const sheet = this.makeEmptySheet(this.uniqueSheetName('Importado'), Math.max(100, rows.length + 5), 26);
                rows.forEach((row, r) => row.forEach((val, c) => { sheet.data[r + ',' + c] = val; }));
                this.state.sheets.push(sheet);
                this.switchSheet(this.state.sheets.length - 1);
                this.DOM.statusReady.textContent = 'CSV importado!';
                this.autosave();
            };
            reader.readAsText(file);
        };
        input.click();
    }

    parseCSV(text) {
        const rows = [];
        let row = [], field = '', inQuotes = false;
        for (let i = 0; i < text.length; i++) {
            const c = text[i];
            if (inQuotes) {
                if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false; }
                else field += c;
            } else {
                if (c === '"') inQuotes = true;
                else if (c === ',') { row.push(field); field = ''; }
                else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
                else if (c === '\r') { /* skip */ }
                else field += c;
            }
        }
        if (field.length || row.length) { row.push(field); rows.push(row); }
        return rows.filter(r => r.some(f => f !== ''));
    }

    importXLSXFile(file) {
        if (typeof XLSX === 'undefined') { alert('Biblioteca XLSX indisponível (verifique a conexão com a internet).'); return; }
        const reader = new FileReader();
        reader.onload = (e) => {
            const wb = XLSX.read(e.target.result, { type: 'binary', cellFormula: false });
            wb.SheetNames.forEach(name => {
                const ws = wb.Sheets[name];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
                const sheet = this.makeEmptySheet(this.uniqueSheetName(name), Math.max(100, json.length + 5), 26);
                json.forEach((row, r) => row.forEach((val, c) => { sheet.data[r + ',' + c] = val; }));
                this.state.sheets.push(sheet);
            });
            this.switchSheet(this.state.sheets.length - 1);
            this.DOM.statusReady.textContent = 'Pasta de trabalho .xlsx importada!';
            this.autosave();
        };
        reader.readAsBinaryString(file);
    }

    importJSONFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = JSON.parse(e.target.result);
                if (wb.sheets) {
                    this.state.sheets = wb.sheets;
                    this.state.activeSheet = 0;
                    this.render();
                    this.renderSheetTabs();
                    this.DOM.statusReady.textContent = 'Pasta de trabalho carregada!';
                    this.autosave();
                }
            } catch (err) { alert('Arquivo JSON inválido.'); }
        };
        reader.readAsText(file);
    }

    exportPDF() {
        window.print();
    }

    saveWorkbookXLSX() {
        if (typeof XLSX === 'undefined') { alert('Biblioteca XLSX indisponível (verifique a conexão). Exportando como CSV.'); this.exportCSV(); return; }
        const wb = XLSX.utils.book_new();
        this.state.sheets.forEach(s => {
            const aoa = [];
            for (let r = 0; r < s.rowCount; r++) {
                const row = [];
                let hasContent = false;
                for (let c = 0; c < s.colCount; c++) {
                    const v = this.getDisplayValue(r, c, this.state.sheets.indexOf(s));
                    if (v !== '') hasContent = true;
                    row.push(v);
                }
                if (hasContent || r < aoa.length + 1) aoa.push(row);
            }
            const ws = XLSX.utils.aoa_to_sheet(aoa);
            XLSX.utils.book_append_sheet(wb, ws, s.name.substring(0, 31));
        });
        XLSX.writeFile(wb, `${(this.DOM.workbookName.textContent || 'Pasta1').split(' - ')[0]}.xlsx`);
        this.DOM.statusReady.textContent = 'Pasta de trabalho salva (.xlsx)!';
        this.markSaved();
    }

    saveWorkbookJSON() {
        const data = { sheets: this.state.sheets, activeSheet: this.state.activeSheet };
        this.downloadBlob(new Blob([JSON.stringify(data)], { type: 'application/json' }), `pasta_de_trabalho.json`);
        this.markSaved();
    }

    downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = filename; a.click();
        URL.revokeObjectURL(url);
    }

    markSaved() {
        this.DOM.fileStatus.textContent = '- Salvo';
    }
    markUnsaved() {
        this.DOM.fileStatus.textContent = '- Não salvo';
    }

    // ===================================================
    // AUTOSAVE (localStorage)
    // ===================================================
    autosave() {
        this.markUnsaved();
        clearTimeout(this._autosaveTimer);
        this._autosaveTimer = setTimeout(() => {
            try {
                localStorage.setItem('excelSimuladorPro.autosave', JSON.stringify({
                    sheets: this.state.sheets, activeSheet: this.state.activeSheet, savedAt: Date.now()
                }));
                this.markSaved();
            } catch (e) { /* quota exceeded etc */ }
        }, 800);
    }

    restoreAutosave() {
        try {
            const raw = localStorage.getItem('excelSimuladorPro.autosave');
            if (!raw) return;
            const parsed = JSON.parse(raw);
            if (!parsed.sheets || !parsed.sheets.length) return;
            const when = new Date(parsed.savedAt).toLocaleString('pt-BR');
            if (confirm(`Foi encontrado um trabalho salvo automaticamente (${when}). Deseja restaurá-lo?`)) {
                this.state.sheets = parsed.sheets;
                this.state.activeSheet = Math.min(parsed.activeSheet || 0, parsed.sheets.length - 1);
                this.render();
                this.renderSheetTabs();
                this.updateUI();
            }
        } catch (e) { /* ignore */ }
    }

    // ===================================================
    // MODELOS
    // ===================================================
    openTemplatesModal() {
        const names = (typeof ExcelDatabase !== 'undefined') ? ExcelDatabase.listTemplates() : [];
        let html = '<div class="template-list"><div class="template-item" data-tpl="__blank__"><i class="fas fa-file"></i><span>Pasta em branco</span></div>';
        names.forEach(n => { html += `<div class="template-item" data-tpl="${n}"><i class="fas fa-layer-group"></i><span>${n}</span></div>`; });
        html += '</div>';
        this.openGenericModal('Escolher Modelo', html, () => {
            const selected = document.querySelector('.template-item.selected');
            const tpl = selected ? selected.dataset.tpl : '__blank__';
            this.applyTemplate(tpl);
        });
        document.querySelectorAll('.template-item').forEach(item => {
            item.addEventListener('click', () => {
                document.querySelectorAll('.template-item').forEach(i => i.classList.remove('selected'));
                item.classList.add('selected');
            });
        });
    }

    applyTemplate(name) {
        if (name === '__blank__' || typeof ExcelDatabase === 'undefined') {
            this.state.sheets.push(this.makeEmptySheet(this.uniqueSheetName('Planilha')));
        } else {
            const tpl = ExcelDatabase.getTemplate(name);
            const sheet = this.makeEmptySheet(this.uniqueSheetName(name.replace(/\s/g, '')));
            if (tpl) {
                tpl.data.forEach((row, r) => {
                    row.forEach((val, c) => {
                        sheet.data[r + ',' + c] = val;
                        if (r === 0) sheet.styles[r + ',' + c] = { ...this.getDefaultStyle(), bold: true, backgroundColor: '#107C41', color: '#fff' };
                    });
                });
            }
            this.state.sheets.push(sheet);
        }
        this.switchSheet(this.state.sheets.length - 1);
        this.autosave();
    }

    // ===================================================
    // COMENTÁRIOS
    // ===================================================
    addComment() {
        const { row, col } = this.state.selectedCell;
        const key = this.getCellKey(row, col);
        const existing = this.sheet.comments[key] || '';
        const text = prompt('Comentário para ' + this.getCellReference(row, col) + ':', existing);
        if (text === null) return;
        if (text.trim() === '') delete this.sheet.comments[key];
        else this.sheet.comments[key] = text;
        this.render();
        this.autosave();
    }

    // ===================================================
    // BUSCA DE FUNÇÕES / MODAL
    // ===================================================
    buildFunctionList() {
        this.functionDescriptions = {
            SUM: 'Soma os valores em um intervalo', AVERAGE: 'Calcula a média', MAX: 'Retorna o maior valor', MIN: 'Retorna o menor valor',
            COUNT: 'Conta células numéricas', COUNTA: 'Conta células não vazias', COUNTBLANK: 'Conta células vazias',
            MEDIAN: 'Retorna a mediana', STDEV: 'Desvio padrão', VAR: 'Variância', RANK: 'Classifica um valor',
            LARGE: 'K-ésimo maior valor', SMALL: 'K-ésimo menor valor', MODE: 'Valor mais frequente',
            IF: 'Teste lógico condicional', AND: 'Verdadeiro se todos forem verdadeiros', OR: 'Verdadeiro se algum for verdadeiro',
            NOT: 'Inverte um valor lógico', IFERROR: 'Retorna valor alternativo em caso de erro', ISERROR: 'Verifica se é erro',
            ISBLANK: 'Verifica se está vazio', ISNUMBER: 'Verifica se é número', ISTEXT: 'Verifica se é texto',
            SUMIF: 'Soma com condição', COUNTIF: 'Conta com condição', SUMIFS: 'Soma com múltiplas condições', COUNTIFS: 'Conta com múltiplas condições',
            AVERAGEIF: 'Média com condição', VLOOKUP: 'Procura vertical', HLOOKUP: 'Procura horizontal',
            INDEX: 'Retorna valor por posição', MATCH: 'Retorna posição de um valor', CHOOSE: 'Escolhe valor de uma lista',
            CONCATENATE: 'Junta textos', CONCAT: 'Junta textos/intervalos', LEFT: 'Extrai caracteres à esquerda', RIGHT: 'Extrai caracteres à direita',
            MID: 'Extrai do meio do texto', LEN: 'Conta caracteres', UPPER: 'Maiúsculas', LOWER: 'Minúsculas', PROPER: 'Capitaliza palavras',
            TRIM: 'Remove espaços extras', SUBSTITUTE: 'Substitui texto', REPLACE: 'Substitui por posição', FIND: 'Localiza texto (sensível)',
            SEARCH: 'Localiza texto (não sensível)', TEXT: 'Formata número como texto', VALUE: 'Converte texto em número', REPT: 'Repete texto', EXACT: 'Compara textos',
            ROUND: 'Arredonda', ROUNDUP: 'Arredonda para cima', ROUNDDOWN: 'Arredonda para baixo', ABS: 'Valor absoluto',
            SQRT: 'Raiz quadrada', POWER: 'Potência', MOD: 'Resto da divisão', INT: 'Parte inteira', TRUNC: 'Trunca decimais',
            PI: 'Retorna π', RAND: 'Número aleatório 0-1', RANDBETWEEN: 'Número aleatório entre dois valores',
            CEILING: 'Arredonda para cima (múltiplo)', FLOOR: 'Arredonda para baixo (múltiplo)', SIGN: 'Sinal do número', PRODUCT: 'Multiplica valores',
            TODAY: 'Data atual', NOW: 'Data e hora atuais', DATE: 'Cria uma data', YEAR: 'Extrai o ano', MONTH: 'Extrai o mês', DAY: 'Extrai o dia', WEEKDAY: 'Dia da semana'
        };
    }

    openFunctionModal() {
        const modal = document.getElementById('functionModal');
        modal.style.display = 'flex';
        const list = document.getElementById('functionList');
        const render = (filter = '') => {
            list.innerHTML = '';
            Object.keys(this.functionDescriptions).sort().forEach(fn => {
                if (filter && !fn.toLowerCase().includes(filter.toLowerCase()) && !this.functionDescriptions[fn].toLowerCase().includes(filter.toLowerCase())) return;
                const div = document.createElement('div');
                div.className = 'function-item';
                div.dataset.function = fn;
                div.innerHTML = `<strong>${fn}</strong> - ${this.functionDescriptions[fn]}`;
                div.onclick = () => {
                    document.querySelectorAll('.function-item').forEach(el => el.classList.remove('selected'));
                    div.classList.add('selected');
                };
                list.appendChild(div);
            });
        };
        render();
        document.getElementById('functionSearch').value = '';
        document.getElementById('functionSearch').oninput = (e) => render(e.target.value);

        document.getElementById('insertFunction').onclick = () => {
            const selected = document.querySelector('.function-item.selected');
            if (selected) {
                const func = selected.dataset.function;
                const template = this.getFunctionTemplate(func);
                const { row, col } = this.state.selectedCell;
                this.startEditing(row, col, template);
                modal.style.display = 'none';
            }
        };
        document.getElementById('cancelFunction').onclick = () => { modal.style.display = 'none'; };
    }

    getFunctionTemplate(func) {
        const arity = {
            SUM: '(A1:A10)', AVERAGE: '(A1:A10)', MAX: '(A1:A10)', MIN: '(A1:A10)', COUNT: '(A1:A10)', COUNTA: '(A1:A10)',
            IF: '(A1>10,"Sim","Não")', SUMIF: '(A1:A10,">10")', COUNTIF: '(A1:A10,">10")', VLOOKUP: '(A1,A1:B10,2,FALSE)',
            HLOOKUP: '(A1,A1:B10,2,FALSE)', INDEX: '(A1:B10,1,1)', MATCH: '(A1,A1:A10,0)', CONCATENATE: '(A1," ",B1)',
            LEFT: '(A1,5)', RIGHT: '(A1,5)', MID: '(A1,2,3)', LEN: '(A1)', UPPER: '(A1)', LOWER: '(A1)', PROPER: '(A1)',
            TODAY: '()', NOW: '()', ROUND: '(A1,2)', ABS: '(A1)', IFERROR: '(A1/B1,"erro")'
        };
        return '=' + func + (arity[func] || '(A1)');
    }

    openHelpModal() { document.getElementById('helpModal').style.display = 'flex'; }

    openGenericModal(title, bodyHTML, onOk) {
        document.getElementById('genericModalTitle').textContent = title;
        document.getElementById('genericModalBody').innerHTML = bodyHTML;
        const modal = document.getElementById('genericModal');
        modal.style.display = 'flex';
        document.getElementById('genericModalOk').onclick = () => { onOk(); modal.style.display = 'none'; };
        document.getElementById('genericModalCancel').onclick = () => { modal.style.display = 'none'; };
        return modal;
    }

    // ===================================================
    // MENU / RIBBON
    // ===================================================
    switchMenu(menu) {
        document.querySelectorAll('.menu-item').forEach(el => el.classList.toggle('active', el.dataset.menu === menu));
        document.querySelectorAll('.ribbon-tab').forEach(el => {
            const isTarget = el.id === 'tab-' + menu;
            el.classList.toggle('active', isTarget);
            el.style.display = isTarget ? 'flex' : 'none';
        });
    }

    // ===================================================
    // STATUS / UI
    // ===================================================
    updateUI() {
        this.DOM.zoomLevel.textContent = this.state.zoom + '%';
        this.renderSheetTabs();
        this.markUnsaved();
    }

    updateStatus() {
        const { row, col } = this.state.selectedCell;
        const { r1, r2, c1, c2 } = this.normSelection();
        this.DOM.statusCell.textContent = `Célula: ${this.getCellReference(row, col)}`;
        this.DOM.statusSelection.textContent = `Seleção: ${r2 - r1 + 1}x${c2 - c1 + 1}`;
        this.DOM.cellReference.textContent = (r1 !== r2 || c1 !== c2)
            ? `${this.getCellReference(r1, c1)}:${this.getCellReference(r2, c2)}`
            : this.getCellReference(row, col);
        if (!this.state.isEditing) this.DOM.formulaField.value = this.getCellRaw(row, col);

        let sum = 0, count = 0;
        this.forEachSelectedCell((r, c) => {
            const v = parseFloat(this.getDisplayValue(r, c));
            if (!isNaN(v)) { sum += v; count++; }
        });
        this.DOM.statusAgg.textContent = count > 1 ? `Soma: ${sum.toLocaleString('pt-BR')}  Média: ${(sum / count).toLocaleString('pt-BR', { maximumFractionDigits: 2 })}  Contagem: ${count}` : '';
    }

    // ===================================================
    // ZOOM
    // ===================================================
    zoomTo(z) {
        this.state.zoom = Math.max(50, Math.min(200, z));
        this.applyZoom();
    }
    zoomIn() { this.zoomTo(this.state.zoom + 10); }
    zoomOut() { this.zoomTo(this.state.zoom - 10); }
    applyZoom() {
        this.DOM.sheetScroll.style.transform = `scale(${this.state.zoom / 100})`;
        this.DOM.sheetScroll.style.transformOrigin = 'top left';
        this.DOM.zoomLevel.textContent = this.state.zoom + '%';
    }

    // ===================================================
    // CONGELAR PAINÉIS / GRADE
    // ===================================================
    toggleFreezeRow() {
        this.sheet.freeze.row = !this.sheet.freeze.row;
        this.render();
    }
    toggleFreezeCol() {
        this.sheet.freeze.col = !this.sheet.freeze.col;
        this.render();
    }
    unfreeze() {
        this.sheet.freeze = { row: false, col: false };
        this.render();
    }
    toggleGridlines() {
        this.sheet.showGridlines = !this.sheet.showGridlines;
        this.render();
    }

    setPrintOrientation(orientation) {
        this.state.orientation = orientation;
        document.body.classList.toggle('landscape', orientation === 'landscape');
        const styleEl = document.getElementById('printOrientationStyle');
        if (styleEl) styleEl.textContent = `@page { size: ${orientation}; margin: 10mm; }`;
    }

    // ===================================================
    // PWA
    // ===================================================
    setupPWA() {
        if ('serviceWorker' in navigator) {
            navigator.serviceWorker.register('sw.js').catch(() => {});
        }
    }

    // ===================================================
    // MENU DE CONTEXTO
    // ===================================================
    showContextMenu(x, y) {
        const menu = document.getElementById('contextMenu');
        menu.style.display = 'block';
        const maxX = window.innerWidth - menu.offsetWidth - 8;
        const maxY = window.innerHeight - menu.offsetHeight - 8;
        menu.style.left = Math.min(x, maxX) + 'px';
        menu.style.top = Math.min(y, maxY) + 'px';
    }
    hideContextMenu() {
        document.getElementById('contextMenu').style.display = 'none';
    }
    handleContextAction(action) {
        switch (action) {
            case 'cut': this.cut(); break;
            case 'copy': this.copy(); break;
            case 'paste': this.paste(); break;
            case 'insertRow': this.insertRow(); break;
            case 'deleteRow': this.deleteRow(); break;
            case 'insertCol': this.insertColumn(); break;
            case 'deleteCol': this.deleteColumn(); break;
            case 'clear': this.clearCells(); break;
            case 'comment': this.addComment(); break;
        }
    }

    // ===================================================
    // MODAIS DE RECURSOS (gráfico de dados, localizar, formatação condicional, etc.)
    // ===================================================
    openFindReplaceModal() {
        const body = `
            <div class="form-row"><label>Localizar:</label><input type="text" id="findTerm" placeholder="Texto a localizar"></div>
            <div class="form-row"><label>Substituir por:</label><input type="text" id="replaceTerm" placeholder="(opcional)"></div>
            <div class="form-row"><label id="findResultMsg" style="font-size:12px;color:#666;"></label></div>
        `;
        const modal = this.openGenericModal('Localizar e Substituir', body, () => {
            const term = document.getElementById('findTerm').value;
            if (!term) return;
            const replacement = document.getElementById('replaceTerm').value;
            const found = this.findReplace(term, replacement || replacement === '' ? replacement : null, true);
            this.DOM.statusReady.textContent = `${found} ocorrência(s) processada(s)`;
        });
        setTimeout(() => document.getElementById('findTerm') && document.getElementById('findTerm').focus(), 50);
    }

    openConditionalFormatModal() {
        const body = `
            <div class="form-row"><label>Condição:</label>
                <select id="cfOp">
                    <option value=">">Maior que</option>
                    <option value="<">Menor que</option>
                    <option value="=">Igual a</option>
                    <option value=">=">Maior ou igual a</option>
                    <option value="<=">Menor ou igual a</option>
                </select>
            </div>
            <div class="form-row"><label>Valor:</label><input type="number" id="cfValue" value="0"></div>
            <div class="form-row"><label>Cor:</label><input type="color" id="cfColor" value="#ffd966"></div>
        `;
        this.openGenericModal('Formatação Condicional', body, () => {
            const op = document.getElementById('cfOp').value;
            const value = document.getElementById('cfValue').value;
            const color = document.getElementById('cfColor').value;
            this.addConditionalFormat(op, value, color);
        });
    }

    openDataValidationModal() {
        const body = `
            <div class="form-row"><label>Lista de valores (separados por vírgula):</label>
            <input type="text" id="dvList" placeholder="Ex: Sim,Não,Talvez"></div>
        `;
        this.openGenericModal('Validação de Dados', body, () => {
            const list = document.getElementById('dvList').value.split(',').map(s => s.trim()).filter(Boolean);
            this.forEachSelectedCell((r, c) => {
                this.sheet.validations[this.getCellKey(r, c)] = list;
            });
            this.autosave();
        });
    }

    openColWidthModal() {
        const body = `<div class="form-row"><label>Largura (px):</label><input type="number" id="colWidthInput" value="90" min="30"></div>`;
        this.openGenericModal('Largura da Coluna', body, () => {
            const w = parseInt(document.getElementById('colWidthInput').value) || 90;
            const { c1, c2 } = this.normSelection();
            for (let c = c1; c <= c2; c++) this.sheet.colWidths[c] = w;
            this.render();
            this.autosave();
        });
    }

    openRowHeightModal() {
        const body = `<div class="form-row"><label>Altura (px):</label><input type="number" id="rowHeightInput" value="28" min="16"></div>`;
        this.openGenericModal('Altura da Linha', body, () => {
            const h = parseInt(document.getElementById('rowHeightInput').value) || 28;
            const { r1, r2 } = this.normSelection();
            for (let r = r1; r <= r2; r++) this.sheet.rowHeights[r] = h;
            this.render();
            this.autosave();
        });
    }

    openHyperlinkModal() {
        const body = `<div class="form-row"><label>URL:</label><input type="text" id="linkUrl" placeholder="https://..."></div>
            <div class="form-row"><label>Texto exibido:</label><input type="text" id="linkText" placeholder="(opcional)"></div>`;
        this.openGenericModal('Inserir Hiperlink', body, () => {
            const url = document.getElementById('linkUrl').value.trim();
            const text = document.getElementById('linkText').value.trim() || url;
            if (!url) return;
            const { row, col } = this.state.selectedCell;
            this.setCellValue(row, col, text);
            const key = this.getCellKey(row, col);
            this.sheet.styles[key].color = '#0563C1';
            this.sheet.styles[key].underline = true;
            this.sheet.validations[key + ':link'] = url;
            this.render();
            this.autosave();
        });
    }

    openPivotTableModal() {
        const body = `<p style="font-size:13px;color:#555;line-height:1.5;">
            Selecione um intervalo com cabeçalhos e clique em OK para gerar uma tabela dinâmica simples
            (contagem e soma agrupadas pela primeira coluna) em uma nova planilha.</p>`;
        this.openGenericModal('Tabela Dinâmica', body, () => this.generateSimplePivot());
    }

    generateSimplePivot() {
        const { r1, r2, c1, c2 } = this.normSelection();
        if (r2 <= r1) { alert('Selecione um intervalo com cabeçalho e ao menos uma linha de dados.'); return; }
        const groups = {};
        for (let r = r1 + 1; r <= r2; r++) {
            const key = String(this.getDisplayValue(r, c1));
            if (!groups[key]) groups[key] = { count: 0, sum: 0 };
            groups[key].count++;
            for (let c = c1 + 1; c <= c2; c++) {
                const v = parseFloat(this.getDisplayValue(r, c));
                if (!isNaN(v)) groups[key].sum += v;
            }
        }
        const sheet = this.makeEmptySheet(this.uniqueSheetName('TabelaDinamica'));
        sheet.data['0,0'] = String(this.getDisplayValue(r1, c1));
        sheet.data['0,1'] = 'Contagem';
        sheet.data['0,2'] = 'Soma';
        sheet.styles['0,0'] = { ...this.getDefaultStyle(), bold: true, backgroundColor: '#107C41', color: '#fff' };
        sheet.styles['0,1'] = { ...this.getDefaultStyle(), bold: true, backgroundColor: '#107C41', color: '#fff' };
        sheet.styles['0,2'] = { ...this.getDefaultStyle(), bold: true, backgroundColor: '#107C41', color: '#fff' };
        Object.keys(groups).forEach((k, i) => {
            sheet.data[(i + 1) + ',0'] = k;
            sheet.data[(i + 1) + ',1'] = groups[k].count;
            sheet.data[(i + 1) + ',2'] = groups[k].sum;
        });
        this.state.sheets.push(sheet);
        this.switchSheet(this.state.sheets.length - 1);
        this.autosave();
    }

    // ===================================================
    // VINCULAÇÃO DE EVENTOS (defensiva - nunca trava a UI)
    // ===================================================
    on(id, event, handler) {
        const el = document.getElementById(id);
        if (!el) { console.warn(`[ExcelSimulador] elemento não encontrado: #${id}`); return; }
        el.addEventListener(event, handler);
    }

    bindEvents() {
        // Menu principal (abas do ribbon)
        document.querySelectorAll('.menu-item').forEach(item => {
            item.addEventListener('click', () => this.switchMenu(item.dataset.menu));
        });

        // Barra de fórmulas
        this.DOM.formulaField.addEventListener('focus', () => {
            if (!this.state.isEditing) {
                this.state.isEditing = true;
                const cell = this.getCellElement(this.state.selectedCell.row, this.state.selectedCell.col);
                if (cell) cell.classList.add('editing');
            }
        });
        this.DOM.formulaField.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); this.finishEditing('down'); }
            else if (e.key === 'Tab') { e.preventDefault(); this.finishEditing('right'); }
            else if (e.key === 'Escape') { e.preventDefault(); this.cancelEditing(); }
        });
        this.DOM.formulaField.addEventListener('blur', () => { if (this.state.isEditing) this.finishEditing(); });
        this.on('formula-functions', 'click', () => this.openFunctionModal());

        // Seleção via mouse (arrastar / alça de preenchimento)
        document.addEventListener('mouseup', (e) => {
            if (this.state.isFilling) {
                const target = document.elementFromPoint(e.clientX, e.clientY);
                if (target && target.closest('td[data-row]')) {
                    const td = target.closest('td[data-row]');
                    this.finishFill(parseInt(td.dataset.row), parseInt(td.dataset.col));
                }
                this.state.isFilling = false;
            }
            this.state.isSelecting = false;
        });
        this.DOM.spreadsheet.addEventListener('mousemove', (e) => {
            if (!this.state.isFilling) return;
            const target = e.target.closest('td[data-row]');
            if (!target) return;
        });
        this.DOM.spreadsheet.addEventListener('contextmenu', (e) => {
            const td = e.target.closest('td[data-row]');
            if (!td) return;
            e.preventDefault();
            const r = parseInt(td.dataset.row), c = parseInt(td.dataset.col);
            const inSel = (() => { const { r1, r2, c1, c2 } = this.normSelection(); return r >= r1 && r <= r2 && c >= c1 && c <= c2; })();
            if (!inSel) this.selectCell(r, c);
            this.showContextMenu(e.clientX, e.clientY);
        });
        document.addEventListener('click', (e) => {
            if (!e.target.closest('#contextMenu')) this.hideContextMenu();
        });
        document.querySelectorAll('.context-item').forEach(item => {
            item.addEventListener('click', () => { this.handleContextAction(item.dataset.action); this.hideContextMenu(); });
        });

        // Direct typing on a selected cell starts editing
        document.addEventListener('keypress', (e) => {
            if (this.state.isEditing) return;
            if (document.activeElement === this.DOM.formulaField) return;
            if (e.ctrlKey || e.metaKey || e.altKey) return;
            if (!this.state.selectedCell) return;
            if (e.key.length === 1) {
                const { row, col } = this.state.selectedCell;
                this.startEditing(row, col, e.key);
            }
        });

        // Formatação de fonte
        this.on('btnBold', 'click', () => this.toggleStyle('bold'));
        this.on('btnItalic', 'click', () => this.toggleStyle('italic'));
        this.on('btnUnderline', 'click', () => this.toggleStyle('underline'));
        this.on('btnStrike', 'click', () => this.toggleStyle('strike'));
        this.on('fontFamily', 'change', (e) => this.applyStyleToSelection('fontFamily', e.target.value));
        this.on('fontSize', 'change', (e) => this.applyStyleToSelection('fontSize', e.target.value));
        this.on('numberFormat', 'change', (e) => this.applyNumberFormat(e.target.value));
        this.on('btnIncreaseDecimal', 'click', () => this.adjustDecimals(1));
        this.on('btnDecreaseDecimal', 'click', () => this.adjustDecimals(-1));

        // Alinhamento
        this.on('btnAlignLeft', 'click', () => this.applyStyleToSelection('align', 'left'));
        this.on('btnAlignCenter', 'click', () => this.applyStyleToSelection('align', 'center'));
        this.on('btnAlignRight', 'click', () => this.applyStyleToSelection('align', 'right'));
        this.on('btnAlignTop', 'click', () => this.applyStyleToSelection('valign', 'top'));
        this.on('btnAlignMiddle', 'click', () => this.applyStyleToSelection('valign', 'middle'));
        this.on('btnWrapText', 'click', () => this.toggleStyle('wrapText'));

        // Cores e bordas
        this.on('btnCellColor', 'click', () => {
            const panel = document.getElementById('side-panel');
            panel.style.display = 'block';
        });
        this.on('btnFontColor', 'click', () => { document.getElementById('side-panel').style.display = 'block'; });
        this.on('btnBorder', 'click', () => { document.getElementById('side-panel').style.display = 'block'; });
        this.on('closePanel', 'click', () => { document.getElementById('side-panel').style.display = 'none'; });
        this.on('fillColorPicker', 'input', (e) => this.applyStyleToSelection('backgroundColor', e.target.value));
        this.on('fontColorPicker', 'input', (e) => this.applyStyleToSelection('color', e.target.value));
        this.on('borderAll', 'click', () => this.applyStyleToSelection('border', '1px solid #444'));
        this.on('borderOutside', 'click', () => this.applyStyleToSelection('border', '2px solid #444'));
        this.on('borderNone', 'click', () => this.applyStyleToSelection('border', 'none'));
        this.on('btnConditionalFormat', 'click', () => this.openConditionalFormatModal());

        // Mesclar / Edição
        this.on('btnMerge', 'click', () => this.mergeCells());
        this.on('btnUnmerge', 'click', () => this.unmergeCells());
        this.on('btnInsertRow', 'click', () => this.insertRow());
        this.on('btnDeleteRow', 'click', () => this.deleteRow());
        this.on('btnInsertCol', 'click', () => this.insertColumn());
        this.on('btnDeleteCol', 'click', () => this.deleteColumn());
        this.on('btnClear', 'click', () => this.clearCells());
        this.on('btnFind', 'click', () => this.openFindReplaceModal());

        // Área de transferência
        this.on('btnCopy', 'click', () => this.copy());
        this.on('btnCut', 'click', () => this.cut());
        this.on('btnPaste', 'click', () => this.paste());
        this.on('btnFormatPainter', 'click', () => {
            const { r1, c1 } = this.normSelection();
            this.state.formatPainter = { ...this.sheet.styles[this.getCellKey(r1, c1)] };
            this.DOM.statusReady.textContent = 'Pincel de formatação ativo — clique na célula de destino';
        });
        this.DOM.spreadsheet.addEventListener('click', (e) => {
            if (!this.state.formatPainter) return;
            const td = e.target.closest('td[data-row]');
            if (!td) return;
            this.forEachSelectedCell((r, c) => {
                this.sheet.styles[this.getCellKey(r, c)] = { ...this.state.formatPainter };
            });
            this.state.formatPainter = null;
            this.render();
            this.selectCell(this.state.selectedCell.row, this.state.selectedCell.col);
        });

        // Desfazer / Refazer
        this.on('btnUndo', 'click', () => this.undo());
        this.on('btnRedo', 'click', () => this.redo());

        // AutoSoma / Funções
        this.on('btnAutoSum', 'click', () => this.autoSum());
        this.on('btnAutoSum2', 'click', () => this.autoSum());
        this.on('btnFunction', 'click', () => this.openFunctionModal());
        this.on('btnFunction2', 'click', () => this.openFunctionModal());

        // Arquivo
        this.on('btnNew', 'click', () => {
            if (confirm('Criar nova pasta de trabalho? Dados não salvos serão perdidos.')) {
                localStorage.removeItem('excelSimuladorPro.autosave');
                location.reload();
            }
        });
        this.on('btnOpen', 'click', () => this.importCSV());
        this.on('btnSave', 'click', () => this.saveWorkbookXLSX());
        this.on('btnSaveAs', 'click', () => this.saveWorkbookXLSX());
        this.on('btnExportCSV', 'click', () => this.exportCSV());
        this.on('btnImportCSV', 'click', () => this.importCSV());
        this.on('btnExportPDF', 'click', () => this.exportPDF());
        this.on('btnTemplates', 'click', () => this.openTemplatesModal());
        this.on('btnOptions', 'click', () => alert('Painel de opções: zoom, grade e fórmulas podem ser ajustados nas abas Exibição e Layout da Página.'));
        this.on('btnHelp', 'click', () => this.openHelpModal());

        // Inserir - gráficos
        this.on('btnBarChart', 'click', () => this.createChart('bar'));
        this.on('btnPieChart', 'click', () => this.createChart('pie'));
        this.on('btnLineChart', 'click', () => this.createChart('line'));
        this.on('btnAreaChart', 'click', () => this.createChart('area'));
        this.on('btnDoughnutChart', 'click', () => this.createChart('doughnut'));
        this.on('btnRadarChart', 'click', () => this.createChart('radar'));
        this.on('btnTable', 'click', () => this.applyStyleToSelection('backgroundColor', '#eef6f0'));
        this.on('btnPivotTable', 'click', () => this.openPivotTableModal());
        this.on('btnHyperlink', 'click', () => this.openHyperlinkModal());
        this.on('btnComment', 'click', () => this.addComment());
        this.on('btnImage', 'click', () => alert('Para inserir imagens, use um editor de texto (Word/PowerPoint) — este recurso ainda não está disponível na planilha.'));
        this.on('btnTextBox', 'click', () => alert('Caixas de texto ainda não são suportadas nesta versão.'));

        // Layout da página
        this.on('btnOrientationPortrait', 'click', () => this.setPrintOrientation('portrait'));
        this.on('btnOrientationLandscape', 'click', () => this.setPrintOrientation('landscape'));
        this.on('btnPrintArea', 'click', () => alert('Área de impressão definida pela seleção atual ao exportar/imprimir.'));
        this.on('btnColWidth', 'click', () => this.openColWidthModal());
        this.on('btnRowHeight', 'click', () => this.openRowHeightModal());
        this.on('btnToggleGridlines', 'click', () => this.toggleGridlines());
        this.on('btnToggleHeadings', 'click', () => this.DOM.spreadsheet.classList.toggle('no-headings'));

        // Fórmulas
        this.on('btnShowFormulas', 'click', () => { this.state.showFormulas = !this.state.showFormulas; this.render(); });
        this.on('btnTraceErrors', 'click', () => this.traceErrors());
        this.on('btnRecalculate', 'click', () => { this.render(); this.DOM.statusReady.textContent = 'Recalculado!'; });

        // Dados
        this.on('btnSortAsc', 'click', () => this.sortSelection(true));
        this.on('btnSortDesc', 'click', () => this.sortSelection(false));
        this.on('btnFilter', 'click', () => alert('Use Classificar Crescente/Decrescente na aba Dados para organizar sua tabela. Filtro interativo completo não está disponível nesta versão.'));
        this.on('btnDataValidation', 'click', () => this.openDataValidationModal());
        this.on('btnRemoveDuplicates', 'click', () => this.removeDuplicates());
        this.on('btnTextToColumns', 'click', () => this.textToColumns());
        this.on('btnGroupRows', 'click', () => alert('Agrupamento de linhas/colunas ainda não é suportado nesta versão.'));
        this.on('btnUngroupRows', 'click', () => {});

        // Revisão
        this.on('btnSpellCheck', 'click', () => alert('Verificação ortográfica não está disponível nesta versão web.'));
        this.on('btnNewComment', 'click', () => this.addComment());
        this.on('btnShowComments', 'click', () => {
            const entries = Object.entries(this.sheet.comments);
            if (!entries.length) { alert('Nenhum comentário nesta planilha.'); return; }
            alert(entries.map(([key, text]) => {
                const [r, c] = key.split(',').map(Number);
                return `${this.getCellReference(r, c)}: ${text}`;
            }).join('\n'));
        });
        this.on('btnProtectSheet', 'click', () => alert('Proteção de planilha simulada — nesta versão os dados continuam editáveis.'));

        // Exibição
        this.on('btnNormalView', 'click', () => {});
        this.on('btnPageLayoutView', 'click', () => alert('Modo de exibição de layout de página simulado via impressão (Arquivo > Exportar/Imprimir PDF).'));
        this.on('btnToggleGridlinesView', 'click', () => this.toggleGridlines());
        this.on('btnToggleFormulaBar', 'click', () => {
            const bar = document.getElementById('formula-bar');
            bar.style.display = bar.style.display === 'none' ? 'flex' : 'none';
        });
        this.on('btnFreezeTopRow', 'click', () => this.toggleFreezeRow());
        this.on('btnFreezeFirstCol', 'click', () => this.toggleFreezeCol());
        this.on('btnUnfreeze', 'click', () => this.unfreeze());
        this.on('btnZoomIn2', 'click', () => this.zoomIn());
        this.on('btnZoomOut2', 'click', () => this.zoomOut());
        this.on('btnZoom100', 'click', () => this.zoomTo(100));

        // Barra de status
        this.on('btnAddSheet', 'click', () => this.addSheet());
        this.on('zoomIn', 'click', () => this.zoomIn());
        this.on('zoomOut', 'click', () => this.zoomOut());

        // Modais - fechar
        document.querySelectorAll('.close').forEach(el => {
            el.addEventListener('click', () => { document.querySelectorAll('.modal').forEach(m => m.style.display = 'none'); });
        });
        window.addEventListener('click', (e) => {
            if (e.target.classList.contains('modal')) e.target.style.display = 'none';
        });
        this.on('closeHelp', 'click', () => { document.getElementById('helpModal').style.display = 'none'; });

        // Atalhos de teclado
        document.addEventListener('keydown', (e) => this.handleKeyDown(e));

        // Redimensionamento da janela (recalcula posições de gráficos etc. se necessário)
        window.addEventListener('beforeunload', (e) => {
            // autosave already handles persistence; no forced dialog to keep UX smooth
        });
    }

    traceErrors() {
        const s = this.sheet;
        let errors = [];
        for (let r = 0; r < s.rowCount; r++) {
            for (let c = 0; c < s.colCount; c++) {
                const v = this.getDisplayValue(r, c);
                if (typeof v === 'string' && /^#[A-Z/!?0-9]+$/.test(v)) errors.push(this.getCellReference(r, c) + ': ' + v);
            }
        }
        alert(errors.length ? `Erros encontrados:\n${errors.join('\n')}` : 'Nenhum erro encontrado na planilha.');
    }

    removeDuplicates() {
        const { r1, r2, c1, c2 } = this.normSelection();
        if (r2 <= r1) { alert('Selecione um intervalo com mais de uma linha.'); return; }
        const seen = new Set();
        const rowsToKeep = [];
        for (let r = r1; r <= r2; r++) {
            const rowVals = [];
            for (let c = c1; c <= c2; c++) rowVals.push(this.getCellRaw(r, c));
            const sig = JSON.stringify(rowVals);
            if (!seen.has(sig)) { seen.add(sig); rowsToKeep.push(rowVals); }
        }
        const removed = (r2 - r1 + 1) - rowsToKeep.length;
        rowsToKeep.forEach((rowVals, i) => {
            rowVals.forEach((val, ci) => this.setCellRaw(r1 + i, c1 + ci, val));
        });
        for (let r = r1 + rowsToKeep.length; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) this.setCellRaw(r, c, '');
        }
        this.render();
        this.autosave();
        this.DOM.statusReady.textContent = `${removed} linha(s) duplicada(s) removida(s)`;
    }

    textToColumns() {
        const { r1, r2, c1 } = this.normSelection();
        for (let r = r1; r <= r2; r++) {
            const val = String(this.getCellRaw(r, c1));
            if (!val.includes(',') && !val.includes(';')) continue;
            const parts = val.split(/[,;]\s*/);
            parts.forEach((p, i) => this.setCellRaw(r, c1 + i, p));
        }
        this.render();
        this.autosave();
    }

    handleKeyDown(e) {
        const ctrl = e.ctrlKey || e.metaKey;
        if (ctrl && e.key.toLowerCase() === 'z') { e.preventDefault(); this.undo(); return; }
        if (ctrl && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) { e.preventDefault(); this.redo(); return; }
        if (ctrl && e.key.toLowerCase() === 'c') { e.preventDefault(); this.copy(); return; }
        if (ctrl && e.key.toLowerCase() === 'v') { e.preventDefault(); this.paste(); return; }
        if (ctrl && e.key.toLowerCase() === 'x') { e.preventDefault(); this.cut(); return; }
        if (ctrl && e.key.toLowerCase() === 'b') { e.preventDefault(); this.toggleStyle('bold'); return; }
        if (ctrl && e.key.toLowerCase() === 'i') { e.preventDefault(); this.toggleStyle('italic'); return; }
        if (ctrl && e.key.toLowerCase() === 'u') { e.preventDefault(); this.toggleStyle('underline'); return; }
        if (ctrl && e.key.toLowerCase() === 'f') { e.preventDefault(); this.openFindReplaceModal(); return; }
        if (ctrl && e.key.toLowerCase() === 's') { e.preventDefault(); this.saveWorkbookXLSX(); return; }

        if (this.state.isEditing) return;

        if (e.key === 'Delete' || e.key === 'Backspace') {
            if (this.state.selectedCell) { e.preventDefault(); this.clearCells(); }
            return;
        }
        if (e.key === 'F2') {
            e.preventDefault();
            if (this.state.selectedCell) this.startEditing(this.state.selectedCell.row, this.state.selectedCell.col);
            return;
        }
        if (e.key === 'Enter') {
            e.preventDefault();
            this.moveSelection('down');
            return;
        }
        if (e.key === 'Tab') {
            e.preventDefault();
            this.moveSelection(e.shiftKey ? 'left' : 'right');
            return;
        }
        if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
            e.preventDefault();
            let { row, col } = this.state.selectedCell;
            if (e.key === 'ArrowUp') row = Math.max(0, row - 1);
            if (e.key === 'ArrowDown') row = Math.min(this.sheet.rowCount - 1, row + 1);
            if (e.key === 'ArrowLeft') col = Math.max(0, col - 1);
            if (e.key === 'ArrowRight') col = Math.min(this.sheet.colCount - 1, col + 1);
            this.selectCell(row, col, e.shiftKey);
            return;
        }
        if (ctrl && e.key === 'Home') { e.preventDefault(); this.selectCell(0, 0); return; }
        if (ctrl && e.key === 'End') {
            e.preventDefault();
            this.selectCell(this.sheet.rowCount - 1, this.sheet.colCount - 1);
            return;
        }
    }
}

// ===== INICIALIZAÇÃO =====
document.addEventListener('DOMContentLoaded', () => {
    const app = new ExcelApplication();
    window.excelApp = app;
});