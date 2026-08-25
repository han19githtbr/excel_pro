// ============================================
// DATABASE E UTILITÁRIOS PARA O EXCEL
// ============================================

const ExcelDatabase = {
    // Templates de planilhas
    templates: {
        'Orçamento Mensal': {
            data: [
                ['Categoria', 'Janeiro', 'Fevereiro', 'Março', 'Total'],
                ['Alimentação', 800, 750, 900, '=SUM(B2:D2)'],
                ['Transporte', 300, 320, 280, '=SUM(B3:D3)'],
                ['Moradia', 1200, 1200, 1200, '=SUM(B4:D4)'],
                ['Lazer', 150, 200, 180, '=SUM(B5:D5)'],
                ['Total', '=SUM(B2:B5)', '=SUM(C2:C5)', '=SUM(D2:D5)', '=SUM(B6:D6)']
            ]
        },
        'Controle de Estoque': {
            data: [
                ['Produto', 'Quantidade', 'Mínimo', 'Máximo', 'Status'],
                ['Mouse Gamer', 45, 10, 100, '=IF(B2<C2,"Baixo",IF(B2>D2,"Alto","OK"))'],
                ['Teclado Mecânico', 23, 5, 50, '=IF(B3<C3,"Baixo",IF(B3>D3,"Alto","OK"))'],
                ['Monitor 24"', 12, 3, 20, '=IF(B4<C4,"Baixo",IF(B4>D4,"Alto","OK"))'],
                ['Headset', 34, 8, 40, '=IF(B5<C5,"Baixo",IF(B5>D5,"Alto","OK"))']
            ]
        },
        'Vendas por Região': {
            data: [
                ['Região', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'],
                ['Norte', 15000, 18000, 22000, 20000, '=SUM(B2:E2)'],
                ['Sul', 18000, 21000, 19000, 23000, '=SUM(B3:E3)'],
                ['Leste', 12000, 14000, 16000, 15000, '=SUM(B4:E4)'],
                ['Oeste', 20000, 19000, 21000, 24000, '=SUM(B5:E5)'],
                ['Total', '=SUM(B2:B5)', '=SUM(C2:C5)', '=SUM(D2:D5)', '=SUM(E2:E5)', '=SUM(F2:F5)']
            ]
        }
    },

    // Funções de análise estatística
    statistics: {
        mean: (arr) => arr.reduce((a, b) => a + b, 0) / arr.length,
        median: (arr) => {
            const sorted = [...arr].sort((a, b) => a - b);
            const mid = Math.floor(sorted.length / 2);
            return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
        },
        mode: (arr) => {
            const freq = {};
            arr.forEach(v => freq[v] = (freq[v] || 0) + 1);
            let max = 0, mode = null;
            for (const [key, val] of Object.entries(freq)) {
                if (val > max) { max = val; mode = key; }
            }
            return mode;
        },
        variance: (arr) => {
            const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
            return arr.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / arr.length;
        },
        stddev: (arr) => Math.sqrt(ExcelDatabase.statistics.variance(arr))
    },

    // Formatação de números
    format: {
        currency: (value) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value),
        percent: (value) => new Intl.NumberFormat('pt-BR', { style: 'percent', minimumFractionDigits: 2 }).format(value / 100),
        number: (value) => new Intl.NumberFormat('pt-BR').format(value),
        date: (value) => new Intl.DateTimeFormat('pt-BR').format(new Date(value)),
        time: (value) => new Intl.DateTimeFormat('pt-BR', { hour: '2-digit', minute: '2-digit' }).format(new Date(value)),
        scientific: (value) => value.toExponential(2),
        fraction: (value) => {
            const frac = value % 1;
            if (frac === 0) return String(value);
            const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
            const denominator = 1000;
            const numerator = Math.round(frac * denominator);
            const divisor = gcd(numerator, denominator);
            return `${Math.floor(value)} ${numerator/divisor}/${denominator/divisor}`;
        }
    },

    // Validação de dados
    validators: {
        isNumber: (v) => !isNaN(parseFloat(v)) && isFinite(v),
        isInteger: (v) => Number.isInteger(parseFloat(v)),
        isPositive: (v) => parseFloat(v) > 0,
        isEmail: (v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v),
        isPhone: (v) => /^\(?[1-9]{2}\)? ?[0-9]{4,5}-?[0-9]{4}$/.test(v),
        isDate: (v) => !isNaN(Date.parse(v)),
        isURL: (v) => /^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$/.test(v),
        isCPF: (v) => {
            const cpf = v.replace(/[^\d]/g, '');
            if (cpf.length !== 11) return false;
            let sum = 0, rest;
            for (let i = 1; i <= 9; i++) sum += parseInt(cpf[i-1]) * (11 - i);
            rest = (sum * 10) % 11;
            if (rest === 10 || rest === 11) rest = 0;
            if (rest !== parseInt(cpf[9])) return false;
            sum = 0;
            for (let i = 1; i <= 10; i++) sum += parseInt(cpf[i-1]) * (12 - i);
            rest = (sum * 10) % 11;
            if (rest === 10 || rest === 11) rest = 0;
            return rest === parseInt(cpf[10]);
        },
        isCNPJ: (v) => {
            const cnpj = v.replace(/[^\d]/g, '');
            if (cnpj.length !== 14) return false;
            let sum = 0, size = 12, weight = 5;
            for (let i = 0; i < size; i++) {
                sum += parseInt(cnpj[i]) * weight;
                weight = weight === 2 ? 9 : weight - 1;
            }
            let rest = sum % 11;
            if (rest === 10 || rest === 11) rest = 0;
            if (rest !== parseInt(cnpj[12])) return false;
            sum = 0; size = 13; weight = 6;
            for (let i = 0; i < size; i++) {
                sum += parseInt(cnpj[i]) * weight;
                weight = weight === 2 ? 9 : weight - 1;
            }
            rest = sum % 11;
            return (rest === 10 || rest === 11) ? 0 === parseInt(cnpj[13]) : rest === parseInt(cnpj[13]);
        }
    },

    // Geradores de dados
    generators: {
        randomNumber: (min = 0, max = 100) => Math.floor(Math.random() * (max - min + 1)) + min,
        randomDate: (start = new Date(2023, 0, 1), end = new Date()) => {
            const date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
            return date.toLocaleDateString('pt-BR');
        },
        randomText: (length = 10) => {
            const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
            let result = '';
            for (let i = 0; i < length; i++) {
                result += chars.charAt(Math.floor(Math.random() * chars.length));
            }
            return result;
        },
        randomEmail: () => {
            const domains = ['gmail.com', 'outlook.com', 'yahoo.com', 'empresa.com'];
            return `${ExcelDatabase.generators.randomText(8)}@${domains[Math.floor(Math.random() * domains.length)]}`;
        },
        randomPhone: () => {
            const ddd = Math.floor(Math.random() * 90) + 10;
            const num = Math.floor(Math.random() * 90000000) + 10000000;
            return `(${ddd}) ${String(num).slice(0,4)}-${String(num).slice(4)}`;
        }
    },

    // Funções de planilha avançadas
    advanced: {
        // Análise de regressão linear
        linearRegression: (x, y) => {
            const n = x.length;
            const sumX = x.reduce((a, b) => a + b, 0);
            const sumY = y.reduce((a, b) => a + b, 0);
            const sumXY = x.reduce((a, b, i) => a + b * y[i], 0);
            const sumX2 = x.reduce((a, b) => a + b * b, 0);
            const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
            const intercept = (sumY - slope * sumX) / n;
            return { slope, intercept, predict: (xVal) => slope * xVal + intercept };
        },
        
        // Correlação de Pearson
        correlation: (x, y) => {
            const n = x.length;
            const sumX = x.reduce((a, b) => a + b, 0);
            const sumY = y.reduce((a, b) => a + b, 0);
            const sumXY = x.reduce((a, b, i) => a + b * y[i], 0);
            const sumX2 = x.reduce((a, b) => a + b * b, 0);
            const sumY2 = y.reduce((a, b) => a + b * b, 0);
            const corr = (n * sumXY - sumX * sumY) / 
                Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
            return corr;
        },

        // Média móvel
        movingAverage: (data, windowSize) => {
            const result = [];
            for (let i = 0; i <= data.length - windowSize; i++) {
                const window = data.slice(i, i + windowSize);
                result.push(window.reduce((a, b) => a + b, 0) / windowSize);
            }
            return result;
        },

        // Suavização exponencial
        exponentialSmoothing: (data, alpha) => {
            const result = [data[0]];
            for (let i = 1; i < data.length; i++) {
                result.push(alpha * data[i] + (1 - alpha) * result[i - 1]);
            }
            return result;
        }
    },

    getTemplate(name) {
        return this.templates[name] || null;
    },

    listTemplates() {
        return Object.keys(this.templates);
    }
};

// Exporta para uso global
window.ExcelDatabase = ExcelDatabase;