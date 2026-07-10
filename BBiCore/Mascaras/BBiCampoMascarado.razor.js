// Máscaras client-side. Formatam o input enquanto o usuário digita, mantendo o
// cursor ao fim (suficiente para digitação normal). O Blazor lê o valor no change.

function digitos(s) { return (s || "").replace(/\D/g, ""); }

function fmtCpf(s) {
    let d = digitos(s).slice(0, 11), r = "";
    for (let i = 0; i < d.length; i++) {
        if (i === 3 || i === 6) r += ".";
        else if (i === 9) r += "-";
        r += d[i];
    }
    return r;
}

function fmtCnpj(s) {
    let d = digitos(s).slice(0, 14), r = "";
    for (let i = 0; i < d.length; i++) {
        if (i === 2 || i === 5) r += ".";
        else if (i === 8) r += "/";
        else if (i === 12) r += "-";
        r += d[i];
    }
    return r;
}

function fmtDoc(s) { return digitos(s).length <= 11 ? fmtCpf(s) : fmtCnpj(s); }

function fmtTel(s) {
    let d = digitos(s).slice(0, 11), cel = d.length > 10, r = "";
    for (let i = 0; i < d.length; i++) {
        if (i === 0) r += "(";
        else if (i === 2) r += ") ";
        else if ((cel && i === 7) || (!cel && i === 6)) r += "-";
        r += d[i];
    }
    return r;
}

function fmtCep(s) {
    let d = digitos(s).slice(0, 8);
    return d.length <= 5 ? d : d.slice(0, 5) + "-" + d.slice(5);
}

function fmtMoeda(s) {
    let d = digitos(s);
    if (!d) return "";
    let v = parseInt(d, 10) / 100;
    return v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtData(s) {
    let d = digitos(s).slice(0, 8), r = "";
    for (let i = 0; i < d.length; i++) {
        if (i === 2 || i === 4) r += "/";
        r += d[i];
    }
    return r;
}

const fmts = { doc: fmtDoc, cpf: fmtCpf, cnpj: fmtCnpj, tel: fmtTel, cep: fmtCep, moeda: fmtMoeda, data: fmtData };

// Liga a máscara a um input. Idempotente.
export function ligar(el, tipo) {
    if (!el || el.dataset.bbiLigado) return;
    el.dataset.bbiLigado = "1";
    const fmt = fmts[tipo] || (x => x);
    el.addEventListener("input", () => {
        el.value = fmt(el.value);
        const n = el.value.length;
        el.setSelectionRange(n, n);
    });
}
