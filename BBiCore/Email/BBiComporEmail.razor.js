// =============================================================================
//  bbiComporEmail.js  —  Interop mínimo do editor de e-mail (RCL BBiCore)
//  Servido como _content/BBiCore/bbiComporEmail.js. Faz só o que o Blazor não
//  faz sozinho: inserir texto na posição do cursor de um <textarea> e devolver
//  o controle ao @bind via evento 'input'. Clipboard fica com o teclado
//  (Ctrl+C/V/X), de propósito — nada de Clipboard API aqui.
// =============================================================================

/**
 * Insere um texto na posição atual do cursor de um textarea e dispara 'input'
 * para o Blazor capturar o novo valor via binding.
 * @param {string} id  Id do elemento textarea.
 * @param {string} texto  Texto a inserir (ex.: "{{Nome}}").
 */
export function inserirNoCursor(id, texto) {
    const el = document.getElementById(id);
    if (!el) return;

    const ini = el.selectionStart ?? el.value.length;
    const fim = el.selectionEnd ?? el.value.length;

    el.value = el.value.substring(0, ini) + texto + el.value.substring(fim);

    const pos = ini + texto.length;
    el.selectionStart = el.selectionEnd = pos;

    // Avisa o Blazor (@bind / @oninput) que o valor mudou.
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.focus();
}
