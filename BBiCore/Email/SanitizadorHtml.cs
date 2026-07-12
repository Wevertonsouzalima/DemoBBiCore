// =============================================================================
//  SanitizadorHtml.cs  —  Limpeza do HTML do e-mail (BBiCore.Email)
// -----------------------------------------------------------------------------
//  AUTOMÁTICO E NÃO NEGOCIÁVEL: roda sempre ao salvar um corpo escrito no modo
//  avançado. Não há parâmetro para desligar e não é decisão do dev — se fosse
//  opcional, faltaria em algum sistema.
//
//  FRONTEIRA: passa TUDO que é aparência e navegação; barra APENAS execução de
//  código.
//
//   PASSA  · tabelas completas, listas, formatação, parágrafos
//          · imagens (cid:, data:, http/https, GIF), imagem de fundo
//          · links http/https/mailto/tel — INCLUSIVE estilizados como botão
//            ("clique aqui para acessar o Registro" continua funcionando)
//          · style inline, bgcolor, background, width, height, align, border...
//
//   BARRA  · <script>, <iframe>, <object>, <embed>, <form>, <base>
//          · atributos de evento (onclick, onerror, onload — qualquer on*)
//          · href/src com esquema javascript: ou vbscript:
//          · dentro de style: expression() e url(javascript:)
//
//  TRANSPARÊNCIA: o que for removido é DEVOLVIDO na lista de remoções. A tela
//  avisa o usuário (ele precisa entender por que o que colou mudou) e o log
//  registra — assim, se ele reclamar depois, sabemos de onde veio.
//
//  Depende de HtmlAgilityPack (MIT). Análise por árvore, não por expressão
//  regular: HTML malformado é justamente o que engana filtros baseados em regex.
// =============================================================================

using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace BBiCore.Email
{
    // #region Resultado  ->  destino sugerido: Models/

    /// <summary>Um item retirado do HTML pela sanitização.</summary>
    /// <param name="Elemento">Tag ou atributo removido (ex.: "script", "onclick").</param>
    /// <param name="Motivo">Explicação em português, para exibir ao usuário e gravar no log.</param>
    public sealed record RemocaoHtml(string Elemento, string Motivo);

    /// <summary>Resultado da sanitização: o HTML limpo e o que precisou ser retirado.</summary>
    /// <param name="Html">HTML já limpo — é o que deve ser gravado.</param>
    /// <param name="Remocoes">Itens removidos. Vazio quando nada foi alterado.</param>
    public sealed record ResultadoSanitizacao(string Html, IReadOnlyList<RemocaoHtml> Remocoes)
    {
        /// <summary>Indica que o HTML gravado é DIFERENTE do que o usuário escreveu.</summary>
        public bool Alterado => Remocoes.Count > 0;

        /// <summary>Resumo das remoções, pronto para a mensagem de tela e para o log.</summary>
        /// <returns>Texto único descrevendo o que saiu.</returns>
        public string Resumo()
        {
            if (Remocoes.Count == 0)
                return string.Empty;

            IEnumerable<string> itens = Remocoes
                .Select(r => r.Elemento)
                .Distinct(StringComparer.OrdinalIgnoreCase);

            return string.Join(", ", itens);
        }
    }

    // #endregion

    // #region Sanitizador  ->  destino sugerido: Services/

    /// <summary>Remove do HTML tudo o que pode executar código, preservando toda a parte visual.</summary>
    public static class SanitizadorHtml
    {
        /// <summary>Tags que não têm uso legítimo em e-mail e são removidas inteiras (com o conteúdo).</summary>
        private static readonly HashSet<string> TagsProibidas = new(StringComparer.OrdinalIgnoreCase)
        {
            "script", "iframe", "object", "embed", "applet", "form", "input",
            "button", "select", "textarea", "base", "link", "meta", "frame", "frameset"
        };

        /// <summary>Atributos que carregam URL e por isso precisam ter o esquema conferido.</summary>
        private static readonly HashSet<string> AtributosDeUrl = new(StringComparer.OrdinalIgnoreCase)
        {
            "href", "src", "background", "action", "formaction", "poster"
        };

        /// <summary>Esquemas de URL que executam código e nunca são aceitos.</summary>
        private static readonly string[] EsquemasProibidos = ["javascript:", "vbscript:", "data:text/html"];

        /// <summary>Construções de CSS que executam código (IE antigo e url(javascript:)).</summary>
        private static readonly Regex CssPerigoso = new(
            @"expression\s*\(|url\s*\(\s*['""]?\s*(javascript|vbscript)\s*:",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Limpa o HTML: remove execução de código e devolve o que foi retirado. O HTML devolvido é o
        /// que deve ser gravado.
        /// </summary>
        /// <param name="html">HTML escrito pelo usuário (modo avançado).</param>
        /// <returns>HTML limpo e a lista de remoções.</returns>
        public static ResultadoSanitizacao Sanitizar(string? html)
        {
            if (string.IsNullOrWhiteSpace(html))
                return new ResultadoSanitizacao(string.Empty, []);

            HtmlDocument documento = new();
            documento.LoadHtml(html);

            List<RemocaoHtml> remocoes = [];

            RemoverTagsProibidas(documento, remocoes);
            LimparAtributos(documento, remocoes);
            RemoverComentarios(documento);

            return new ResultadoSanitizacao(documento.DocumentNode.OuterHtml, remocoes);
        }

        /// <summary>Remove as tags que executam código ou não fazem sentido em e-mail.</summary>
        /// <param name="documento">Documento HTML.</param>
        /// <param name="remocoes">Lista que acumula o que foi retirado.</param>
        private static void RemoverTagsProibidas(HtmlDocument documento, List<RemocaoHtml> remocoes)
        {
            List<HtmlNode> alvos = documento.DocumentNode
                .Descendants()
                .Where(n => n.NodeType == HtmlNodeType.Element && TagsProibidas.Contains(n.Name))
                .ToList();

            foreach (HtmlNode no in alvos)
            {
                remocoes.Add(new RemocaoHtml(
                    $"<{no.Name}>",
                    $"a tag <{no.Name}> não é permitida em e-mail (pode executar código)."));

                no.Remove();
            }
        }

        /// <summary>Remove atributos de evento, URLs com esquema perigoso e CSS que executa código.</summary>
        /// <param name="documento">Documento HTML.</param>
        /// <param name="remocoes">Lista que acumula o que foi retirado.</param>
        private static void LimparAtributos(HtmlDocument documento, List<RemocaoHtml> remocoes)
        {
            List<HtmlNode> elementos = documento.DocumentNode
                .Descendants()
                .Where(n => n.NodeType == HtmlNodeType.Element && n.HasAttributes)
                .ToList();

            foreach (HtmlNode elemento in elementos)
            {
                List<HtmlAttribute> atributos = [.. elemento.Attributes];

                foreach (HtmlAttribute atributo in atributos)
                {
                    // Eventos: onclick, onerror, onload... nada disso é visual.
                    if (atributo.Name.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                    {
                        remocoes.Add(new RemocaoHtml(
                            atributo.Name,
                            $"o atributo {atributo.Name} executa código e foi removido."));

                        elemento.Attributes.Remove(atributo);
                        continue;
                    }

                    // URLs: o link fica; o esquema que executa código sai.
                    if (AtributosDeUrl.Contains(atributo.Name) && EsquemaProibido(atributo.Value))
                    {
                        remocoes.Add(new RemocaoHtml(
                            $"{atributo.Name}=javascript:",
                            $"o endereço em {atributo.Name} executava código (javascript:) e foi removido."));

                        elemento.Attributes.Remove(atributo);
                        continue;
                    }

                    // CSS: mantém o estilo legítimo, tira só a DECLARAÇÃO que executa código
                    // (remover apenas o trecho deixaria resíduo do tipo "background:alert(1))").
                    if (atributo.Name.Equals("style", StringComparison.OrdinalIgnoreCase)
                        && CssPerigoso.IsMatch(atributo.Value ?? string.Empty))
                    {
                        string limpo = LimparEstilo(atributo.Value ?? string.Empty);

                        remocoes.Add(new RemocaoHtml(
                            "style",
                            "o estilo continha uma instrução que executa código (expression/javascript) e ela foi removida."));

                        if (string.IsNullOrWhiteSpace(limpo))
                            elemento.Attributes.Remove(atributo);
                        else
                            atributo.Value = limpo;
                    }
                }
            }
        }

        /// <summary>
        /// Limpa o atributo style descartando as DECLARAÇÕES que executam código e preservando as
        /// demais (ex.: de "width:100%;background:expression(x)" sobra "width:100%").
        /// </summary>
        /// <param name="estilo">Conteúdo do atributo style.</param>
        /// <returns>Estilo sem as declarações perigosas.</returns>
        private static string LimparEstilo(string estilo)
        {
            List<string> mantidas = [];

            foreach (string declaracao in estilo.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                if (CssPerigoso.IsMatch(declaracao))
                    continue;

                string limpa = declaracao.Trim();

                if (!string.IsNullOrEmpty(limpa))
                    mantidas.Add(limpa);
            }

            return mantidas.Count == 0 ? string.Empty : string.Join("; ", mantidas);
        }

        /// <summary>Remove comentários HTML (podem esconder conteúdo condicional executável).</summary>
        /// <param name="documento">Documento HTML.</param>
        private static void RemoverComentarios(HtmlDocument documento)
        {
            List<HtmlNode> comentarios = documento.DocumentNode
                .Descendants()
                .Where(n => n.NodeType == HtmlNodeType.Comment)
                .ToList();

            foreach (HtmlNode comentario in comentarios)
                comentario.Remove();
        }

        /// <summary>Confere se o valor de uma URL usa um esquema que executa código.</summary>
        /// <param name="valor">Valor do atributo.</param>
        /// <returns>Verdadeiro quando o esquema é proibido.</returns>
        private static bool EsquemaProibido(string? valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return false;

            // Espaços, quebras e maiúsculas são usados para disfarçar o esquema.
            string normalizado = valor
                .Replace(" ", string.Empty)
                .Replace("\t", string.Empty)
                .Replace("\n", string.Empty)
                .Replace("\r", string.Empty)
                .ToLowerInvariant();

            foreach (string esquema in EsquemasProibidos)
                if (normalizado.StartsWith(esquema, StringComparison.Ordinal))
                    return true;

            return false;
        }
    }

    // #endregion
}
