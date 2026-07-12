// =============================================================================
//  CorpoNormal.cs  —  O corpo do MODO NORMAL (BBiCore.Email)
// -----------------------------------------------------------------------------
//  NO MODO NORMAL, A BIBLIOTECA É DONA DE 100% DO HTML. Ela emite a tabela, a
//  célula da imagem de cabeçalho (cid:), a célula de texto e a do rodapé. Dentro
//  da célula de texto só existe: texto ESCAPADO, quebra de linha, parágrafo e
//  {{marcador}}. Não há atributo externo, não há tag que a lib não tenha escrito.
//
//  POR ISSO O TEXTO PODE SER LIDO DE VOLTA. Não é "parsear HTML arbitrário" — é
//  ler um formato FECHADO que nós mesmos produzimos, com gramática conhecida e
//  finita. É isso que permite ao assistente REABRIR um template salvo em modo
//  normal, sem precisar de uma segunda coluna guardando o texto (o corpo continua
//  sendo a única fonte da verdade).
//
//  A garantia de que o HTML está neste formato vem de OrigemCriacao = Normal. Se
//  o template foi para o avançado, não se volta. E se um HTML marcado como normal
//  não corresponder ao formato (alguém editou no banco), a leitura AVISA em vez
//  de adivinhar — quem chama decide abrir no avançado.
// =============================================================================

using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace BBiCore.Email
{
    /// <summary>Monta e relê o corpo do modo normal (imagem de cabeçalho + texto + imagem de rodapé).</summary>
    public static partial class CorpoNormal
    {
        /// <summary>Marca que identifica a célula de texto na volta da leitura.</summary>
        private const string MarcaTexto = "bbi-texto";

        /// <summary>
        /// Gera o HTML do modo normal, em tabela Outlook-safe. As imagens entram por <c>cid:</c>, usando
        /// os ContentId dos vínculos de cabeçalho e rodapé (quando existirem).
        /// </summary>
        /// <param name="texto">Texto digitado pelo usuário (pode conter marcadores).</param>
        /// <param name="contentIdCabecalho">ContentId da imagem de cabeçalho; nulo quando não há.</param>
        /// <param name="contentIdRodape">ContentId da imagem de rodapé; nulo quando não há.</param>
        /// <returns>HTML do corpo.</returns>
        public static string Gerar(string? texto, string? contentIdCabecalho, string? contentIdRodape)
        {
            StringBuilder sb = new();

            sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"border-collapse:collapse\">");

            if (!string.IsNullOrWhiteSpace(contentIdCabecalho))
                sb.Append("<tr><td align=\"center\" style=\"padding:0\">")
                  .Append($"<img src=\"cid:{contentIdCabecalho}\" alt=\"\" style=\"display:block;max-width:100%;height:auto;border:0\">")
                  .Append("</td></tr>");

            sb.Append($"<tr><td class=\"{MarcaTexto}\" style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;color:#000000;padding:16px\">");
            sb.Append(TextoParaHtml(texto ?? string.Empty));
            sb.Append("</td></tr>");

            if (!string.IsNullOrWhiteSpace(contentIdRodape))
                sb.Append("<tr><td align=\"center\" style=\"padding:0\">")
                  .Append($"<img src=\"cid:{contentIdRodape}\" alt=\"\" style=\"display:block;max-width:100%;height:auto;border:0\">")
                  .Append("</td></tr>");

            sb.Append("</table>");
            return sb.ToString();
        }

        /// <summary>
        /// Lê de volta o texto do usuário a partir de um HTML gerado por <see cref="Gerar"/>. É a operação
        /// que permite reabrir o assistente sem guardar o texto numa segunda coluna.
        /// </summary>
        /// <param name="html">Corpo do template (deve ter sido gerado no modo normal).</param>
        /// <param name="texto">Texto recuperado, quando o formato é reconhecido.</param>
        /// <returns>Verdadeiro quando o HTML está no formato esperado e o texto foi recuperado.</returns>
        public static bool TentarLerTexto(string? html, out string texto)
        {
            texto = string.Empty;

            if (string.IsNullOrWhiteSpace(html))
                return true;   // corpo vazio é um normal válido (template recém-criado)

            HtmlDocument documento = new();
            documento.LoadHtml(html);

            HtmlNode? celula = documento.DocumentNode
                .Descendants("td")
                .FirstOrDefault(n => n.GetAttributeValue("class", string.Empty).Contains(MarcaTexto, StringComparison.OrdinalIgnoreCase));

            // Formato não reconhecido: alguém editou o HTML fora do assistente.
            // Quem chamou decide o que fazer (o certo é avisar e abrir no avançado).
            if (celula is null)
                return false;

            texto = HtmlParaTexto(celula.InnerHtml);
            return true;
        }

        /// <summary>Converte o texto do usuário em HTML seguro: escapa tags e transforma quebras em &lt;br&gt;. Marcadores passam intactos.</summary>
        /// <param name="texto">Texto simples.</param>
        /// <returns>Fragmento HTML.</returns>
        private static string TextoParaHtml(string texto)
        {
            string escapado = WebUtility.HtmlEncode(texto);

            escapado = escapado
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");

            return escapado.Replace("\n", "<br>\n");
        }

        /// <summary>Converte de volta o fragmento HTML da célula de texto no texto original do usuário.</summary>
        /// <param name="html">Conteúdo da célula de texto.</param>
        /// <returns>Texto como o usuário digitou.</returns>
        private static string HtmlParaTexto(string html)
        {
            // O caminho de volta é exatamente o inverso da ida: cada <br> (com a nova linha que o
            // acompanha) vira uma quebra, e as entidades voltam ao caractere original. Note que a
            // regex já consome o "\n" que segue o <br> — mexer nisso colapsa as linhas em branco.
            string texto = QuebraDeLinha().Replace(html, "\n");

            return WebUtility.HtmlDecode(texto).Trim();
        }

        /// <summary>Reconhece a quebra de linha gerada na ida (com ou sem a nova linha que a acompanha).</summary>
        /// <returns>Regex compilada.</returns>
        [GeneratedRegex(@"<br\s*/?>(\r?\n)?", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex QuebraDeLinha();
    }
}
