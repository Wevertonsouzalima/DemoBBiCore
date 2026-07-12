// =============================================================================
//  MontadorMime.cs  —  Montagem da mensagem MIME (BBiCore.Email)
// -----------------------------------------------------------------------------
//  Ponto ÚNICO de montagem MIME, usado por quem precisa de um MimeMessage:
//    · TransporteSmtp     (envia pelo MailKit)
//    · TransporteSimulado (grava o .eml em disco)
//    · GeradorRascunho    (gera o .eml do rascunho)
//
//  Assim o .eml gerado é EXATAMENTE a mensagem que seria enviada — mesmo corpo,
//  mesmas imagens embutidas, mesmos anexos, mesma classificação.
//
//  Usa MimeKit (vem junto com o pacote MailKit).
// =============================================================================

using MimeKit;

namespace BBiCore.Email
{
    /// <summary>Monta a mensagem MIME a partir de uma <see cref="MensagemEmail"/> já pronta.</summary>
    internal static class MontadorMime
    {
        /// <summary>Monta a mensagem MIME, com imagens embutidas (cid:) e anexos comuns.</summary>
        /// <param name="mensagem">Mensagem pronta (tudo resolvido, anexos em memória).</param>
        /// <returns>Mensagem MIME equivalente.</returns>
        public static MimeMessage Montar(MensagemEmail mensagem)
        {
            ArgumentNullException.ThrowIfNull(mensagem);

            MimeMessage mime = new()
            {
                Subject = mensagem.Assunto
            };

            mime.From.Add(new MailboxAddress(
                mensagem.NomeRemetente ?? mensagem.Remetente,
                mensagem.Remetente));

            foreach (string destino in mensagem.Destinatarios)
                mime.To.Add(MailboxAddress.Parse(destino));

            foreach (string cc in mensagem.Cc)
                mime.Cc.Add(MailboxAddress.Parse(cc));

            foreach (string cco in mensagem.Cco)
                mime.Bcc.Add(MailboxAddress.Parse(cco));

            string? sensibilidade = MapearSensibilidade(mensagem.Classificacao);

            if (sensibilidade is not null)
                mime.Headers.Add("Sensitivity", sensibilidade);

            BodyBuilder corpo = new()
            {
                HtmlBody = mensagem.CorpoHtml
            };

            foreach (AnexoMensagem anexo in mensagem.Anexos)
            {
                if (anexo.EhInline)
                {
                    MimeEntity embutido = corpo.LinkedResources.Add(
                        anexo.NomeArquivo,
                        anexo.Conteudo,
                        ObterTipoConteudo(anexo.ContentType));

                    // O corpo referencia esta imagem por cid:{ContentId}.
                    embutido.ContentId = anexo.ContentId;
                }
                else
                {
                    corpo.Attachments.Add(
                        anexo.NomeArquivo,
                        anexo.Conteudo,
                        ObterTipoConteudo(anexo.ContentType));
                }
            }

            mime.Body = corpo.ToMessageBody();
            return mime;
        }

        /// <summary>Serializa a mensagem MIME como bytes de um arquivo .eml.</summary>
        /// <param name="mensagem">Mensagem pronta.</param>
        /// <returns>Bytes do .eml (abre no Outlook).</returns>
        public static byte[] GerarEml(MensagemEmail mensagem)
        {
            using MimeMessage mime = Montar(mensagem);
            using MemoryStream ms = new();

            mime.WriteTo(ms);
            return ms.ToArray();
        }

        /// <summary>Converte o content-type textual no tipo do MimeKit; usa binário genérico quando não informado.</summary>
        /// <param name="contentType">Content-type (ex.: "image/png").</param>
        /// <returns>Tipo de conteúdo correspondente.</returns>
        public static ContentType ObterTipoConteudo(string? contentType)
        {
            if (string.IsNullOrWhiteSpace(contentType))
                return new ContentType("application", "octet-stream");

            if (ContentType.TryParse(contentType, out ContentType? tipo) && tipo is not null)
                return tipo;

            return new ContentType("application", "octet-stream");
        }

        /// <summary>Mapeia a classificação para o cabeçalho Sensitivity do MIME.</summary>
        /// <param name="classificacao">Classificação do e-mail.</param>
        /// <returns>Valor do cabeçalho, ou nulo quando não se aplica.</returns>
        private static string? MapearSensibilidade(ClassificacaoEmail classificacao)
        {
            switch (classificacao)
            {
                case ClassificacaoEmail.Confidencial:
                    return "Company-Confidential";
                case ClassificacaoEmail.Interno:
                case ClassificacaoEmail.Publico:
                default:
                    return null;
            }
        }
    }
}
