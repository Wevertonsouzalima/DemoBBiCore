// =============================================================================
//  LogEmail.cs  —  Registro (log) dos envios de e-mail (BBiCore.Email)
// -----------------------------------------------------------------------------
//  ESTE ARQUIVO É O PONTO DE INTEGRAÇÃO COM O BANCO DE LOGS.
//
//  A estrutura do processo já está pronta e LIGADA ao fluxo de envio: o
//  ServicoEmail chama o registrador SEMPRE — no sucesso e na falha —, sem
//  depender de o dev de cada aplicação lembrar de nada. O que falta é apenas o
//  CORPO dos métodos (a gravação em si), que depende dos bancos corporativos.
//
//  PARA IMPLEMENTAR (marcados com "TODO" abaixo):
//    · RegistrarSucessoAsync  -> grava a linha de log do envio bem-sucedido.
//    · RegistrarFalhaAsync    -> grava a linha de log com o erro ocorrido.
//
//  OBSERVAÇÃO SOBRE RETENÇÃO (decisão já tomada): o registro do envio é
//  permanente e leve; o CORPO é o dado pesado e sensível, e deve ter prazo. O
//  expurgo previsto ANULA o corpo (e os nomes de anexo, se for o caso) e MANTÉM
//  a linha — assim não se perde a prova de que houve envio. Uma coluna de data
//  indexada é o que torna esse job viável.
//
//  OBSERVAÇÃO SOBRE DADO PESSOAL: o corpo resolvido contém o que os marcadores
//  trouxeram (nome, documento, valores). Ele é o campo a criptografar na
//  gravação.
// =============================================================================

namespace BBiCore.Email
{
    // #region Registro  ->  destino sugerido: Models/

    /// <summary>
    /// Retrato de um envio, pronto para ser gravado no log. Traz tudo o que a biblioteca sabe no
    /// momento do envio — o que gravar (e o que criptografar ou descartar) é decidido na gravação.
    /// </summary>
    /// <param name="DataHora">Momento do envio.</param>
    /// <param name="Remetente">Conta que enviou.</param>
    /// <param name="Destinatarios">Destinatários já resolvidos.</param>
    /// <param name="Cc">Cópia já resolvida.</param>
    /// <param name="Cco">Cópia oculta já resolvida.</param>
    /// <param name="Assunto">Assunto já resolvido.</param>
    /// <param name="CorpoHtml">Corpo já resolvido. É o dado sensível: criptografe na gravação e expurgue no prazo.</param>
    /// <param name="NomesAnexos">Nomes dos anexos (apenas os nomes — o conteúdo NÃO entra no log).</param>
    /// <param name="Classificacao">Rótulo de sensibilidade do e-mail.</param>
    /// <param name="NomeTemplate">Nome do template de origem; nulo quando o envio foi avulso.</param>
    /// <param name="Sucesso">Se o envio foi concluído.</param>
    /// <param name="MensagemErro">Descrição do erro, quando houve falha.</param>
    public sealed record RegistroEnvioEmail(
        DateTime DataHora,
        string Remetente,
        IReadOnlyList<string> Destinatarios,
        IReadOnlyList<string> Cc,
        IReadOnlyList<string> Cco,
        string Assunto,
        string CorpoHtml,
        IReadOnlyList<string> NomesAnexos,
        ClassificacaoEmail Classificacao,
        string? NomeTemplate,
        bool Sucesso,
        string? MensagemErro)
    {
        /// <summary>Cria o registro a partir da mensagem enviada.</summary>
        /// <param name="mensagem">Mensagem entregue (ou tentada).</param>
        /// <param name="nomeTemplate">Template de origem; nulo no envio avulso.</param>
        /// <param name="sucesso">Se o envio deu certo.</param>
        /// <param name="mensagemErro">Erro, quando houve.</param>
        /// <returns>Registro pronto para gravação.</returns>
        public static RegistroEnvioEmail De(
            MensagemEmail mensagem,
            string? nomeTemplate,
            bool sucesso,
            string? mensagemErro)
        {
            ArgumentNullException.ThrowIfNull(mensagem);

            return new RegistroEnvioEmail(
                DateTime.Now,
                mensagem.Remetente,
                mensagem.Destinatarios,
                mensagem.Cc,
                mensagem.Cco,
                mensagem.Assunto,
                mensagem.CorpoHtml,
                [.. mensagem.Anexos.Select(a => a.NomeArquivo)],
                mensagem.Classificacao,
                nomeTemplate,
                sucesso,
                mensagemErro);
        }
    }

    // #endregion

    // #region Registrador  ->  destino sugerido: Services/

    /// <summary>
    /// Grava os envios no banco de logs. É acionado pelo <see cref="ServicoEmail"/> em TODO envio —
    /// não há caminho de envio que não passe por aqui, e nenhuma aplicação precisa registrar nada
    /// por conta própria.
    /// </summary>
    public sealed class RegistradorEnvioEmail
    {
        /// <summary>Registra um envio bem-sucedido.</summary>
        /// <param name="registro">Retrato do envio.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        public Task RegistrarSucessoAsync(RegistroEnvioEmail registro, CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(registro);

            // TODO: gravar a linha no banco de LOGS.
            //   · corpo (registro.CorpoHtml) -> criptografar antes de gravar;
            //   · anexos -> gravar apenas os nomes (registro.NomesAnexos);
            //   · gravar a data (registro.DataHora) em coluna INDEXADA, para viabilizar o expurgo.
            return Task.CompletedTask;
        }

        /// <summary>
        /// Registra que a sanitização ALTEROU o HTML salvo pelo usuário. É registrado sempre que algo
        /// é removido — assim, se o usuário reclamar depois ("meu e-mail está diferente do que colei"),
        /// sabemos exatamente o que saiu, quando e de qual template.
        /// </summary>
        /// <param name="nomeTemplate">Template que estava sendo salvo.</param>
        /// <param name="remocoes">O que foi removido do HTML.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        public Task RegistrarSanitizacaoAsync(
            string nomeTemplate,
            IReadOnlyList<RemocaoHtml> remocoes,
            CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(remocoes);

            // TODO: gravar no banco de LOGS que o HTML de 'nomeTemplate' foi alterado na gravação,
            //       com a lista de remoções (cada uma traz Elemento e Motivo).
            return Task.CompletedTask;
        }

        /// <summary>Registra uma tentativa de envio que FALHOU, com o erro ocorrido.</summary>
        /// <param name="registro">Retrato do envio (o e-mail que se tentou enviar).</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        public Task RegistrarFalhaAsync(RegistroEnvioEmail registro, CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(registro);

            // TODO: gravar a linha no banco de LOGS, marcando a falha.
            //   · registro.MensagemErro traz a causa;
            //   · o e-mail que se tentou enviar vai junto (mesmo tratamento do sucesso:
            //     corpo criptografado, só nomes de anexo).
            return Task.CompletedTask;
        }
    }

    // #endregion
}
