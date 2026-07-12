// =============================================================================
//  TransporteSmtp.cs  —  Entrega via SMTP usando MailKit (BBiCore.Email)
// -----------------------------------------------------------------------------
//  Depende do pacote MailKit (NuGet público). É o caminho moderno: o SmtpClient
//  da BCL está obsoleto e a própria Microsoft recomenda o MailKit no lugar.
//
//  Se preferir o Exchange (EWS), use o TransporteExchange — a troca é só o
//  registro na injeção de dependência; nada mais muda.
// =============================================================================

using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace BBiCore.Email
{
    /// <summary>Entrega e-mails por SMTP (relay interno ou servidor autenticado), via MailKit.</summary>
    public sealed class TransporteSmtp : ITransporteEmail
    {
        /// <summary>Configurações do sistema.</summary>
        private readonly OpcoesEmail _opcoes;

        /// <summary>Cadastro do sistema: é dele que saem as credenciais (a aplicação não as fornece).</summary>
        private readonly CadastroSistema _cadastro;

        /// <summary>Cria o transporte SMTP.</summary>
        /// <param name="opcoes">Configurações do sistema.</param>
        /// <param name="cadastro">Acesso ao cadastro do sistema (fonte das credenciais).</param>
        public TransporteSmtp(OpcoesEmail opcoes, CadastroSistema cadastro)
        {
            _opcoes = opcoes ?? throw new ArgumentNullException(nameof(opcoes));
            _cadastro = cadastro ?? throw new ArgumentNullException(nameof(cadastro));
        }

        /// <inheritdoc/>
        public async Task EntregarAsync(MensagemEmail mensagem, CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(mensagem);

            CredenciaisEmail credenciais = await ObterCredenciaisAsync(cancelamento);
            MimeMessage mime = MontadorMime.Montar(AjustarRemetente(mensagem, credenciais));

            using SmtpClient cliente = new()
            {
                Timeout = _opcoes.TimeoutSegundos * 1000
            };

            if (!_opcoes.ValidarCertificadoServidor)
                cliente.ServerCertificateValidationCallback = (remetente, certificado, cadeia, erros) => true;

            SecureSocketOptions seguranca = _opcoes.UsarSsl
                ? SecureSocketOptions.StartTlsWhenAvailable
                : SecureSocketOptions.None;

            await cliente.ConnectAsync(_opcoes.Host, _opcoes.Porta, seguranca, cancelamento);

            // Relay interno costuma não exigir autenticação: só autentica se houver usuário.
            if (!string.IsNullOrWhiteSpace(credenciais.Usuario))
                await cliente.AuthenticateAsync(credenciais.Usuario, credenciais.Senha, cancelamento);

            await cliente.SendAsync(mime, cancelamento);
            await cliente.DisconnectAsync(true, cancelamento);
        }

        /// <inheritdoc/>
        public async Task<ResultadoEnvio> TestarConexaoAsync(CancellationToken cancelamento = default)
        {
            try
            {
                CredenciaisEmail credenciais = await ObterCredenciaisAsync(cancelamento);

                using SmtpClient cliente = new()
                {
                    Timeout = _opcoes.TimeoutSegundos * 1000
                };

                if (!_opcoes.ValidarCertificadoServidor)
                    cliente.ServerCertificateValidationCallback = (remetente, certificado, cadeia, erros) => true;

                SecureSocketOptions seguranca = _opcoes.UsarSsl
                    ? SecureSocketOptions.StartTlsWhenAvailable
                    : SecureSocketOptions.None;

                await cliente.ConnectAsync(_opcoes.Host, _opcoes.Porta, seguranca, cancelamento);

                if (!string.IsNullOrWhiteSpace(credenciais.Usuario))
                    await cliente.AuthenticateAsync(credenciais.Usuario, credenciais.Senha, cancelamento);

                await cliente.DisconnectAsync(true, cancelamento);

                return new ResultadoEnvio(true, $"Conexão com {_opcoes.Host}:{_opcoes.Porta} bem-sucedida.");
            }
            catch (Exception ex)
            {
                return new ResultadoEnvio(false, $"Falha na conexão: {ex.Message}");
            }
        }

        /// <summary>Obtém as credenciais do cadastro do sistema.</summary>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Credenciais prontas.</returns>
        private async Task<CredenciaisEmail> ObterCredenciaisAsync(CancellationToken cancelamento)
            => await _cadastro.ObterCredenciaisAsync(cancelamento);

        /// <summary>Troca o remetente da mensagem pela conta das credenciais, quando elas informam uma.</summary>
        /// <param name="mensagem">Mensagem pronta.</param>
        /// <param name="credenciais">Credenciais em uso.</param>
        /// <returns>A mensagem, com o remetente ajustado quando necessário.</returns>
        private static MensagemEmail AjustarRemetente(MensagemEmail mensagem, CredenciaisEmail credenciais)
        {
            if (string.IsNullOrWhiteSpace(credenciais.EnderecoRemetente))
                return mensagem;

            return mensagem with { Remetente = credenciais.EnderecoRemetente };
        }
    }
}
