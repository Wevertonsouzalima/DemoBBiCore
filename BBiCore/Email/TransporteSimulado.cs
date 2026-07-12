// =============================================================================
//  TransporteSimulado.cs  —  Entrega simulada em disco (BBiCore.Email)
// -----------------------------------------------------------------------------
//  Em vez de transmitir, grava o e-mail montado como arquivo .eml numa pasta.
//  Serve para testar o fluxo inteiro — inclusive anexos e imagens embutidas —
//  sem depender de servidor. O .eml gerado abre no Outlook.
//
//  Registre este transporte em desenvolvimento/homologação; em produção, troque
//  pelo TransporteExchange ou pelo TransporteSmtp. Nada mais muda.
// =============================================================================

namespace BBiCore.Email
{
    /// <summary>Transporte de SIMULAÇÃO: grava o e-mail em disco (.eml) em vez de enviá-lo.</summary>
    public sealed class TransporteSimulado : ITransporteEmail
    {
        /// <summary>Configurações do sistema.</summary>
        private readonly OpcoesEmail _opcoes;

        /// <summary>Cria o transporte simulado.</summary>
        /// <param name="opcoes">Configurações do sistema (usa PastaSimulacao).</param>
        public TransporteSimulado(OpcoesEmail opcoes)
            => _opcoes = opcoes ?? throw new ArgumentNullException(nameof(opcoes));

        /// <summary>Pasta onde os arquivos são gravados.</summary>
        public string PastaDestino => string.IsNullOrWhiteSpace(_opcoes.PastaSimulacao)
            ? Path.Combine(Path.GetTempPath(), "bbi-emails")
            : _opcoes.PastaSimulacao;

        /// <inheritdoc/>
        public async Task EntregarAsync(MensagemEmail mensagem, CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(mensagem);

            Directory.CreateDirectory(PastaDestino);

            // Mesma montagem do envio real: o .eml reflete fielmente o que seria transmitido.
            byte[] eml = MontadorMime.GerarEml(mensagem);
            string caminho = Path.Combine(PastaDestino, $"{DateTime.Now:yyyyMMdd-HHmmss-fff}.eml");

            await File.WriteAllBytesAsync(caminho, eml, cancelamento);
        }

        /// <inheritdoc/>
        public Task<ResultadoEnvio> TestarConexaoAsync(CancellationToken cancelamento = default)
        {
            try
            {
                // Não há servidor: o "teste" é confirmar que a pasta de destino pode ser criada/escrita.
                Directory.CreateDirectory(PastaDestino);

                return Task.FromResult(new ResultadoEnvio(
                    true,
                    $"Transporte simulado pronto — os e-mails serão gravados em {PastaDestino}."));
            }
            catch (Exception ex)
            {
                return Task.FromResult(new ResultadoEnvio(false, $"Pasta de simulação inacessível: {ex.Message}"));
            }
        }
    }
}
