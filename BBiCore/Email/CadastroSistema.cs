// =============================================================================
//  CadastroSistema.cs  —  Acesso ao cadastro do aplicativo (BBiCore.Email)
// -----------------------------------------------------------------------------
//  ESTE ARQUIVO É O PONTO DE INTEGRAÇÃO COM O BANCO CENTRALIZADOR.
//
//  A aplicação só se IDENTIFICA (OpcoesEmail.NomeSistema, vindo do appsettings:
//  "meu nome é X"). A partir daí é a biblioteca que busca o cadastro do app —
//  conta de e-mail, usuário e senha — e descriptografa a senha. Nenhuma
//  aplicação passa credencial, e nenhum dev implementa contrato para isso.
//
//  PARA IMPLEMENTAR (marcado com "TODO" abaixo):
//    · ObterCredenciaisAsync -> ler o cadastro no centralizador pelo nome do
//      sistema e devolver as credenciais JÁ descriptografadas.
//
//  Enquanto o acesso não estiver ligado, a busca cai no que estiver preenchido
//  em OpcoesEmail (usuário/senha), o que mantém o desenvolvimento local e o
//  projeto de demonstração funcionando fora da rede corporativa.
// =============================================================================

namespace BBiCore.Email
{
    /// <summary>
    /// Busca, no banco centralizador, os dados de e-mail cadastrados para a aplicação em execução.
    /// É usado pelos transportes: eles nunca recebem credencial de fora.
    /// </summary>
    public sealed class CadastroSistema
    {
        /// <summary>Configurações da aplicação (traz o nome do sistema, usado na busca).</summary>
        private readonly OpcoesEmail _opcoes;

        /// <summary>Cria o acesso ao cadastro.</summary>
        /// <param name="opcoes">Configurações da aplicação.</param>
        public CadastroSistema(OpcoesEmail opcoes)
            => _opcoes = opcoes ?? throw new ArgumentNullException(nameof(opcoes));

        /// <summary>Obtém as credenciais de e-mail da aplicação em execução, já prontas para autenticar.</summary>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Credenciais com a senha em texto claro.</returns>
        public Task<CredenciaisEmail> ObterCredenciaisAsync(CancellationToken cancelamento = default)
        {
            // TODO: buscar no banco CENTRALIZADOR o cadastro cujo nome seja _opcoes.NomeSistema
            //       e devolver conta de e-mail, usuário e senha DESCRIPTOGRAFADA.
            //
            //       A conexão do centralizador e a rotina de descriptografia vivem na própria
            //       biblioteca (decisão tomada: a aplicação não participa disso).
            //
            //       Enquanto isso não existe, vale o que estiver no appsettings da aplicação:

            CredenciaisEmail credenciais = new(
                _opcoes.Usuario,
                _opcoes.Senha,
                _opcoes.EnderecoRemetente);

            return Task.FromResult(credenciais);
        }
    }
}
