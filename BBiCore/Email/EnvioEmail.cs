// =============================================================================
//  EnvioEmail.cs  —  Núcleo do envio de e-mail (BBiCore.Email)
// -----------------------------------------------------------------------------
//  ARQUITETURA POR COMPOSIÇÃO (não herança):
//
//      IServicoEmail  (o que o DEV usa)
//           |
//           +--> ServicoEmail  ── faz o que é COMUM a todo envio:
//           |        · resolve o template (quando o e-mail vem de um template)
//           |        · materializa anexos (lê caminho, confina à pasta base,
//           |          aplica a política de falha, agenda a exclusão)
//           |        · aplica o redirecionamento de homologação
//           |        · repete em falha transitória (retry com espera crescente)
//           |
//           +--> ITransporteEmail  ── só ENTREGA a mensagem pronta:
//                    · TransporteSmtp     (MailKit)
//                    · TransporteExchange (EWS)
//                    · TransporteSimulado (grava .eml em disco)
//
//  DUAS SEMÂNTICAS, DE PROPÓSITO:
//   · E-mail AVULSO (EmailAvulso)  -> texto LITERAL. Um "{{campo}}" digitado
//     pelo dev vai para o destinatário exatamente assim. Não há resolução.
//   · E-mail POR TEMPLATE          -> aí sim os marcadores {{campo}} são
//     resolvidos com o objeto de dados. É o caminho que o componente usa.
//
//  >>> NOTA PARA REORGANIZAÇÃO: cada tipo em sua #region, nomeada pelo destino.
// =============================================================================

using BBiCore.Templates;

namespace BBiCore.Email
{
    // #region Mensagem e anexos  ->  destino sugerido: Models/

    /// <summary>Anexo já materializado em memória (bytes lidos e política de falha aplicada).</summary>
    /// <param name="NomeArquivo">Nome final do anexo.</param>
    /// <param name="ContentType">Content-type, quando conhecido.</param>
    /// <param name="Conteudo">Bytes do anexo.</param>
    /// <param name="ContentId">ContentId quando é imagem embutida (cid:); nulo para anexo comum.</param>
    public sealed record AnexoMensagem(string NomeArquivo, string? ContentType, byte[] Conteudo, string? ContentId)
    {
        /// <summary>Indica que o anexo é uma imagem embutida no corpo (referenciada por cid:).</summary>
        public bool EhInline => ContentId is not null;
    }

    /// <summary>Mensagem pronta para o transporte: tudo já resolvido, anexos já em memória.</summary>
    /// <param name="Remetente">Endereço do remetente.</param>
    /// <param name="NomeRemetente">Nome de exibição do remetente.</param>
    /// <param name="Destinatarios">Destinatários.</param>
    /// <param name="Cc">Cópia.</param>
    /// <param name="Cco">Cópia oculta.</param>
    /// <param name="Assunto">Assunto final.</param>
    /// <param name="CorpoHtml">Corpo em HTML.</param>
    /// <param name="Classificacao">Rótulo de sensibilidade.</param>
    /// <param name="Anexos">Anexos materializados.</param>
    public sealed record MensagemEmail(
        string Remetente,
        string? NomeRemetente,
        IReadOnlyList<string> Destinatarios,
        IReadOnlyList<string> Cc,
        IReadOnlyList<string> Cco,
        string Assunto,
        string CorpoHtml,
        ClassificacaoEmail Classificacao,
        IReadOnlyList<AnexoMensagem> Anexos);

    /// <summary>Resultado de uma tentativa de envio.</summary>
    /// <param name="Sucesso">Se o e-mail foi enviado.</param>
    /// <param name="Mensagem">Detalhe do erro (quando não houve sucesso) ou confirmação.</param>
    public sealed record ResultadoEnvio(bool Sucesso, string? Mensagem);

    // #endregion

    // #region E-mail avulso (uso programático pelo dev)  ->  destino sugerido: Models/

    /// <summary>
    /// Anexo de um e-mail avulso. Use <see cref="DeBytes"/> quando o conteúdo já estiver em memória,
    /// ou <see cref="DeCaminho"/> para um arquivo em disco (confinado à pasta base configurada).
    /// </summary>
    public sealed class AnexoAvulso
    {
        /// <summary>Nome do arquivo apresentado ao destinatário.</summary>
        public string NomeArquivo { get; set; } = string.Empty;

        /// <summary>Content-type do anexo (opcional).</summary>
        public string? ContentType { get; set; }

        /// <summary>Bytes do anexo; nulo quando vem de um caminho.</summary>
        public byte[]? Conteudo { get; set; }

        /// <summary>Caminho do arquivo em disco; nulo quando o conteúdo já está em memória.</summary>
        public string? Caminho { get; set; }

        /// <summary>Se o arquivo de origem deve ser excluído após o envio bem-sucedido.</summary>
        public bool ExcluirAposEnviar { get; set; }

        /// <summary>ContentId, quando o anexo é uma imagem embutida referenciada por cid: no corpo.</summary>
        public string? ContentId { get; set; }

        /// <summary>Cria um anexo a partir de bytes em memória.</summary>
        /// <param name="nomeArquivo">Nome do arquivo.</param>
        /// <param name="conteudo">Bytes do anexo.</param>
        /// <param name="contentType">Content-type (opcional).</param>
        /// <returns>Anexo pronto.</returns>
        public static AnexoAvulso DeBytes(string nomeArquivo, byte[] conteudo, string? contentType = null)
            => new() { NomeArquivo = nomeArquivo, Conteudo = conteudo, ContentType = contentType };

        /// <summary>Cria um anexo a partir de um arquivo em disco.</summary>
        /// <param name="caminho">Caminho do arquivo (relativo à pasta base de anexos).</param>
        /// <param name="excluirAposEnviar">Se o arquivo deve ser excluído após o envio.</param>
        /// <returns>Anexo pronto.</returns>
        public static AnexoAvulso DeCaminho(string caminho, bool excluirAposEnviar = false)
            => new()
            {
                Caminho = caminho,
                NomeArquivo = Path.GetFileName(caminho),
                ExcluirAposEnviar = excluirAposEnviar
            };
    }

    /// <summary>
    /// E-mail montado pelo próprio dev, no código, sem template. O conteúdo é LITERAL: se você escrever
    /// "{{algo}}" no assunto ou no corpo, isso chega ao destinatário exatamente assim — marcadores só
    /// têm efeito no envio por template.
    /// </summary>
    public sealed class EmailAvulso
    {
        /// <summary>Destinatários.</summary>
        public IList<string> Para { get; } = [];

        /// <summary>Cópia.</summary>
        public IList<string> Cc { get; } = [];

        /// <summary>Cópia oculta.</summary>
        public IList<string> Cco { get; } = [];

        /// <summary>Assunto (texto literal).</summary>
        public string Assunto { get; set; } = string.Empty;

        /// <summary>Corpo em HTML (texto literal). Para texto simples, use <see cref="DefinirCorpoTexto"/>.</summary>
        public string CorpoHtml { get; set; } = string.Empty;

        /// <summary>Rótulo de sensibilidade da mensagem.</summary>
        public ClassificacaoEmail Classificacao { get; set; } = ClassificacaoEmail.Interno;

        /// <summary>Anexos do e-mail.</summary>
        public IList<AnexoAvulso> Anexos { get; } = [];

        /// <summary>Política quando um anexo falhar (ler arquivo inexistente, por exemplo).</summary>
        public AcaoFalhaAnexo AcaoNaFalhaDeAnexo { get; set; } = AcaoFalhaAnexo.FalharEnvio;

        /// <summary>Remetente alternativo; nulo usa o configurado em <see cref="OpcoesEmail"/>.</summary>
        public string? Remetente { get; set; }

        /// <summary>Define o corpo a partir de texto simples, convertendo para HTML seguro (escapa tags, quebras viram &lt;br&gt;).</summary>
        /// <param name="texto">Texto simples.</param>
        public void DefinirCorpoTexto(string texto)
        {
            string escapado = System.Net.WebUtility.HtmlEncode(texto ?? string.Empty);
            escapado = escapado.Replace("\r\n", "\n").Replace("\r", "\n");
            CorpoHtml = escapado.Replace("\n", "<br>\n");
        }
    }

    // #endregion

    // #region Configuração e credenciais  ->  destino sugerido: Models/

    /// <summary>Configurações de e-mail de um sistema. Uma instância por aplicação.</summary>
    public sealed class OpcoesEmail
    {
        /// <summary>
        /// Nome com que a aplicação se identifica ("meu nome é X", no appsettings). É por ele que a
        /// biblioteca acha o cadastro do app no banco centralizador e obtém as credenciais de envio.
        /// </summary>
        public string NomeSistema { get; set; } = string.Empty;

        /// <summary>Endereço do remetente (o "de").</summary>
        public string EnderecoRemetente { get; set; } = string.Empty;

        /// <summary>Nome de exibição do remetente (opcional).</summary>
        public string? NomeExibicao { get; set; }

        /// <summary>Usuário de autenticação. Usado apenas enquanto a busca no cadastro (CadastroSistema) não estiver ligada.</summary>
        public string Usuario { get; set; } = string.Empty;

        /// <summary>Senha em texto claro. Usada apenas enquanto a busca no cadastro (CadastroSistema) não estiver ligada.</summary>
        public string Senha { get; set; } = string.Empty;

        /// <summary>Servidor SMTP (transporte SMTP).</summary>
        public string Host { get; set; } = string.Empty;

        /// <summary>Porta do servidor SMTP (587 para STARTTLS, 25 para relay interno).</summary>
        public int Porta { get; set; } = 587;

        /// <summary>Se a conexão SMTP usa TLS.</summary>
        public bool UsarSsl { get; set; } = true;

        /// <summary>Endereço do serviço EWS (transporte Exchange). Ex.: https://servidor/ews/exchange.asmx</summary>
        public string UrlEws { get; set; } = string.Empty;

        /// <summary>Versão do Exchange usada pelo EWS (ex.: "Exchange2010_SP2").</summary>
        public string VersaoExchange { get; set; } = "Exchange2010_SP2";

        /// <summary>Se guarda o e-mail na pasta de Itens Enviados da conta (transporte Exchange).</summary>
        public bool SalvarCopiaEmEnviados { get; set; } = true;

        /// <summary>Se valida o certificado do servidor. Desligar aceita QUALQUER certificado — use só em ambiente interno controlado.</summary>
        public bool ValidarCertificadoServidor { get; set; } = true;

        /// <summary>
        /// Trava de homologação: quando preenchida, TODOS os e-mails vão só para estes endereços
        /// (Cc e Cco são descartados e o assunto ganha um prefixo). Deixe vazia em produção.
        /// </summary>
        public IList<string> RedirecionarPara { get; } = [];

        /// <summary>Prefixo aplicado ao assunto quando o redirecionamento está ativo.</summary>
        public string PrefixoAssuntoTeste { get; set; } = "[TESTE] ";

        /// <summary>Pasta base dos anexos lidos de disco. Confina o acesso (barreira contra path traversal).</summary>
        public string? PastaBaseAnexos { get; set; }

        /// <summary>Pasta onde o transporte simulado grava os .eml (padrão: subpasta temporária).</summary>
        public string? PastaSimulacao { get; set; }

        /// <summary>Tempo limite do envio, em segundos.</summary>
        public int TimeoutSegundos { get; set; } = 60;

        /// <summary>Quantas vezes repetir o envio em falha transitória de rede (0 desliga).</summary>
        public int TentativasEmFalha { get; set; } = 2;

        /// <summary>Espera inicial entre tentativas, em milissegundos (dobra a cada tentativa).</summary>
        public int EsperaEntreTentativasMs { get; set; } = 500;
    }

    /// <summary>Credenciais de envio prontas para uso (senha já descriptografada).</summary>
    /// <param name="Usuario">Usuário de autenticação.</param>
    /// <param name="Senha">Senha em texto claro.</param>
    /// <param name="EnderecoRemetente">Conta de e-mail do sistema; nulo mantém o das opções.</param>
    public sealed record CredenciaisEmail(string Usuario, string Senha, string? EnderecoRemetente = null);

    // #endregion

    // #region Rascunho  ->  destino sugerido: Models/

    /// <summary>
    /// Para onde vai o rascunho gerado. É combinável: o dev pode pedir mais de um destino ao mesmo
    /// tempo (ex.: <c>Bytes | Arquivo</c>).
    /// </summary>
    [Flags]
    public enum DestinoRascunho
    {
        /// <summary>Nenhum destino (não gera nada).</summary>
        Nenhum = 0,

        /// <summary>Devolve os bytes do .eml (para o dev baixar pelo navegador ou tratar como quiser).</summary>
        Bytes = 1,

        /// <summary>Grava o .eml num arquivo em disco.</summary>
        Arquivo = 2,

        /// <summary>Salva na pasta Rascunhos da caixa postal (exige um transporte que suporte — hoje, o Exchange).</summary>
        CaixaPostal = 4
    }

    /// <summary>Opções da geração do rascunho.</summary>
    public sealed class OpcoesRascunho
    {
        /// <summary>Destinos desejados (combináveis). Padrão: apenas os bytes.</summary>
        public DestinoRascunho Destino { get; set; } = DestinoRascunho.Bytes;

        /// <summary>Pasta onde gravar o .eml quando <see cref="DestinoRascunho.Arquivo"/> estiver ligado. Nulo usa a pasta de simulação das opções.</summary>
        public string? PastaArquivo { get; set; }

        /// <summary>Nome do arquivo .eml (sem caminho). Nulo gera um nome a partir do assunto e do horário.</summary>
        public string? NomeArquivo { get; set; }
    }

    /// <summary>Resultado da geração de um rascunho.</summary>
    /// <param name="Sucesso">Se a geração ocorreu.</param>
    /// <param name="Mensagem">Detalhe do erro ou confirmação.</param>
    /// <param name="Conteudo">Bytes do .eml, quando o destino incluiu <see cref="DestinoRascunho.Bytes"/>.</param>
    /// <param name="NomeArquivo">Nome sugerido do arquivo .eml.</param>
    /// <param name="CaminhoArquivo">Caminho gravado, quando o destino incluiu <see cref="DestinoRascunho.Arquivo"/>.</param>
    /// <param name="SalvoNaCaixaPostal">Se foi salvo na pasta Rascunhos da caixa postal.</param>
    public sealed record ResultadoRascunho(
        bool Sucesso,
        string? Mensagem,
        byte[]? Conteudo = null,
        string? NomeArquivo = null,
        string? CaminhoArquivo = null,
        bool SalvoNaCaixaPostal = false);

    // #endregion

    // #region Transporte (contrato)  ->  destino sugerido: Services/

    /// <summary>
    /// Entrega de fato a mensagem. É a ÚNICA parte que conhece o protocolo — trocar Exchange por SMTP
    /// (ou pelo simulado) é trocar a implementação registrada, sem tocar em mais nada.
    /// </summary>
    public interface ITransporteEmail
    {
        /// <summary>Entrega a mensagem já pronta.</summary>
        /// <param name="mensagem">Mensagem com tudo resolvido e anexos em memória.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        Task EntregarAsync(MensagemEmail mensagem, CancellationToken cancelamento = default);

        /// <summary>
        /// Autentica no servidor e desconecta, SEM enviar nada. Serve para o usuário descobrir que a
        /// credencial ou o endereço estão errados na hora de configurar — e não no primeiro envio real.
        /// </summary>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Sucesso, ou a mensagem do erro encontrado.</returns>
        Task<ResultadoEnvio> TestarConexaoAsync(CancellationToken cancelamento = default);
    }

    /// <summary>
    /// Capacidade OPCIONAL de um transporte: salvar a mensagem na pasta Rascunhos da caixa postal
    /// (o e-mail aparece no Outlook da conta, pronto para revisão e disparo manual). Hoje só o
    /// transporte do Exchange implementa; pedir <see cref="DestinoRascunho.CaixaPostal"/> com um
    /// transporte que não implementa devolve erro explicativo, sem quebrar os outros destinos.
    /// </summary>
    public interface ITransporteRascunho
    {
        /// <summary>Salva a mensagem como rascunho na caixa postal da conta.</summary>
        /// <param name="mensagem">Mensagem pronta.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        Task SalvarRascunhoAsync(MensagemEmail mensagem, CancellationToken cancelamento = default);
    }

    /// <summary>Falha ao preparar um anexo, quando a política manda abortar o envio.</summary>
    public sealed class FalhaAnexoException : Exception
    {
        /// <summary>Cria a exceção com a mensagem informada.</summary>
        /// <param name="mensagem">Descrição da falha.</param>
        public FalhaAnexoException(string mensagem) : base(mensagem)
        {
        }
    }

    // #endregion

    // #region Serviço de e-mail (o que o dev usa)  ->  destino sugerido: Services/

    /// <summary>
    /// Fachada de e-mail da biblioteca. É o que o dev injeta para enviar — nas três formas:
    /// avulso (literal), por template salvo (resolve marcadores) e a partir de um template já em mãos
    /// (usado pelo componente).
    /// </summary>
    public interface IServicoEmail
    {
        /// <summary>Envia um e-mail montado no código. O conteúdo é LITERAL: marcadores não são resolvidos.</summary>
        /// <param name="email">E-mail avulso.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado do envio.</returns>
        Task<ResultadoEnvio> EnviarAsync(EmailAvulso email, CancellationToken cancelamento = default);

        /// <summary>
        /// Envia a partir de um template SALVO, resolvendo os marcadores com o objeto de dados.
        /// Só envia templates PUBLICADOS: um template em rascunho é recusado (é a trava que impede um
        /// template pela metade de sair para o cliente).
        /// </summary>
        /// <param name="nomeTemplate">Nome do template no repositório.</param>
        /// <param name="dados">Objeto que preenche os marcadores.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado do envio.</returns>
        Task<ResultadoEnvio> EnviarPorTemplateAsync(string nomeTemplate, object dados, CancellationToken cancelamento = default);

        /// <summary>
        /// Envia a partir de um template já em mãos, resolvendo os marcadores. Template em RASCUNHO é
        /// recusado — a menos que <paramref name="permitirRascunho"/> seja verdadeiro, o que o
        /// componente usa no "enviar teste" (testar antes de publicar é justamente o ponto).
        /// </summary>
        /// <param name="template">Template de e-mail.</param>
        /// <param name="dados">Objeto que preenche os marcadores.</param>
        /// <param name="permitirRascunho">Se permite enviar um template ainda não publicado (uso consciente: teste).</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado do envio.</returns>
        Task<ResultadoEnvio> EnviarTemplateAsync(ITemplateEmail template, object dados, bool permitirRascunho = false, CancellationToken cancelamento = default);

        /// <summary>
        /// Gera o RASCUNHO de um e-mail avulso em vez de enviá-lo: monta a MESMA mensagem que seria
        /// transmitida e a entrega como .eml (bytes e/ou arquivo) e/ou na caixa postal.
        /// </summary>
        /// <param name="email">E-mail avulso (conteúdo literal, como no envio).</param>
        /// <param name="opcoes">Destinos do rascunho; nulo usa apenas os bytes.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado com os bytes e/ou o caminho gravado.</returns>
        Task<ResultadoRascunho> GerarRascunhoAsync(EmailAvulso email, OpcoesRascunho? opcoes = null, CancellationToken cancelamento = default);

        /// <summary>
        /// Gera o RASCUNHO de um template, já RESOLVIDO com o objeto de dados — o .eml sai exatamente
        /// como o e-mail que seria enviado (sem marcadores), pronto para revisar e disparar.
        /// </summary>
        /// <param name="template">Template de e-mail.</param>
        /// <param name="dados">Objeto que preenche os marcadores.</param>
        /// <param name="opcoes">Destinos do rascunho; nulo usa apenas os bytes.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado com os bytes e/ou o caminho gravado.</returns>
        Task<ResultadoRascunho> GerarRascunhoTemplateAsync(ITemplateEmail template, object dados, OpcoesRascunho? opcoes = null, CancellationToken cancelamento = default);

        /// <summary>
        /// Salva o template, SANITIZANDO o corpo antes de gravar (automático: não é escolha do dev nem
        /// do usuário). Se algo for removido do HTML, o resultado avisa — e o log registra o que saiu.
        /// </summary>
        /// <param name="template">Template a salvar.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado do salvamento, com o aviso de sanitização quando houver.</returns>
        Task<ResultadoSalvamento> SalvarTemplateAsync(ITemplateEmail template, CancellationToken cancelamento = default);

        /// <summary>Autentica no servidor e desconecta, sem enviar nada. Serve para conferir a configuração.</summary>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Sucesso, ou a mensagem do erro.</returns>
        Task<ResultadoEnvio> TestarConexaoAsync(CancellationToken cancelamento = default);
    }

    /// <summary>Resultado do salvamento de um template.</summary>
    /// <param name="Sucesso">Se foi salvo.</param>
    /// <param name="Mensagem">Confirmação ou erro.</param>
    /// <param name="HtmlAlterado">Verdadeiro quando a sanitização mudou o corpo em relação ao que o usuário escreveu.</param>
    /// <param name="Remocoes">O que a sanitização removeu (vazio quando nada mudou).</param>
    public sealed record ResultadoSalvamento(
        bool Sucesso,
        string? Mensagem,
        bool HtmlAlterado = false,
        IReadOnlyList<RemocaoHtml>? Remocoes = null);

    /// <summary>
    /// Implementação da fachada. Concentra tudo o que é comum ao envio e delega a entrega ao
    /// <see cref="ITransporteEmail"/> registrado.
    /// </summary>
    public sealed class ServicoEmail : IServicoEmail
    {
        /// <summary>Transporte que entrega a mensagem (Exchange, SMTP ou simulado).</summary>
        private readonly ITransporteEmail _transporte;

        /// <summary>Configurações do sistema.</summary>
        private readonly OpcoesEmail _opcoes;

        /// <summary>Motor de marcadores, usado apenas no envio por template.</summary>
        private readonly MotorTemplate _motor;

        /// <summary>Repositório de templates; necessário apenas para <see cref="EnviarPorTemplateAsync"/>.</summary>
        private readonly IRepositorioTemplateEmail? _repositorio;

        /// <summary>Registrador do log de envios. Acionado em TODO envio — sucesso e falha.</summary>
        private readonly RegistradorEnvioEmail _log = new();

        /// <summary>Acesso aos anexos vinculados ao template (acervo + vínculo).</summary>
        private readonly ServicoAnexos? _anexos;

        /// <summary>Cria o serviço de e-mail.</summary>
        /// <param name="transporte">Transporte registrado.</param>
        /// <param name="opcoes">Configurações do sistema.</param>
        /// <param name="motor">Motor de marcadores (envio por template).</param>
        /// <param name="repositorio">Repositório de templates (opcional; exigido só no envio por nome).</param>
        /// <param name="anexos">Serviço de anexos; necessário quando os templates têm anexos vinculados.</param>
        public ServicoEmail(
            ITransporteEmail transporte,
            OpcoesEmail opcoes,
            MotorTemplate motor,
            IRepositorioTemplateEmail? repositorio = null,
            ServicoAnexos? anexos = null)
        {
            _transporte = transporte ?? throw new ArgumentNullException(nameof(transporte));
            _opcoes = opcoes ?? throw new ArgumentNullException(nameof(opcoes));
            _motor = motor ?? throw new ArgumentNullException(nameof(motor));
            _repositorio = repositorio;
            _anexos = anexos;
        }

        /// <inheritdoc/>
        public async Task<ResultadoEnvio> EnviarAsync(EmailAvulso email, CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(email);

            List<string> excluirAposEnvio = [];

            try
            {
                // Avulso: NADA é resolvido — o texto vai literal, marcadores inclusive.
                IReadOnlyList<AnexoMensagem> anexos = MaterializarAvulsos(email, excluirAposEnvio);

                MensagemEmail mensagem = new(
                    string.IsNullOrWhiteSpace(email.Remetente) ? _opcoes.EnderecoRemetente : email.Remetente,
                    _opcoes.NomeExibicao,
                    [.. email.Para],
                    [.. email.Cc],
                    [.. email.Cco],
                    email.Assunto,
                    email.CorpoHtml,
                    email.Classificacao,
                    anexos);

                // Avulso: não há template de origem.
                return await DespacharAsync(mensagem, excluirAposEnvio, nomeTemplate: null, cancelamento);
            }
            catch (FalhaAnexoException ex)
            {
                return new ResultadoEnvio(false, ex.Message);
            }
            catch (Exception ex)
            {
                return new ResultadoEnvio(false, $"Falha no envio: {ex.Message}");
            }
        }

        /// <inheritdoc/>
        public async Task<ResultadoEnvio> EnviarPorTemplateAsync(string nomeTemplate, object dados, CancellationToken cancelamento = default)
        {
            if (_repositorio is null)
                return new ResultadoEnvio(false, "Nenhum IRepositorioTemplateEmail registrado — não há como buscar o template pelo nome.");

            ITemplateEmail? template = await _repositorio.ObterAsync(nomeTemplate, cancelamento);

            if (template is null)
                return new ResultadoEnvio(false, $"Template '{nomeTemplate}' não encontrado.");

            return await EnviarTemplateAsync(template, dados, permitirRascunho: false, cancelamento);
        }

        /// <inheritdoc/>
        public async Task<ResultadoEnvio> EnviarTemplateAsync(
            ITemplateEmail template,
            object dados,
            bool permitirRascunho = false,
            CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(template);

            // TRAVA: rascunho não vai para o mundo. Publicar é o que libera o template para uso.
            if (template.Situacao == SituacaoTemplate.Rascunho && !permitirRascunho)
                return new ResultadoEnvio(
                    false,
                    $"O template '{template.Nome}' está em rascunho — publique-o antes de enviar.");

            List<string> excluirAposEnvio = [];

            try
            {
                // Template: AQUI os marcadores {{campo}} são resolvidos com o objeto de dados.
                EmailResolvido resolvido = ResolvedorEmail.Resolver(template, dados, _motor);

                // Sobrou marcador sem valor? A política decide: abortar, ou enviar sem aquele trecho.
                (EmailResolvido tratado, string? erroMarcador) = AplicarPoliticaDeMarcadores(resolvido);

                if (erroMarcador is not null)
                    return new ResultadoEnvio(false, erroMarcador);

                resolvido = tratado;

                IReadOnlyList<AnexoMensagem> anexos = await MaterializarDoTemplateAsync(
                    template, dados, resolvido.AcaoNaFalhaDeAnexo, excluirAposEnvio, cancelamento);

                MensagemEmail mensagem = new(
                    _opcoes.EnderecoRemetente,
                    _opcoes.NomeExibicao,
                    resolvido.Destinatarios,
                    resolvido.Cc,
                    resolvido.Cco,
                    resolvido.Assunto,
                    resolvido.Corpo,
                    resolvido.Classificacao,
                    anexos);

                return await DespacharAsync(mensagem, excluirAposEnvio, template.Nome, cancelamento);
            }
            catch (FalhaAnexoException ex)
            {
                return new ResultadoEnvio(false, ex.Message);
            }
            catch (Exception ex)
            {
                return new ResultadoEnvio(false, $"Falha no envio: {ex.Message}");
            }
        }

        /// <inheritdoc/>
        public async Task<ResultadoRascunho> GerarRascunhoAsync(
            EmailAvulso email,
            OpcoesRascunho? opcoes = null,
            CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(email);

            try
            {
                // Rascunho não é envio: os anexos são materializados, mas NADA é excluído do disco
                // (o arquivo de origem ainda pode ser necessário quando o e-mail for realmente enviado).
                List<string> ignorado = [];
                IReadOnlyList<AnexoMensagem> anexos = MaterializarAvulsos(email, ignorado);

                MensagemEmail mensagem = new(
                    string.IsNullOrWhiteSpace(email.Remetente) ? _opcoes.EnderecoRemetente : email.Remetente,
                    _opcoes.NomeExibicao,
                    [.. email.Para],
                    [.. email.Cc],
                    [.. email.Cco],
                    email.Assunto,
                    email.CorpoHtml,
                    email.Classificacao,
                    anexos);

                return await MaterializarRascunhoAsync(mensagem, opcoes, cancelamento);
            }
            catch (FalhaAnexoException ex)
            {
                return new ResultadoRascunho(false, ex.Message);
            }
            catch (Exception ex)
            {
                return new ResultadoRascunho(false, $"Falha ao gerar o rascunho: {ex.Message}");
            }
        }

        /// <inheritdoc/>
        public async Task<ResultadoRascunho> GerarRascunhoTemplateAsync(
            ITemplateEmail template,
            object dados,
            OpcoesRascunho? opcoes = null,
            CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(template);

            try
            {
                // O rascunho sai SEMPRE resolvido: é exatamente o e-mail que seria enviado —
                // por isso passa pela MESMA política de marcadores (o .eml nunca leva "{{campo}}" cru).
                EmailResolvido resolvido = ResolvedorEmail.Resolver(template, dados, _motor);

                (EmailResolvido tratado, string? erroMarcador) = AplicarPoliticaDeMarcadores(resolvido);

                if (erroMarcador is not null)
                    return new ResultadoRascunho(false, erroMarcador);

                resolvido = tratado;

                // No rascunho, os anexos de disco NÃO são excluídos: o e-mail ainda não foi enviado.
                List<string> ignorado = [];

                IReadOnlyList<AnexoMensagem> anexos = await MaterializarDoTemplateAsync(
                    template, dados, resolvido.AcaoNaFalhaDeAnexo, ignorado, cancelamento);

                MensagemEmail mensagem = new(
                    _opcoes.EnderecoRemetente,
                    _opcoes.NomeExibicao,
                    resolvido.Destinatarios,
                    resolvido.Cc,
                    resolvido.Cco,
                    resolvido.Assunto,
                    resolvido.Corpo,
                    resolvido.Classificacao,
                    anexos);

                return await MaterializarRascunhoAsync(mensagem, opcoes, cancelamento);
            }
            catch (FalhaAnexoException ex)
            {
                return new ResultadoRascunho(false, ex.Message);
            }
            catch (Exception ex)
            {
                return new ResultadoRascunho(false, $"Falha ao gerar o rascunho: {ex.Message}");
            }
        }

        /// <summary>Gera o .eml e o entrega nos destinos pedidos (bytes, arquivo e/ou caixa postal).</summary>
        /// <param name="mensagem">Mensagem pronta (a mesma que seria enviada).</param>
        /// <param name="opcoes">Destinos desejados.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado com o que foi produzido.</returns>
        private async Task<ResultadoRascunho> MaterializarRascunhoAsync(
            MensagemEmail mensagem,
            OpcoesRascunho? opcoes,
            CancellationToken cancelamento)
        {
            OpcoesRascunho destino = opcoes ?? new OpcoesRascunho();

            if (destino.Destino == DestinoRascunho.Nenhum)
                return new ResultadoRascunho(false, "Nenhum destino de rascunho informado.");

            string nomeArquivo = string.IsNullOrWhiteSpace(destino.NomeArquivo)
                ? GerarNomeArquivo(mensagem.Assunto)
                : destino.NomeArquivo;

            byte[]? conteudo = null;
            string? caminhoGravado = null;
            bool naCaixaPostal = false;
            List<string> avisos = [];

            // O .eml é a MESMA mensagem que o transporte enviaria.
            bool precisaDoEml = destino.Destino.HasFlag(DestinoRascunho.Bytes)
                || destino.Destino.HasFlag(DestinoRascunho.Arquivo);

            byte[]? eml = precisaDoEml ? MontadorMime.GerarEml(mensagem) : null;

            if (destino.Destino.HasFlag(DestinoRascunho.Bytes))
                conteudo = eml;

            if (destino.Destino.HasFlag(DestinoRascunho.Arquivo) && eml is not null)
            {
                string pasta = string.IsNullOrWhiteSpace(destino.PastaArquivo)
                    ? (string.IsNullOrWhiteSpace(_opcoes.PastaSimulacao)
                        ? Path.Combine(Path.GetTempPath(), "bbi-rascunhos")
                        : _opcoes.PastaSimulacao)
                    : destino.PastaArquivo;

                Directory.CreateDirectory(pasta);
                caminhoGravado = Path.Combine(pasta, nomeArquivo);

                await File.WriteAllBytesAsync(caminhoGravado, eml, cancelamento);
            }

            if (destino.Destino.HasFlag(DestinoRascunho.CaixaPostal))
            {
                if (_transporte is ITransporteRascunho suportaRascunho)
                {
                    await suportaRascunho.SalvarRascunhoAsync(mensagem, cancelamento);
                    naCaixaPostal = true;
                }
                else
                {
                    avisos.Add("o transporte registrado não salva na caixa postal (só o Exchange faz isso)");
                }
            }

            string resumo = avisos.Count > 0
                ? "Rascunho gerado, com ressalva: " + string.Join("; ", avisos) + "."
                : "Rascunho gerado.";

            return new ResultadoRascunho(true, resumo, conteudo, nomeArquivo, caminhoGravado, naCaixaPostal);
        }

        /// <summary>Gera um nome de arquivo .eml a partir do assunto e do horário, sem caracteres inválidos.</summary>
        /// <param name="assunto">Assunto do e-mail.</param>
        /// <returns>Nome do arquivo (ex.: "aviso-pedido-20260711-1530.eml").</returns>
        private static string GerarNomeArquivo(string assunto)
        {
            string baseNome = string.IsNullOrWhiteSpace(assunto) ? "rascunho" : assunto;

            foreach (char invalido in Path.GetInvalidFileNameChars())
                baseNome = baseNome.Replace(invalido, '-');

            baseNome = baseNome.Trim().Replace(' ', '-');

            if (baseNome.Length > 60)
                baseNome = baseNome[..60];

            return $"{baseNome}-{DateTime.Now:yyyyMMdd-HHmmss}.eml";
        }

        /// <inheritdoc/>
        public async Task<ResultadoSalvamento> SalvarTemplateAsync(ITemplateEmail template, CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(template);

            if (_repositorio is null)
                return new ResultadoSalvamento(false, "Nenhum IRepositorioTemplateEmail registrado.");

            IReadOnlyList<RemocaoHtml> remocoes = [];

            // A sanitização vale para o corpo do MODO AVANÇADO — o normal é gerado pela própria
            // biblioteca e não tem como conter código.
            if (template.OrigemCriacao == OrigemCriacaoTemplate.Avancado)
            {
                ResultadoSanitizacao limpeza = SanitizadorHtml.Sanitizar(template.Corpo);

                if (limpeza.Alterado)
                {
                    template.Corpo = limpeza.Html;
                    remocoes = limpeza.Remocoes;

                    // O que foi removido FICA REGISTRADO: se o usuário reclamar depois de que o e-mail
                    // saiu diferente do que ele colou, sabemos exatamente o que saiu e de onde veio.
                    await _log.RegistrarSanitizacaoAsync(template.Nome, remocoes, cancelamento);
                }
            }

            await _repositorio.SalvarAsync(template, cancelamento);

            string mensagem = remocoes.Count > 0
                ? $"Template salvo. Removemos do HTML elementos não permitidos: {string.Join(", ", remocoes.Select(r => r.Elemento).Distinct())}."
                : "Template salvo.";

            return new ResultadoSalvamento(true, mensagem, remocoes.Count > 0, remocoes);
        }

        /// <inheritdoc/>
        public Task<ResultadoEnvio> TestarConexaoAsync(CancellationToken cancelamento = default)
            => _transporte.TestarConexaoAsync(cancelamento);

        /// <summary>
        /// Trata os marcadores que ficaram SEM valor, conforme a política do template:
        /// aborta o envio, ou REMOVE os trechos não resolvidos do texto (o destinatário nunca vê
        /// um "{{campo}}" cru). Depois da remoção, confere se ainda sobrou destinatário.
        /// </summary>
        /// <param name="resolvido">E-mail já resolvido, possivelmente com marcadores pendentes.</param>
        /// <returns>O e-mail tratado e, quando o envio deve ser abortado, a mensagem de erro.</returns>
        private static (EmailResolvido Email, string? Erro) AplicarPoliticaDeMarcadores(EmailResolvido resolvido)
        {
            if (resolvido.MarcadoresNaoResolvidos.Count == 0)
                return (resolvido, null);

            string lista = string.Join(", ", resolvido.MarcadoresNaoResolvidos.Select(m => "{{" + m + "}}"));

            switch (resolvido.AcaoNaFalhaDeMarcador)
            {
                case AcaoFalhaMarcador.FalharEnvio:
                    return (resolvido, $"Envio abortado — marcadores sem valor: {lista}.");

                case AcaoFalhaMarcador.EnviarMesmoAssim:
                default:
                    IReadOnlyList<string> pendentes = resolvido.MarcadoresNaoResolvidos;

                    // Os anexos NÃO entram aqui: eles vivem no acervo e são materializados à parte
                    // (o caminho dinâmico é resolvido lá, com o mesmo objeto de dados).
                    EmailResolvido limpo = resolvido with
                    {
                        Assunto = MotorTemplate.RemoverMarcadores(resolvido.Assunto, pendentes),
                        Corpo = MotorTemplate.RemoverMarcadores(resolvido.Corpo, pendentes),
                        Destinatarios = LimparEnderecos(resolvido.Destinatarios, pendentes),
                        Cc = LimparEnderecos(resolvido.Cc, pendentes),
                        Cco = LimparEnderecos(resolvido.Cco, pendentes)
                    };

                    // Remover o marcador de um endereço pode ter zerado a lista — aí não há e-mail a enviar.
                    if (limpo.Destinatarios.Count == 0)
                        return (limpo, $"Envio abortado — sem destinatário depois de remover os marcadores sem valor: {lista}.");

                    return (limpo, null);
            }
        }

        /// <summary>Remove os marcadores pendentes de uma lista de endereços e descarta os que ficaram vazios.</summary>
        /// <param name="enderecos">Endereços resolvidos.</param>
        /// <param name="pendentes">Marcadores sem valor.</param>
        /// <returns>Endereços válidos restantes.</returns>
        private static IReadOnlyList<string> LimparEnderecos(IReadOnlyList<string> enderecos, IReadOnlyList<string> pendentes)
        {
            List<string> limpos = [];

            foreach (string endereco in enderecos)
            {
                string tratado = MotorTemplate.RemoverMarcadores(endereco, pendentes).Trim();

                if (!string.IsNullOrWhiteSpace(tratado))
                    limpos.Add(tratado);
            }

            return limpos;
        }

        /// <summary>Aplica o redirecionamento, entrega (com repetição em falha transitória) e faz a limpeza pós-envio.</summary>
        /// <param name="mensagem">Mensagem pronta.</param>
        /// <param name="excluirAposEnvio">Arquivos a excluir depois do sucesso.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado do envio.</returns>
        private async Task<ResultadoEnvio> DespacharAsync(
            MensagemEmail mensagem,
            List<string> excluirAposEnvio,
            string? nomeTemplate,
            CancellationToken cancelamento)
        {
            MensagemEmail efetiva = AplicarRedirecionamento(mensagem);

            try
            {
                await EntregarComRepeticaoAsync(efetiva, cancelamento);
            }
            catch (OperationCanceledException)
            {
                // Cancelamento também é falha do ponto de vista do log: houve tentativa.
                await _log.RegistrarFalhaAsync(
                    RegistroEnvioEmail.De(efetiva, nomeTemplate, sucesso: false, "Envio cancelado."),
                    CancellationToken.None);

                return new ResultadoEnvio(false, "Envio cancelado.");
            }
            catch (Exception ex)
            {
                // LOG DE FALHA: grava o e-mail que se tentou enviar e a causa.
                await _log.RegistrarFalhaAsync(
                    RegistroEnvioEmail.De(efetiva, nomeTemplate, sucesso: false, ex.Message),
                    cancelamento);

                return new ResultadoEnvio(false, $"Falha no envio: {ex.Message}");
            }

            // LOG DE SUCESSO: logo após o envio, antes de qualquer limpeza.
            await _log.RegistrarSucessoAsync(
                RegistroEnvioEmail.De(efetiva, nomeTemplate, sucesso: true, null),
                cancelamento);

            foreach (string caminho in excluirAposEnvio)
                TentarExcluir(caminho);

            return new ResultadoEnvio(true, "E-mail enviado.");
        }

        /// <summary>Entrega a mensagem, repetindo em falha transitória com espera crescente.</summary>
        /// <param name="mensagem">Mensagem pronta.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        private async Task EntregarComRepeticaoAsync(MensagemEmail mensagem, CancellationToken cancelamento)
        {
            int tentativasRestantes = Math.Max(0, _opcoes.TentativasEmFalha);
            int espera = Math.Max(1, _opcoes.EsperaEntreTentativasMs);

            while (true)
            {
                try
                {
                    await _transporte.EntregarAsync(mensagem, cancelamento);
                    return;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception) when (tentativasRestantes > 0)
                {
                    tentativasRestantes--;
                    await Task.Delay(espera, cancelamento);
                    espera *= 2;
                }
            }
        }

        /// <summary>Aplica a trava de homologação: troca os destinatários e prefixa o assunto.</summary>
        /// <param name="mensagem">Mensagem original.</param>
        /// <returns>A própria mensagem, ou uma cópia redirecionada.</returns>
        private MensagemEmail AplicarRedirecionamento(MensagemEmail mensagem)
        {
            if (_opcoes.RedirecionarPara.Count == 0)
                return mensagem;

            return mensagem with
            {
                Assunto = _opcoes.PrefixoAssuntoTeste + mensagem.Assunto,
                Destinatarios = [.. _opcoes.RedirecionarPara],
                Cc = [],
                Cco = []
            };
        }

        /// <summary>Materializa os anexos de um e-mail avulso.</summary>
        /// <param name="email">E-mail avulso.</param>
        /// <param name="excluirAposEnvio">Recebe os caminhos a excluir depois do envio.</param>
        /// <returns>Anexos prontos.</returns>
        private IReadOnlyList<AnexoMensagem> MaterializarAvulsos(EmailAvulso email, List<string> excluirAposEnvio)
        {
            List<AnexoMensagem> materializados = [];

            foreach (AnexoAvulso anexo in email.Anexos)
            {
                byte[]? bytes = ObterBytes(
                    anexo.Conteudo,
                    anexo.Caminho,
                    anexo.NomeArquivo,
                    anexo.ExcluirAposEnviar,
                    email.AcaoNaFalhaDeAnexo,
                    excluirAposEnvio);

                if (bytes is null)
                    continue;

                materializados.Add(new AnexoMensagem(anexo.NomeArquivo, anexo.ContentType, bytes, anexo.ContentId));
            }

            return materializados;
        }

        /// <summary>
        /// Materializa os anexos VINCULADOS ao template: busca cada arquivo do acervo, resolve os
        /// marcadores do caminho (quando dinâmico) e lê os bytes. O ContentId do vínculo é o que
        /// transforma a imagem em recurso do corpo (cid:) em vez de anexo no clipe.
        /// </summary>
        /// <param name="template">Template de origem.</param>
        /// <param name="dados">Objeto de dados (resolve o caminho dinâmico).</param>
        /// <param name="politica">Política em caso de anexo que falhe.</param>
        /// <param name="excluirAposEnvio">Recebe os caminhos a excluir depois do envio.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Anexos prontos.</returns>
        private async Task<IReadOnlyList<AnexoMensagem>> MaterializarDoTemplateAsync(
            ITemplateEmail template,
            object dados,
            AcaoFalhaAnexo politica,
            List<string> excluirAposEnvio,
            CancellationToken cancelamento)
        {
            if (_anexos is null)
                return [];

            IReadOnlyList<AnexoVinculado> vinculados = await _anexos.ListarDoTemplateAsync(template.Id, cancelamento);
            List<AnexoMensagem> materializados = [];

            foreach (AnexoVinculado item in vinculados.OrderBy(v => v.Vinculo.Ordem))
            {
                string? caminho = null;

                switch (item.Anexo.ModoObtencao)
                {
                    case ModoObtencaoAnexo.CaminhoFixo:
                        caminho = item.Anexo.Caminho;
                        break;
                    case ModoObtencaoAnexo.CaminhoDinamico:
                        // O caminho tem marcadores: resolve com o objeto de dados.
                        caminho = _motor.Resolver(item.Anexo.Caminho ?? string.Empty, dados).Texto;
                        break;
                    case ModoObtencaoAnexo.BytesNoBanco:
                    default:
                        break;
                }

                byte[]? bytes = ObterBytes(
                    item.Anexo.Conteudo,
                    caminho,
                    item.Anexo.NomeArquivo,
                    item.Vinculo.ExcluirAposAnexar,
                    politica,
                    excluirAposEnvio);

                if (bytes is null)
                    continue;

                materializados.Add(new AnexoMensagem(
                    item.Anexo.NomeArquivo,
                    item.Anexo.ContentType,
                    bytes,
                    item.Vinculo.ContentId));   // presente = recurso do corpo; nulo = anexo no clipe
            }

            return materializados;
        }

        /// <summary>Obtém os bytes de um anexo (de memória ou de disco), aplicando a política de falha.</summary>
        /// <param name="conteudo">Bytes já em memória, quando houver.</param>
        /// <param name="caminho">Caminho em disco, quando houver.</param>
        /// <param name="nomeArquivo">Nome do anexo (para a mensagem de erro).</param>
        /// <param name="excluirDepois">Se o arquivo deve ser excluído após o envio.</param>
        /// <param name="politica">Política em caso de falha.</param>
        /// <param name="excluirAposEnvio">Lista que recebe o caminho a excluir.</param>
        /// <returns>Bytes do anexo, ou nulo para pular (quando a política permite).</returns>
        private byte[]? ObterBytes(
            byte[]? conteudo,
            string? caminho,
            string nomeArquivo,
            bool excluirDepois,
            AcaoFalhaAnexo politica,
            List<string> excluirAposEnvio)
        {
            if (conteudo is not null)
                return conteudo;

            if (string.IsNullOrWhiteSpace(caminho))
                return TratarFalha(politica, $"anexo '{nomeArquivo}' sem conteúdo nem caminho");

            try
            {
                string seguro = CaminhoSeguro(caminho);
                byte[] bytes = File.ReadAllBytes(seguro);

                if (excluirDepois)
                    excluirAposEnvio.Add(seguro);

                return bytes;
            }
            catch (FalhaAnexoException)
            {
                throw;
            }
            catch (Exception ex)
            {
                return TratarFalha(politica, $"falha ao ler o anexo '{nomeArquivo}': {ex.Message}");
            }
        }

        /// <summary>Aplica a política de falha: pula o anexo ou aborta o envio.</summary>
        /// <param name="politica">Política configurada.</param>
        /// <param name="motivo">Descrição da falha.</param>
        /// <returns>Nulo quando o anexo deve ser pulado.</returns>
        private static byte[]? TratarFalha(AcaoFalhaAnexo politica, string motivo)
        {
            switch (politica)
            {
                case AcaoFalhaAnexo.FalharEnvio:
                    throw new FalhaAnexoException($"Envio abortado — {motivo}.");
                case AcaoFalhaAnexo.EnviarMesmoAssim:
                default:
                    return null;
            }
        }

        /// <summary>Resolve o caminho do anexo confinando-o à pasta base (barreira contra path traversal).</summary>
        /// <param name="caminho">Caminho informado.</param>
        /// <returns>Caminho absoluto seguro, dentro da pasta base.</returns>
        private string CaminhoSeguro(string caminho)
        {
            if (string.IsNullOrWhiteSpace(_opcoes.PastaBaseAnexos))
                throw new FalhaAnexoException("PastaBaseAnexos não configurada — necessária para anexar arquivos de disco.");

            string baseDir = Path.GetFullPath(_opcoes.PastaBaseAnexos);
            string baseComBarra = baseDir.EndsWith(Path.DirectorySeparatorChar)
                ? baseDir
                : baseDir + Path.DirectorySeparatorChar;

            string combinado = Path.GetFullPath(Path.Combine(baseDir, caminho));

            if (!combinado.StartsWith(baseComBarra, StringComparison.OrdinalIgnoreCase))
                throw new FalhaAnexoException($"Caminho de anexo fora da pasta base permitida: {caminho}");

            return combinado;
        }

        /// <summary>Tenta excluir um arquivo (melhor esforço; não derruba um envio bem-sucedido).</summary>
        /// <param name="caminho">Caminho do arquivo.</param>
        private static void TentarExcluir(string caminho)
        {
            try
            {
                if (File.Exists(caminho))
                    File.Delete(caminho);
            }
            catch
            {
                // Exclusão é melhor esforço.
            }
        }
    }

    // #endregion

}
