// =============================================================================
//  Email.cs  —  Módulo de e-mail: motor de template, contrato, envio e
//  repositório (namespace BBiCore.Email). Componentes BBiEmail/BBiComporEmail
//  na mesma pasta.
// =============================================================================

using System.Collections.Concurrent;
using System.Globalization;
using System.Linq.Expressions;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;
using BBiCore.Templates;

namespace BBiCore.Email
{

// #region Enums do contrato  ->  destino sugerido: Models/

/// <summary>Indica em qual modo o template foi criado.</summary>
public enum OrigemCriacaoTemplate
{
    /// <summary>Modo assistido: cabeçalho/rodapé por imagem + textos simples com marcadores.</summary>
    Normal,

    /// <summary>Modo avançado: HTML editado livremente.</summary>
    Avancado
}

/// <summary>Política quando um anexo não pode ser resolvido no momento do envio.</summary>
public enum AcaoFalhaAnexo
{
    /// <summary>Envia o e-mail mesmo sem o anexo que falhou.</summary>
    EnviarMesmoAssim,

    /// <summary>Falha o envio inteiro se algum anexo não puder ser resolvido.</summary>
    FalharEnvio
}

/// <summary>
/// O que fazer quando, no momento do envio, sobrar algum marcador SEM resolver (o campo não existe
/// no objeto de dados, ou veio nulo). Espelha a política de anexos.
/// </summary>
public enum AcaoFalhaMarcador
{
    /// <summary>Aborta o envio e informa quais marcadores ficaram sem valor. É o padrão seguro.</summary>
    FalharEnvio,

    /// <summary>Envia assim mesmo, REMOVENDO do texto os trechos não resolvidos (o destinatário nunca vê um "{{campo}}" cru).</summary>
    EnviarMesmoAssim
}

/// <summary>Situação do template no cadastro.</summary>
public enum SituacaoTemplate
{
    /// <summary>Em edição: pode ser salvo incompleto (sem destinatário, com marcador faltando, corpo vazio).</summary>
    Rascunho,

    /// <summary>Pronto para uso: passou pelas validações de envio.</summary>
    Publicado
}

/// <summary>Rótulo de sensibilidade do e-mail.</summary>
public enum ClassificacaoEmail
{
    /// <summary>Uso interno.</summary>
    Interno,

    /// <summary>Conteúdo confidencial.</summary>
    Confidencial,

    /// <summary>Conteúdo público.</summary>
    Publico
}

// #endregion

// #region Contrato (interfaces)  ->  destino sugerido: Models/ (ou Contracts/)



/// <summary>Contrato mínimo de um template de e-mail. Cada sistema implementa na sua entidade <c>Email.Templates</c>.</summary>
public interface ITemplateEmail
{
    /// <summary>Identificador do template (chave da tabela). É por ele que os anexos são vinculados.</summary>
    int Id { get; set; }

    /// <summary>Nome/identificação do template (usado como chave na persistência).</summary>
    string Nome { get; set; }

    /// <summary>Assunto do e-mail (pode conter marcadores).</summary>
    string Assunto { get; set; }

    /// <summary>Destinatários (pode conter marcadores). Lista separada por ';'.</summary>
    string Destinatarios { get; set; }

    /// <summary>Cópia (Cc), separada por ';'. Pode conter marcadores.</summary>
    string? Cc { get; set; }

    /// <summary>Cópia oculta (Cco), separada por ';'. Pode conter marcadores.</summary>
    string? Cco { get; set; }

    /// <summary>Corpo do e-mail em HTML (fonte única de verdade; pode conter marcadores).</summary>
    string Corpo { get; set; }

    /// <summary>Modo em que o template foi criado.</summary>
    OrigemCriacaoTemplate OrigemCriacao { get; set; }

    /// <summary>Situação do template: rascunho (em edição, pode estar incompleto) ou publicado (liberado para os sistemas usarem).</summary>
    SituacaoTemplate Situacao { get; set; }

    /// <summary>O que fazer quando sobrar marcador sem resolver no envio.</summary>
    AcaoFalhaMarcador AcaoNaFalhaDeMarcador { get; set; }

    /// <summary>Nome do tipo de dados que este template espera preencher (ex.: "PedidoCliente"). Apenas referência/validação.</summary>
    string TipoDadosNome { get; set; }

    /// <summary>Política a aplicar quando um anexo falhar no envio.</summary>
    AcaoFalhaAnexo AcaoNaFalhaDeAnexo { get; set; }

    /// <summary>Rótulo de sensibilidade do e-mail.</summary>
    ClassificacaoEmail Classificacao { get; set; }
}

// #endregion

// #region DTOs de conveniência  ->  destino sugerido: Models/



/// <summary>Implementação default de <see cref="ITemplateEmail"/>. Um sistema pode usá-la ou mapear sua própria entidade.</summary>
public sealed class TemplateEmailDto : ITemplateEmail
{
    /// <inheritdoc/>
    public int Id { get; set; }

    /// <inheritdoc/>
    public string Nome { get; set; } = string.Empty;

    /// <inheritdoc/>
    public string Assunto { get; set; } = string.Empty;

    /// <inheritdoc/>
    public string Destinatarios { get; set; } = string.Empty;

    /// <inheritdoc/>
    public string? Cc { get; set; }

    /// <inheritdoc/>
    public string? Cco { get; set; }

    /// <inheritdoc/>
    public string Corpo { get; set; } = string.Empty;

    /// <inheritdoc/>
    public OrigemCriacaoTemplate OrigemCriacao { get; set; }

    /// <inheritdoc/>
    public SituacaoTemplate Situacao { get; set; } = SituacaoTemplate.Rascunho;

    /// <inheritdoc/>
    public AcaoFalhaMarcador AcaoNaFalhaDeMarcador { get; set; } = AcaoFalhaMarcador.FalharEnvio;

    /// <inheritdoc/>
    public string TipoDadosNome { get; set; } = string.Empty;

    /// <inheritdoc/>
    public AcaoFalhaAnexo AcaoNaFalhaDeAnexo { get; set; }

    /// <inheritdoc/>
    public ClassificacaoEmail Classificacao { get; set; }

    /// <summary>Cria uma cópia independente de um template (incluindo anexos), para persistência defensiva.</summary>
    /// <param name="origem">Template de origem.</param>
    /// <returns>Nova instância copiada.</returns>
    public static TemplateEmailDto Copiar(ITemplateEmail origem)
    {
        TemplateEmailDto copia = new()
        {
            Id = origem.Id,          // sem o Id, o template perderia o vínculo com os seus anexos
            Nome = origem.Nome,
            Assunto = origem.Assunto,
            Destinatarios = origem.Destinatarios,
            Cc = origem.Cc,
            Cco = origem.Cco,
            Corpo = origem.Corpo,
            OrigemCriacao = origem.OrigemCriacao,
            Situacao = origem.Situacao,
            AcaoNaFalhaDeMarcador = origem.AcaoNaFalhaDeMarcador,
            TipoDadosNome = origem.TipoDadosNome,
            AcaoNaFalhaDeAnexo = origem.AcaoNaFalhaDeAnexo,
            Classificacao = origem.Classificacao
        };

        // Os anexos NÃO são copiados aqui: eles vivem no acervo e são ligados por vínculo (IdTemplate).
        return copia;
    }
}

// #endregion

// #region Repositório de templates  ->  destino sugerido: Services/

/// <summary>Persistência de templates de e-mail. Cada sistema implementa (tipicamente com EF) contra a sua tabela local.</summary>
public interface IRepositorioTemplateEmail
{
    /// <summary>
    /// Lista os nomes dos templates salvos. Informe <paramref name="situacao"/> para filtrar —
    /// numa tela de seleção, o normal é pedir só os <see cref="SituacaoTemplate.Publicado"/>, para
    /// não oferecer rascunhos ao usuário.
    /// </summary>
    /// <param name="situacao">Situação desejada; nulo traz todos.</param>
    /// <param name="cancelamento">Token de cancelamento.</param>
    /// <returns>Nomes disponíveis.</returns>
    Task<IReadOnlyList<string>> ListarNomesAsync(SituacaoTemplate? situacao = null, CancellationToken cancelamento = default);

    /// <summary>Obtém um template pelo nome.</summary>
    /// <param name="nome">Nome do template.</param>
    /// <param name="cancelamento">Token de cancelamento.</param>
    /// <returns>O template, ou nulo se não existir.</returns>
    Task<ITemplateEmail?> ObterAsync(string nome, CancellationToken cancelamento = default);

    /// <summary>Salva (cria ou substitui) um template pelo nome.</summary>
    /// <param name="template">Template a salvar.</param>
    /// <param name="cancelamento">Token de cancelamento.</param>
    Task SalvarAsync(ITemplateEmail template, CancellationToken cancelamento = default);

    /// <summary>Exclui um template pelo nome.</summary>
    /// <param name="nome">Nome do template.</param>
    /// <param name="cancelamento">Token de cancelamento.</param>
    Task ExcluirAsync(string nome, CancellationToken cancelamento = default);
}

/// <summary>
/// Implementação de referência de <see cref="IRepositorioTemplateEmail"/> em memória (thread-safe). Serve para
/// testes e demonstração; em produção, cada sistema implementa a persistência real (EF Core na tabela local).
/// Guarda cópias independentes (não referências), evitando que edições posteriores afetem o que foi salvo.
/// </summary>
public sealed class RepositorioTemplateEmailMemoria : IRepositorioTemplateEmail
{
    /// <summary>Próximo identificador a atribuir (no banco, isto é a coluna IDENTITY).</summary>
    private int _proximoId = 1;

    /// <summary>Armazenamento por nome (case-insensitive).</summary>
    private readonly ConcurrentDictionary<string, ITemplateEmail> _dados =
        new(StringComparer.OrdinalIgnoreCase);

    /// <inheritdoc/>
    public Task<IReadOnlyList<string>> ListarNomesAsync(SituacaoTemplate? situacao = null, CancellationToken cancelamento = default)
    {
        IEnumerable<KeyValuePair<string, ITemplateEmail>> itens = _dados;

        if (situacao is not null)
            itens = itens.Where(par => par.Value.Situacao == situacao.Value);

        IReadOnlyList<string> nomes = [.. itens
            .Select(par => par.Key)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)];

        return Task.FromResult(nomes);
    }

    /// <inheritdoc/>
    public Task<ITemplateEmail?> ObterAsync(string nome, CancellationToken cancelamento = default)
    {
        if (_dados.TryGetValue(nome, out ITemplateEmail? achado))
            return Task.FromResult<ITemplateEmail?>(TemplateEmailDto.Copiar(achado));

        return Task.FromResult<ITemplateEmail?>(null);
    }

    /// <inheritdoc/>
    public Task SalvarAsync(ITemplateEmail template, CancellationToken cancelamento = default)
    {
        ArgumentNullException.ThrowIfNull(template);

        // Template novo ganha identificador — é o que permite vincular anexos a ele.
        if (template.Id == 0)
            template.Id = _proximoId++;

        if (string.IsNullOrWhiteSpace(template.Nome))
            throw new InvalidOperationException("O template precisa de um Nome para ser salvo.");

        _dados[template.Nome] = TemplateEmailDto.Copiar(template);
        return Task.CompletedTask;
    }

    /// <inheritdoc/>
    public Task ExcluirAsync(string nome, CancellationToken cancelamento = default)
    {
        _dados.TryRemove(nome, out _);
        return Task.CompletedTask;
    }
}


/// <summary>E-mail de template com todos os marcadores já resolvidos, pronto para o serviço de envio.</summary>
/// <param name="Assunto">Assunto final.</param>
/// <param name="Destinatarios">Destinatários (já separados).</param>
/// <param name="Cc">Cópia.</param>
/// <param name="Cco">Cópia oculta.</param>
/// <param name="Corpo">Corpo HTML final.</param>
/// <param name="Classificacao">Rótulo de sensibilidade.</param>
/// <param name="AcaoNaFalhaDeAnexo">Política em caso de anexo que falhe.</param>
/// <param name="MarcadoresNaoResolvidos">Marcadores que ficaram SEM valor (nome do campo, sem as chaves). Vazio quando tudo resolveu.</param>
/// <param name="AcaoNaFalhaDeMarcador">Política em caso de marcador não resolvido.</param>
public sealed record EmailResolvido(
    string Assunto,
    IReadOnlyList<string> Destinatarios,
    IReadOnlyList<string> Cc,
    IReadOnlyList<string> Cco,
    string Corpo,
    ClassificacaoEmail Classificacao,
    AcaoFalhaAnexo AcaoNaFalhaDeAnexo,
    IReadOnlyList<string> MarcadoresNaoResolvidos,
    AcaoFalhaMarcador AcaoNaFalhaDeMarcador);

/// <summary>Resolve um <see cref="ITemplateEmail"/> (com marcadores) em um <see cref="EmailResolvido"/> pronto para envio.</summary>
public static class ResolvedorEmail
{
    /// <summary>Resolve todos os campos do template com os dados informados.</summary>
    /// <param name="template">Template com marcadores.</param>
    /// <param name="dados">Objeto de dados (pode ser nulo).</param>
    /// <param name="motor">Motor de template configurado.</param>
    /// <returns>E-mail pronto para envio.</returns>
    public static EmailResolvido Resolver(ITemplateEmail template, object? dados, MotorTemplate motor)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(motor);

        // Acumula, de TODOS os campos, os marcadores que não encontraram valor.
        List<string> naoResolvidos = [];

        string assunto = ResolverCampo(motor, template.Assunto, dados, naoResolvidos);
        IReadOnlyList<string> para = Separar(ResolverCampo(motor, template.Destinatarios, dados, naoResolvidos));
        IReadOnlyList<string> cc = Separar(ResolverCampo(motor, template.Cc, dados, naoResolvidos));
        IReadOnlyList<string> cco = Separar(ResolverCampo(motor, template.Cco, dados, naoResolvidos));
        string corpo = ResolverCampo(motor, template.Corpo, dados, naoResolvidos);

        return new EmailResolvido(
            assunto,
            para,
            cc,
            cco,
            corpo,
            template.Classificacao,
            template.AcaoNaFalhaDeAnexo,
            naoResolvidos.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            template.AcaoNaFalhaDeMarcador);
    }

    /// <summary>Resolve um campo e acumula os marcadores que ficaram sem valor.</summary>
    /// <param name="motor">Motor de template.</param>
    /// <param name="texto">Texto do campo (pode ser nulo).</param>
    /// <param name="dados">Objeto de dados.</param>
    /// <param name="naoResolvidos">Lista que recebe os marcadores sem valor.</param>
    /// <returns>Texto resolvido (marcadores sem valor permanecem, para o serviço decidir o que fazer).</returns>
    private static string ResolverCampo(MotorTemplate motor, string? texto, object? dados, List<string> naoResolvidos)
    {
        ResultadoTemplate resultado = motor.Resolver(texto ?? string.Empty, dados);

        if (resultado.NaoResolvidos.Count > 0)
            naoResolvidos.AddRange(resultado.NaoResolvidos);

        return resultado.Texto;
    }

    /// <summary>Separa uma lista de endereços por ';', descartando vazios e espaços.</summary>
    /// <param name="valor">Texto com endereços.</param>
    /// <returns>Endereços individuais.</returns>
    private static IReadOnlyList<string> Separar(string valor)
        => valor.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}

/// <summary>Problemas encontrados em um campo de endereços.</summary>
/// <param name="Vazio">Verdadeiro quando o campo é obrigatório e não tem nenhum endereço.</param>
/// <param name="Invalidos">Endereços com formato inválido.</param>
/// <param name="Duplicados">Endereços repetidos no mesmo campo.</param>
public sealed record ProblemasEndereco(bool Vazio, IReadOnlyList<string> Invalidos, IReadOnlyList<string> Duplicados)
{
    /// <summary>Verdadeiro quando não há nenhum problema.</summary>
    public bool Ok => !Vazio && Invalidos.Count == 0 && Duplicados.Count == 0;

    /// <summary>Descrição legível dos problemas (vazia quando <see cref="Ok"/>).</summary>
    public string Mensagem
    {
        get
        {
            List<string> partes = [];
            if (Vazio) partes.Add("nenhum destinatário");
            if (Invalidos.Count > 0) partes.Add("inválido(s): " + string.Join(", ", Invalidos));
            if (Duplicados.Count > 0) partes.Add("duplicado(s): " + string.Join(", ", Duplicados));
            return string.Join("; ", partes);
        }
    }
}

/// <summary>Valida campos de endereços de e-mail (formato, vazios e duplicados).</summary>
public static class ValidadorEnderecos
{
    /// <summary>Valida uma lista de endereços separada por ';'.</summary>
    /// <param name="listaSeparadaPorPontoEVirgula">Texto do campo (já resolvido, sem marcadores).</param>
    /// <param name="obrigatorio">Se o campo precisa ter ao menos um endereço.</param>
    /// <returns>Os problemas encontrados.</returns>
    public static ProblemasEndereco Validar(string? listaSeparadaPorPontoEVirgula, bool obrigatorio)
    {
        string[] itens = (listaSeparadaPorPontoEVirgula ?? string.Empty)
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        bool vazio = obrigatorio && itens.Length == 0;

        List<string> invalidos = [];
        List<string> duplicados = [];
        HashSet<string> vistos = new(StringComparer.OrdinalIgnoreCase);

        foreach (string item in itens)
        {
            if (!EhEmailValido(item))
                invalidos.Add(item);

            if (!vistos.Add(item) && !duplicados.Contains(item, StringComparer.OrdinalIgnoreCase))
                duplicados.Add(item);
        }

        return new ProblemasEndereco(vazio, invalidos, duplicados);
    }

    /// <summary>Indica se um endereço tem formato de e-mail válido.</summary>
    /// <param name="endereco">Endereço a testar.</param>
    /// <returns>Verdadeiro se válido.</returns>
    public static bool EhEmailValido(string? endereco)
    {
        if (string.IsNullOrWhiteSpace(endereco))
            return false;

        // MailAddress cobre o formato geral; exige domínio com ponto para evitar "a@b".
        if (!MailAddress.TryCreate(endereco, out MailAddress? addr))
            return false;

        int ponto = addr.Host.LastIndexOf('.');
        return ponto > 0 && ponto < addr.Host.Length - 1;
    }
}

// #endregion

} // fim namespace BBiCore.Email
