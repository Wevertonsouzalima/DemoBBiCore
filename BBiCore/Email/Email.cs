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

/// <summary>Natureza de um anexo persistível do template.</summary>
public enum TipoAnexo
{
    /// <summary>Arquivo fixo já enviado, com os bytes guardados no banco.</summary>
    ArquivoEmbutido,

    /// <summary>Caminho montado em runtime a partir de marcadores (ex.: <c>pedido{{NomeCliente}}.csv</c>).</summary>
    CaminhoDinamico
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

/// <summary>Anexo persistível de um template (formas 1 e 3). A forma 2 (callback) é parâmetro de envio, fora do contrato.</summary>
public interface IAnexoTemplate
{
    /// <summary>Natureza do anexo.</summary>
    TipoAnexo Tipo { get; }

    /// <summary>Nome do arquivo como será anexado ao e-mail. No caminho dinâmico, pode conter marcadores.</summary>
    string NomeArquivo { get; }

    /// <summary>Content-type do anexo (ex.: <c>application/pdf</c>), quando conhecido.</summary>
    string? ContentType { get; }

    /// <summary>Bytes do arquivo (apenas para <see cref="TipoAnexo.ArquivoEmbutido"/>).</summary>
    byte[]? Conteudo { get; }

    /// <summary>Padrão do caminho/nome com marcadores (apenas para <see cref="TipoAnexo.CaminhoDinamico"/>).</summary>
    string? PadraoCaminho { get; }

    /// <summary>
    /// Quando preenchido, o anexo é uma imagem INLINE referenciada no corpo por <c>cid:{ContentId}</c>
    /// (ex.: cabeçalho = "bbi-cabecalho", rodapé = "bbi-rodape"). Nulo indica anexo comum (baixável).
    /// </summary>
    string? ContentId { get; }

    /// <summary>Se verdadeiro, o arquivo deve ser excluído do servidor após ser anexado (limpeza). Editável.</summary>
    bool ExcluirAposAnexar { get; set; }
}

/// <summary>Contrato mínimo de um template de e-mail. Cada sistema implementa na sua entidade <c>Email.Templates</c>.</summary>
public interface ITemplateEmail
{
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

    /// <summary>Nome do tipo de dados que este template espera preencher (ex.: "PedidoCliente"). Apenas referência/validação.</summary>
    string TipoDadosNome { get; set; }

    /// <summary>Política a aplicar quando um anexo falhar no envio.</summary>
    AcaoFalhaAnexo AcaoNaFalhaDeAnexo { get; set; }

    /// <summary>Rótulo de sensibilidade do e-mail.</summary>
    ClassificacaoEmail Classificacao { get; set; }

    /// <summary>Anexos persistíveis do template (mutável para edição pelo componente).</summary>
    IList<IAnexoTemplate> Anexos { get; }
}

// #endregion

// #region DTOs de conveniência  ->  destino sugerido: Models/

/// <summary>Implementação default de <see cref="IAnexoTemplate"/>, usada pelo componente ao criar anexos.</summary>
public sealed class AnexoTemplateDto : IAnexoTemplate
{
    /// <inheritdoc/>
    public TipoAnexo Tipo { get; set; }

    /// <inheritdoc/>
    public string NomeArquivo { get; set; } = string.Empty;

    /// <inheritdoc/>
    public string? ContentType { get; set; }

    /// <inheritdoc/>
    public byte[]? Conteudo { get; set; }

    /// <inheritdoc/>
    public string? PadraoCaminho { get; set; }

    /// <inheritdoc/>
    public string? ContentId { get; set; }

    /// <inheritdoc/>
    public bool ExcluirAposAnexar { get; set; }

    /// <summary>Cria uma cópia independente de um anexo (para persistência defensiva).</summary>
    /// <param name="origem">Anexo de origem.</param>
    /// <returns>Nova instância copiada.</returns>
    public static AnexoTemplateDto Copiar(IAnexoTemplate origem)
        => new()
        {
            Tipo = origem.Tipo,
            NomeArquivo = origem.NomeArquivo,
            ContentType = origem.ContentType,
            Conteudo = origem.Conteudo is null ? null : (byte[])origem.Conteudo.Clone(),
            PadraoCaminho = origem.PadraoCaminho,
            ContentId = origem.ContentId,
            ExcluirAposAnexar = origem.ExcluirAposAnexar
        };
}

/// <summary>Implementação default de <see cref="ITemplateEmail"/>. Um sistema pode usá-la ou mapear sua própria entidade.</summary>
public sealed class TemplateEmailDto : ITemplateEmail
{
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
    public string TipoDadosNome { get; set; } = string.Empty;

    /// <inheritdoc/>
    public AcaoFalhaAnexo AcaoNaFalhaDeAnexo { get; set; }

    /// <inheritdoc/>
    public ClassificacaoEmail Classificacao { get; set; }

    /// <inheritdoc/>
    public IList<IAnexoTemplate> Anexos { get; } = [];

    /// <summary>Cria uma cópia independente de um template (incluindo anexos), para persistência defensiva.</summary>
    /// <param name="origem">Template de origem.</param>
    /// <returns>Nova instância copiada.</returns>
    public static TemplateEmailDto Copiar(ITemplateEmail origem)
    {
        TemplateEmailDto copia = new()
        {
            Nome = origem.Nome,
            Assunto = origem.Assunto,
            Destinatarios = origem.Destinatarios,
            Cc = origem.Cc,
            Cco = origem.Cco,
            Corpo = origem.Corpo,
            OrigemCriacao = origem.OrigemCriacao,
            TipoDadosNome = origem.TipoDadosNome,
            AcaoNaFalhaDeAnexo = origem.AcaoNaFalhaDeAnexo,
            Classificacao = origem.Classificacao
        };

        foreach (IAnexoTemplate anexo in origem.Anexos)
            copia.Anexos.Add(AnexoTemplateDto.Copiar(anexo));

        return copia;
    }
}

// #endregion

// #region Resultado  ->  destino sugerido: Models/

/// <summary>Descreve um campo disponível para inserção como marcador (alimenta o menu "Adicionar propriedade").</summary>
/// <param name="Nome">Nome a inserir dentro de <c>{{ }}</c>.</param>
/// <param name="Rotulo">Rótulo amigável exibido no menu.</param>
/// <param name="Grupo">Grupo do campo (objeto, variável do sistema, valor automático).</param>
public sealed record CampoTemplate(string Nome, string Rotulo, GrupoCampo Grupo);

/// <summary>Origem de um campo disponível para inserção.</summary>
public enum GrupoCampo
{
    /// <summary>Propriedade do objeto de dados (ex.: PedidoCliente).</summary>
    Objeto,

    /// <summary>Variável cadastrada pelo sistema.</summary>
    VariavelSistema,

    /// <summary>Valor computado automático (prefixo "Sistema.").</summary>
    ValorAutomatico
}

/// <summary>Combinação de fontes de campo que um ponto de edição oferece no menu. Controla apenas o que é SUGERIDO, não a validação.</summary>
[Flags]
public enum FontesCampo
{
    /// <summary>Nenhuma fonte.</summary>
    Nenhuma = 0,

    /// <summary>Propriedades do objeto de dados.</summary>
    Objeto = 1,

    /// <summary>Variáveis cadastradas do sistema.</summary>
    VariavelSistema = 2,

    /// <summary>Valores computados automáticos (Sistema.*).</summary>
    ValorAutomatico = 4,

    /// <summary>Todas as fontes (padrão de assunto e corpo).</summary>
    Todas = Objeto | VariavelSistema | ValorAutomatico,

    /// <summary>Objeto + variáveis, sem automáticos (padrão dos campos de endereço).</summary>
    EnderecoEmail = Objeto | VariavelSistema
}

/// <summary>Resultado da resolução de um texto: valor final e marcadores não resolvidos.</summary>
/// <param name="Texto">Texto com os marcadores substituídos.</param>
/// <param name="NaoResolvidos">Marcadores que não existem em nenhuma fonte (substituídos por vazio).</param>
public sealed record ResultadoTemplate(string Texto, IReadOnlyList<string> NaoResolvidos)
{
    /// <summary>Verdadeiro quando todos os marcadores foram resolvidos.</summary>
    public bool Sucesso => NaoResolvidos.Count == 0;
}

// #endregion

// #region Motor multi-fonte  ->  destino sugerido: Services/

/// <summary>
/// Resolve e valida marcadores <c>{{Campo}}</c> / <c>{{Campo:formato}}</c> contra três fontes, em ordem:
/// objeto de dados → variáveis cadastradas → valores computados (prefixo "Sistema."). Configurado uma vez
/// (variáveis + computados) e reutilizado para resolver todos os campos do mesmo e-mail.
/// </summary>
public sealed partial class MotorTemplateEmail
{
    /// <summary>Prefixo reservado dos valores computados (evita colisão com objeto/variáveis).</summary>
    internal const string PrefixoSistema = "Sistema.";

    /// <summary>Cultura padrão de formatação (pt-BR).</summary>
    private static readonly CultureInfo Cultura = CultureInfo.GetCultureInfo("pt-BR");

    /// <summary>Cache de propriedades legíveis por tipo (nome → PropertyInfo, sem diferenciar maiúsculas).</summary>
    private static readonly ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>> Cache = new();

    /// <summary>Variáveis cadastradas do sistema (nome → valor).</summary>
    private readonly IReadOnlyDictionary<string, string> _variaveis;

    /// <summary>Valores computados (chave sem o prefixo → função que recebe formato e cultura).</summary>
    private readonly IReadOnlyDictionary<string, Func<string?, CultureInfo, string>> _computados;

    /// <summary>Cria o motor com as variáveis cadastradas e, opcionalmente, computados extras.</summary>
    /// <param name="variaveis">Variáveis Nome/Valor do sistema. Nulo trata como vazio.</param>
    /// <param name="computadosExtras">Computados adicionais além dos padrão (sobrescrevem os padrão em caso de mesmo nome).</param>
    public MotorTemplateEmail(
        IReadOnlyDictionary<string, string>? variaveis = null,
        IReadOnlyDictionary<string, Func<string?, CultureInfo, string>>? computadosExtras = null)
    {
        _variaveis = variaveis is null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(variaveis, StringComparer.OrdinalIgnoreCase);

        Dictionary<string, Func<string?, CultureInfo, string>> computados = ComputadosPadrao();

        if (computadosExtras is not null)
            foreach (KeyValuePair<string, Func<string?, CultureInfo, string>> par in computadosExtras)
                computados[par.Key] = par.Value;

        _computados = computados;
    }

    /// <summary>Nomes das variáveis cadastradas disponíveis.</summary>
    public IReadOnlyCollection<string> VariaveisDisponiveis => (IReadOnlyCollection<string>)_variaveis.Keys;

    /// <summary>Nomes completos (com prefixo) dos valores computados disponíveis.</summary>
    public IReadOnlyCollection<string> ComputadosDisponiveis
        => _computados.Keys.Select(k => PrefixoSistema + k).ToList();

    /// <summary>Resolve os marcadores de um texto (assunto, corpo, destinatários…) usando as fontes configuradas.</summary>
    /// <param name="texto">Texto com marcadores.</param>
    /// <param name="dados">Objeto de dados (flat) que preenche os marcadores do objeto. Pode ser nulo.</param>
    /// <returns>Resultado com o texto final e a lista de marcadores não resolvidos.</returns>
    public ResultadoTemplate Resolver(string texto, object? dados)
    {
        ArgumentNullException.ThrowIfNull(texto);

        Dictionary<string, PropertyInfo> props = dados is null
            ? []
            : ObterPropriedades(dados.GetType());

        List<string> naoResolvidos = [];

        string final = MarcadorRegex().Replace(texto, m =>
        {
            string nome = m.Groups[1].Value;
            string? formato = m.Groups[2].Success ? m.Groups[2].Value : null;

            if (TentarResolver(nome, formato, dados, props, out string valor))
                return valor;

            if (!naoResolvidos.Contains(nome))
                naoResolvidos.Add(nome);

            return string.Empty;
        });

        return new ResultadoTemplate(final, naoResolvidos);
    }

    /// <summary>Valida um texto contra um tipo (sem instância), considerando as três fontes. Uso na edição.</summary>
    /// <param name="texto">Texto a validar.</param>
    /// <param name="tipoDados">Tipo do objeto que preencheria o template.</param>
    /// <returns>Marcadores que não existem em nenhuma fonte (candidatos a destaque em vermelho).</returns>
    public IReadOnlyList<string> Validar(string texto, Type tipoDados)
    {
        ArgumentNullException.ThrowIfNull(texto);
        ArgumentNullException.ThrowIfNull(tipoDados);

        Dictionary<string, PropertyInfo> props = ObterPropriedades(tipoDados);
        List<string> faltantes = [];

        foreach (Match m in MarcadorRegex().Matches(texto))
        {
            string nome = m.Groups[1].Value;

            if (Conhece(nome, props) || faltantes.Contains(nome))
                continue;

            faltantes.Add(nome);
        }

        return faltantes;
    }

    /// <summary>Tenta resolver um marcador na ordem objeto → variáveis → computados (computado por prefixo).</summary>
    /// <param name="nome">Nome do marcador.</param>
    /// <param name="formato">Formato opcional.</param>
    /// <param name="dados">Objeto de dados.</param>
    /// <param name="props">Propriedades do objeto.</param>
    /// <param name="valor">Recebe o valor resolvido.</param>
    /// <returns>Verdadeiro se resolveu em alguma fonte.</returns>
    private bool TentarResolver(
        string nome, string? formato, object? dados,
        Dictionary<string, PropertyInfo> props, out string valor)
    {
        // Valor computado: identificado pelo prefixo reservado.
        if (nome.StartsWith(PrefixoSistema, StringComparison.OrdinalIgnoreCase))
        {
            string chave = nome[PrefixoSistema.Length..];

            if (_computados.TryGetValue(chave, out Func<string?, CultureInfo, string>? fn))
            {
                valor = fn(formato, Cultura);
                return true;
            }

            valor = string.Empty;
            return false;
        }

        // Objeto de dados.
        if (dados is not null && props.TryGetValue(nome, out PropertyInfo? prop))
        {
            valor = Formatar(prop.GetValue(dados), formato);
            return true;
        }

        // Variável cadastrada.
        if (_variaveis.TryGetValue(nome, out string? v))
        {
            valor = v ?? string.Empty;
            return true;
        }

        valor = string.Empty;
        return false;
    }

    /// <summary>Indica se um marcador é conhecido em alguma fonte (para validação em tempo de edição).</summary>
    /// <param name="nome">Nome do marcador.</param>
    /// <param name="props">Propriedades do tipo.</param>
    /// <returns>Verdadeiro se conhecido.</returns>
    private bool Conhece(string nome, Dictionary<string, PropertyInfo> props)
    {
        if (nome.StartsWith(PrefixoSistema, StringComparison.OrdinalIgnoreCase))
            return _computados.ContainsKey(nome[PrefixoSistema.Length..]);

        return props.ContainsKey(nome) || _variaveis.ContainsKey(nome);
    }

    /// <summary>Formata um valor conforme o formato informado, sempre na cultura pt-BR.</summary>
    /// <param name="valor">Valor da propriedade.</param>
    /// <param name="formato">Formato .NET ou nulo.</param>
    /// <returns>Texto formatado (vazio quando o valor é nulo).</returns>
    private static string Formatar(object? valor, string? formato)
    {
        if (valor is null)
            return string.Empty;

        if (!string.IsNullOrEmpty(formato) && valor is IFormattable formatavel)
            return formatavel.ToString(formato, Cultura);

        return Convert.ToString(valor, Cultura) ?? string.Empty;
    }

    /// <summary>Cria o conjunto padrão de valores computados (saudação, data, hora, dia da semana).</summary>
    /// <returns>Dicionário chave → função.</returns>
    private static Dictionary<string, Func<string?, CultureInfo, string>> ComputadosPadrao()
        => new(StringComparer.OrdinalIgnoreCase)
        {
            ["Saudacao"] = (_, _) => Saudacao(),
            ["DataHoje"] = (formato, cultura) =>
                DateTime.Now.ToString(string.IsNullOrEmpty(formato) ? "dd/MM/yyyy" : formato, cultura),
            ["Hora"] = (formato, cultura) =>
                DateTime.Now.ToString(string.IsNullOrEmpty(formato) ? "HH:mm" : formato, cultura),
            ["DataHora"] = (formato, cultura) =>
                DateTime.Now.ToString(string.IsNullOrEmpty(formato) ? "dd/MM/yyyy HH:mm" : formato, cultura),
            ["DiaSemana"] = (_, cultura) => cultura.DateTimeFormat.GetDayName(DateTime.Now.DayOfWeek)
        };

    /// <summary>Devolve a saudação conforme a hora atual (bom dia / boa tarde / boa noite).</summary>
    /// <returns>Texto da saudação.</returns>
    private static string Saudacao()
    {
        int hora = DateTime.Now.Hour;

        // Faixa horária não é um switch natural (é intervalo), então usa if encadeado.
        if (hora < 12)
            return "Bom dia";

        if (hora < 18)
            return "Boa tarde";

        return "Boa noite";
    }

    /// <summary>Obtém (com cache) as propriedades legíveis do tipo, indexadas por nome sem diferenciar maiúsculas.</summary>
    /// <param name="tipo">Tipo a inspecionar.</param>
    /// <returns>Dicionário nome → PropertyInfo.</returns>
    private static Dictionary<string, PropertyInfo> ObterPropriedades(Type tipo)
        => Cache.GetOrAdd(tipo, t =>
        {
            Dictionary<string, PropertyInfo> mapa = new(StringComparer.OrdinalIgnoreCase);

            foreach (PropertyInfo prop in t.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                if (prop.CanRead && prop.GetIndexParameters().Length == 0 && !mapa.ContainsKey(prop.Name))
                    mapa[prop.Name] = prop;

            return mapa;
        });

    /// <summary>Expressão que casa apenas o marcador duplo <c>{{Nome}}</c> ou <c>{{Nome:formato}}</c>.</summary>
    /// <returns>Regex compilada.</returns>
    [GeneratedRegex(@"\{\{\s*([A-Za-z_][A-Za-z0-9_.]*)\s*(?::(.+?))?\s*\}\}", RegexOptions.Compiled)]
    private static partial Regex MarcadorRegex();
}

// #endregion

// #region Extrator de campos  ->  destino sugerido: Services/

/// <summary>Monta a lista de campos disponíveis (objeto + variáveis + computados) para o menu "Adicionar propriedade".</summary>
public static class ExtratorCamposTemplate
{
    /// <summary>Lista os campos "simples" do tipo de dados, em ordem alfabética, marcados como <see cref="GrupoCampo.Objeto"/>.</summary>
    /// <param name="tipoDados">Tipo do objeto de dados (flat).</param>
    /// <returns>Campos do objeto.</returns>
    private static IReadOnlyList<CampoTemplate> ListarDoObjeto(Type tipoDados)
    {
        ArgumentNullException.ThrowIfNull(tipoDados);

        List<CampoTemplate> campos = [];

        foreach (PropertyInfo prop in tipoDados.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!prop.CanRead || prop.GetIndexParameters().Length > 0)
                continue;

            Type tipo = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

            if (!EhTipoSimples(tipo))
                continue;

            campos.Add(new CampoTemplate(prop.Name, $"{prop.Name} ({TipoAmigavel(tipo)})", GrupoCampo.Objeto));
        }

        campos.Sort((a, b) => string.Compare(a.Nome, b.Nome, StringComparison.OrdinalIgnoreCase));
        return campos;
    }

    /// <summary>Monta a lista completa de campos (objeto + variáveis + computados) a partir do tipo e do motor configurado.</summary>
    /// <param name="tipoDados">Tipo do objeto de dados.</param>
    /// <param name="motor">Motor já configurado com variáveis e computados.</param>
    /// <returns>Todos os campos disponíveis para inserção.</returns>
    public static IReadOnlyList<CampoTemplate> ListarTodos(Type tipoDados, MotorTemplateEmail motor)
        => ListarTodos(tipoDados, motor, FontesCampo.Todas);

    /// <summary>Monta a lista de campos filtrada pelas fontes permitidas naquele ponto de edição.</summary>
    /// <param name="tipoDados">Tipo do objeto de dados.</param>
    /// <param name="motor">Motor já configurado com variáveis e computados.</param>
    /// <param name="fontes">Fontes que o campo oferece (ex.: endereços não oferecem automáticos).</param>
    /// <returns>Campos disponíveis para inserção, conforme as fontes.</returns>
    public static IReadOnlyList<CampoTemplate> ListarTodos(Type tipoDados, MotorTemplateEmail motor, FontesCampo fontes)
    {
        ArgumentNullException.ThrowIfNull(motor);

        List<CampoTemplate> campos = [];

        if (fontes.HasFlag(FontesCampo.Objeto))
            campos.AddRange(ListarDoObjeto(tipoDados));

        if (fontes.HasFlag(FontesCampo.VariavelSistema))
            foreach (string nome in motor.VariaveisDisponiveis.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
                campos.Add(new CampoTemplate(nome, nome, GrupoCampo.VariavelSistema));

        if (fontes.HasFlag(FontesCampo.ValorAutomatico))
            foreach (string nome in motor.ComputadosDisponiveis.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
                campos.Add(new CampoTemplate(nome, nome, GrupoCampo.ValorAutomatico));

        return campos;
    }

    /// <summary>Indica se o tipo é "simples" o bastante para virar um marcador de texto.</summary>
    /// <param name="tipo">Tipo (já desembrulhado de Nullable).</param>
    /// <returns>Verdadeiro para texto, números, data/hora, booleano, enum e Guid.</returns>
    private static bool EhTipoSimples(Type tipo)
    {
        if (tipo.IsEnum) return true;
        if (tipo == typeof(string)) return true;
        if (tipo == typeof(Guid)) return true;
        if (tipo == typeof(DateTime) || tipo == typeof(DateTimeOffset)) return true;
        if (tipo == typeof(DateOnly) || tipo == typeof(TimeOnly) || tipo == typeof(TimeSpan)) return true;

        switch (Type.GetTypeCode(tipo))
        {
            case TypeCode.Boolean:
            case TypeCode.Byte:
            case TypeCode.SByte:
            case TypeCode.Int16:
            case TypeCode.UInt16:
            case TypeCode.Int32:
            case TypeCode.UInt32:
            case TypeCode.Int64:
            case TypeCode.UInt64:
            case TypeCode.Single:
            case TypeCode.Double:
            case TypeCode.Decimal:
                return true;
            default:
                return false;
        }
    }

    /// <summary>Devolve um rótulo amigável em pt-BR para o tipo.</summary>
    /// <param name="tipo">Tipo (já desembrulhado de Nullable).</param>
    /// <returns>Rótulo como "Texto", "Data", "Número"…</returns>
    private static string TipoAmigavel(Type tipo)
    {
        if (tipo.IsEnum) return "Opção";
        if (tipo == typeof(string) || tipo == typeof(Guid)) return "Texto";
        if (tipo == typeof(bool)) return "Sim/Não";

        if (tipo == typeof(DateTime) || tipo == typeof(DateTimeOffset) || tipo == typeof(DateOnly))
            return "Data";

        if (tipo == typeof(TimeOnly) || tipo == typeof(TimeSpan))
            return "Hora";

        switch (Type.GetTypeCode(tipo))
        {
            case TypeCode.Single:
            case TypeCode.Double:
            case TypeCode.Decimal:
                return "Número (decimal)";
            default:
                return "Número";
        }
    }
}

// #endregion

// #region Envio  ->  destino sugerido: Services/ (envio centralizado na DLL)

/// <summary>Anexo já resolvido (marcadores trocados). O envio finaliza a materialização (lê caminho, aplica exclusão).</summary>
/// <param name="NomeArquivo">Nome final do anexo.</param>
/// <param name="ContentType">Content-type, quando conhecido.</param>
/// <param name="Conteudo">Bytes (forma 1 / imagem inline); nulo quando o anexo vem de caminho.</param>
/// <param name="CaminhoResolvido">Caminho já resolvido (forma 3); nulo quando o anexo já traz bytes.</param>
/// <param name="ContentId">ContentId para imagem inline (cid:); nulo para anexo comum.</param>
/// <param name="ExcluirAposAnexar">Se o arquivo de origem deve ser excluído após anexar.</param>
public sealed record AnexoResolvido(
    string NomeArquivo,
    string? ContentType,
    byte[]? Conteudo,
    string? CaminhoResolvido,
    string? ContentId,
    bool ExcluirAposAnexar);

/// <summary>E-mail com todos os marcadores já resolvidos, pronto para o envio.</summary>
/// <param name="Assunto">Assunto final.</param>
/// <param name="Destinatarios">Destinatários (já separados).</param>
/// <param name="Cc">Cópia.</param>
/// <param name="Cco">Cópia oculta.</param>
/// <param name="Corpo">Corpo HTML final.</param>
/// <param name="Classificacao">Rótulo de sensibilidade.</param>
/// <param name="Anexos">Anexos resolvidos.</param>
/// <param name="AcaoNaFalhaDeAnexo">Política em caso de anexo que falhe.</param>
public sealed record EmailResolvido(
    string Assunto,
    IReadOnlyList<string> Destinatarios,
    IReadOnlyList<string> Cc,
    IReadOnlyList<string> Cco,
    string Corpo,
    ClassificacaoEmail Classificacao,
    IReadOnlyList<AnexoResolvido> Anexos,
    AcaoFalhaAnexo AcaoNaFalhaDeAnexo);

/// <summary>Resultado de uma tentativa de envio.</summary>
/// <param name="Sucesso">Se o e-mail foi enviado.</param>
/// <param name="Mensagem">Detalhe do erro (quando não houve sucesso) ou confirmação.</param>
public sealed record ResultadoEnvio(bool Sucesso, string? Mensagem);

/// <summary>
/// Serviço central de envio de e-mail da DLL. Cada sistema registra UMA implementação na injeção de
/// dependência (a real, com Exchange, ou o stub). O componente aciona este serviço; não fala com rede.
/// </summary>
public interface IEnviadorEmail
{
    /// <summary>Envia um e-mail já resolvido.</summary>
    /// <param name="email">E-mail com marcadores resolvidos.</param>
    /// <param name="cancelamento">Token de cancelamento.</param>
    /// <returns>Resultado do envio.</returns>
    Task<ResultadoEnvio> EnviarAsync(EmailResolvido email, CancellationToken cancelamento = default);
}

/// <summary>Implementação temporária de <see cref="IEnviadorEmail"/> — mantida para registros que ainda não trocaram para um enviador real.</summary>
/// <remarks>Prefira <see cref="EnviadorEmailExchange"/> (produção) ou <see cref="EnviadorEmailSimulado"/> (testes).</remarks>
public sealed class EnviadorEmailNaoImplementado : IEnviadorEmail
{
    /// <inheritdoc/>
    public Task<ResultadoEnvio> EnviarAsync(EmailResolvido email, CancellationToken cancelamento = default)
        => Task.FromResult(new ResultadoEnvio(
            false,
            "Envio de e-mail ainda não implementado. Registre um IEnviadorEmail real (Exchange) na injeção de dependência."));
}

/// <summary>Configurações de e-mail de um sistema (remetente, credenciais e servidor). Uma instância por sistema.</summary>
public sealed class OpcoesEmail
{
    /// <summary>Host do servidor SMTP/Exchange.</summary>
    public string Host { get; set; } = string.Empty;

    /// <summary>Porta do servidor (587 para STARTTLS, 25 para relay interno).</summary>
    public int Porta { get; set; } = 587;

    /// <summary>Se a conexão usa SSL/TLS.</summary>
    public bool UsarSsl { get; set; } = true;

    /// <summary>Endereço de e-mail do remetente (o "de").</summary>
    public string EnderecoRemetente { get; set; } = string.Empty;

    /// <summary>Nome de exibição do remetente (opcional).</summary>
    public string? NomeExibicao { get; set; }

    /// <summary>Usuário sistêmico para autenticação.</summary>
    public string Usuario { get; set; } = string.Empty;

    /// <summary>Senha do usuário sistêmico.</summary>
    public string Senha { get; set; } = string.Empty;

    /// <summary>Pasta base onde os anexos de caminho dinâmico podem residir. Confina o acesso (evita path traversal).</summary>
    public string? PastaBaseAnexos { get; set; }

    /// <summary>Pasta onde o enviador simulado grava os .eml (padrão: subpasta temporária).</summary>
    public string? PastaSimulacao { get; set; }

    /// <summary>Tempo limite do envio, em segundos.</summary>
    public int TimeoutSegundos { get; set; } = 60;
}

/// <summary>
/// Base de envio: monta a mensagem MIME a partir do <see cref="EmailResolvido"/> (destinatários, corpo HTML,
/// imagens inline por cid:, anexos das formas 1 e 3), aplica a política de falha de anexo, sanitiza caminhos e
/// trata a exclusão pós-anexo. Subclasses definem apenas a ENTREGA (SMTP real ou gravação simulada).
/// </summary>
public abstract class EnviadorEmailBase : IEnviadorEmail
{
    /// <summary>Configurações do sistema.</summary>
    protected OpcoesEmail Opcoes { get; }

    /// <summary>Cria o enviador com as opções do sistema.</summary>
    /// <param name="opcoes">Configurações de e-mail.</param>
    protected EnviadorEmailBase(OpcoesEmail opcoes)
        => Opcoes = opcoes ?? throw new ArgumentNullException(nameof(opcoes));

    /// <inheritdoc/>
    public async Task<ResultadoEnvio> EnviarAsync(EmailResolvido email, CancellationToken cancelamento = default)
    {
        ArgumentNullException.ThrowIfNull(email);

        List<string> excluirAposEnvio = [];
        MailMessage? mensagem = null;

        try
        {
            mensagem = MontarMensagem(email, excluirAposEnvio);
            await EntregarAsync(mensagem, cancelamento);

            foreach (string caminho in excluirAposEnvio)
                TentarExcluir(caminho);

            return new ResultadoEnvio(true, "E-mail enviado.");
        }
        catch (FalhaAnexoException ex)
        {
            return new ResultadoEnvio(false, ex.Message);
        }
        catch (OperationCanceledException)
        {
            return new ResultadoEnvio(false, "Envio cancelado.");
        }
        catch (Exception ex)
        {
            return new ResultadoEnvio(false, $"Falha no envio: {ex.Message}");
        }
        finally
        {
            mensagem?.Dispose();
        }
    }

    /// <summary>Entrega a mensagem já montada (definida pela subclasse).</summary>
    /// <param name="mensagem">Mensagem MIME pronta.</param>
    /// <param name="cancelamento">Token de cancelamento.</param>
    protected abstract Task EntregarAsync(MailMessage mensagem, CancellationToken cancelamento);

    /// <summary>Monta a mensagem MIME completa a partir do e-mail resolvido.</summary>
    /// <param name="email">E-mail resolvido.</param>
    /// <param name="excluirAposEnvio">Recebe os caminhos a excluir após o envio.</param>
    /// <returns>Mensagem pronta para entrega.</returns>
    private MailMessage MontarMensagem(EmailResolvido email, List<string> excluirAposEnvio)
    {
        MailMessage mensagem = new()
        {
            From = new MailAddress(Opcoes.EnderecoRemetente, Opcoes.NomeExibicao ?? Opcoes.EnderecoRemetente),
            Subject = email.Assunto,
            SubjectEncoding = Encoding.UTF8,
            BodyEncoding = Encoding.UTF8
        };

        foreach (string destino in email.Destinatarios)
            mensagem.To.Add(destino);

        foreach (string cc in email.Cc)
            mensagem.CC.Add(cc);

        foreach (string cco in email.Cco)
            mensagem.Bcc.Add(cco);

        string? sensibilidade = MapearSensibilidade(email.Classificacao);
        if (sensibilidade is not null)
            mensagem.Headers.Add("Sensitivity", sensibilidade);

        List<AnexoResolvido> inline = [.. email.Anexos.Where(a => a.ContentId is not null)];
        List<AnexoResolvido> comuns = [.. email.Anexos.Where(a => a.ContentId is null)];

        if (inline.Count > 0)
        {
            AlternateView visao = AlternateView.CreateAlternateViewFromString(
                email.Corpo, Encoding.UTF8, MediaTypeNames.Text.Html);

            foreach (AnexoResolvido anexo in inline)
            {
                byte[]? bytes = ObterBytes(anexo, email.AcaoNaFalhaDeAnexo, excluirAposEnvio);
                if (bytes is null)
                    continue;

                LinkedResource recurso = new(new MemoryStream(bytes), anexo.ContentType ?? MediaTypeNames.Application.Octet)
                {
                    ContentId = anexo.ContentId
                };
                visao.LinkedResources.Add(recurso);
            }

            mensagem.AlternateViews.Add(visao);
        }
        else
        {
            mensagem.Body = email.Corpo;
            mensagem.IsBodyHtml = true;
        }

        foreach (AnexoResolvido anexo in comuns)
        {
            byte[]? bytes = ObterBytes(anexo, email.AcaoNaFalhaDeAnexo, excluirAposEnvio);
            if (bytes is null)
                continue;

            Attachment anexoMime = new(new MemoryStream(bytes), anexo.NomeArquivo, anexo.ContentType ?? MediaTypeNames.Application.Octet);
            mensagem.Attachments.Add(anexoMime);
        }

        return mensagem;
    }

    /// <summary>Obtém os bytes de um anexo (forma 1 embutida ou forma 3 de caminho), aplicando a política de falha.</summary>
    /// <param name="anexo">Anexo resolvido.</param>
    /// <param name="politica">Política em caso de falha.</param>
    /// <param name="excluirAposEnvio">Lista que recebe o caminho a excluir, se marcado.</param>
    /// <returns>Bytes do anexo, ou nulo para pular (quando a política permite).</returns>
    private byte[]? ObterBytes(AnexoResolvido anexo, AcaoFalhaAnexo politica, List<string> excluirAposEnvio)
    {
        if (anexo.Conteudo is not null)
            return anexo.Conteudo;

        if (string.IsNullOrWhiteSpace(anexo.CaminhoResolvido))
            return TratarFalha(politica, $"anexo '{anexo.NomeArquivo}' sem conteúdo nem caminho");

        try
        {
            string caminho = CaminhoSeguro(anexo.CaminhoResolvido);
            byte[] bytes = File.ReadAllBytes(caminho);

            if (anexo.ExcluirAposAnexar)
                excluirAposEnvio.Add(caminho);

            return bytes;
        }
        catch (FalhaAnexoException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return TratarFalha(politica, $"falha ao ler o anexo '{anexo.NomeArquivo}': {ex.Message}");
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
    /// <param name="caminhoResolvido">Caminho vindo do template (já com marcadores resolvidos).</param>
    /// <returns>Caminho absoluto seguro, dentro da pasta base.</returns>
    private string CaminhoSeguro(string caminhoResolvido)
    {
        if (string.IsNullOrWhiteSpace(Opcoes.PastaBaseAnexos))
            throw new FalhaAnexoException("PastaBaseAnexos não configurada — necessária para anexos de caminho dinâmico.");

        string baseDir = Path.GetFullPath(Opcoes.PastaBaseAnexos);
        string baseComBarra = baseDir.EndsWith(Path.DirectorySeparatorChar)
            ? baseDir
            : baseDir + Path.DirectorySeparatorChar;

        string combinado = Path.GetFullPath(Path.Combine(baseDir, caminhoResolvido));

        if (!combinado.StartsWith(baseComBarra, StringComparison.OrdinalIgnoreCase))
            throw new FalhaAnexoException($"Caminho de anexo fora da pasta base permitida: {caminhoResolvido}");

        return combinado;
    }

    /// <summary>Mapeia a classificação de sensibilidade para o cabeçalho Sensitivity do Exchange.</summary>
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

    /// <summary>Tenta excluir um arquivo (melhor esforço; não propaga erro).</summary>
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
            // Exclusão é melhor esforço; não deve derrubar um envio bem-sucedido.
        }
    }

    /// <summary>Falha de anexo que deve abortar o envio (conforme a política).</summary>
    private sealed class FalhaAnexoException : Exception
    {
        /// <summary>Cria a exceção com a mensagem informada.</summary>
        /// <param name="mensagem">Descrição da falha.</param>
        public FalhaAnexoException(string mensagem) : base(mensagem)
        {
        }
    }
}

/// <summary>Enviador real via SMTP/Exchange, usando as credenciais do sistema.</summary>
public sealed class EnviadorEmailExchange : EnviadorEmailBase
{
    /// <summary>Cria o enviador com as opções do sistema.</summary>
    /// <param name="opcoes">Configurações de e-mail.</param>
    public EnviadorEmailExchange(OpcoesEmail opcoes) : base(opcoes)
    {
    }

    /// <inheritdoc/>
    protected override async Task EntregarAsync(MailMessage mensagem, CancellationToken cancelamento)
    {
        using SmtpClient cliente = new(Opcoes.Host, Opcoes.Porta)
        {
            EnableSsl = Opcoes.UsarSsl,
            Credentials = new NetworkCredential(Opcoes.Usuario, Opcoes.Senha),
            Timeout = Opcoes.TimeoutSegundos * 1000
        };

        await cliente.SendMailAsync(mensagem, cancelamento);
    }
}

/// <summary>
/// Enviador de SIMULAÇÃO: em vez de transmitir, grava o e-mail montado como arquivo .eml numa pasta
/// (via pasta de coleta do SMTP). Serve para testar o fluxo completo — inclusive anexos e imagens inline —
/// sem depender de um servidor. O .eml gerado pode ser aberto no Outlook.
/// </summary>
public sealed class EnviadorEmailSimulado : EnviadorEmailBase
{
    /// <summary>Cria o enviador simulado com as opções do sistema.</summary>
    /// <param name="opcoes">Configurações de e-mail.</param>
    public EnviadorEmailSimulado(OpcoesEmail opcoes) : base(opcoes)
    {
    }

    /// <inheritdoc/>
    protected override async Task EntregarAsync(MailMessage mensagem, CancellationToken cancelamento)
    {
        string pasta = string.IsNullOrWhiteSpace(Opcoes.PastaSimulacao)
            ? Path.Combine(Path.GetTempPath(), "bbi-emails")
            : Opcoes.PastaSimulacao;

        Directory.CreateDirectory(pasta);

        using SmtpClient cliente = new()
        {
            DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
            PickupDirectoryLocation = pasta
        };

        await cliente.SendMailAsync(mensagem, cancelamento);
    }
}

/// <summary>Persistência de templates de e-mail. Cada sistema implementa (tipicamente com EF) contra a sua tabela local.</summary>
public interface IRepositorioTemplateEmail
{
    /// <summary>Lista os nomes dos templates salvos.</summary>
    /// <param name="cancelamento">Token de cancelamento.</param>
    /// <returns>Nomes disponíveis.</returns>
    Task<IReadOnlyList<string>> ListarNomesAsync(CancellationToken cancelamento = default);

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
    /// <summary>Armazenamento por nome (case-insensitive).</summary>
    private readonly ConcurrentDictionary<string, ITemplateEmail> _dados =
        new(StringComparer.OrdinalIgnoreCase);

    /// <inheritdoc/>
    public Task<IReadOnlyList<string>> ListarNomesAsync(CancellationToken cancelamento = default)
    {
        IReadOnlyList<string> nomes = [.. _dados.Keys.OrderBy(n => n, StringComparer.OrdinalIgnoreCase)];
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

/// <summary>Resolve um <see cref="ITemplateEmail"/> (com marcadores) em um <see cref="EmailResolvido"/> pronto para envio.</summary>
public static class ResolvedorEmail
{
    /// <summary>Resolve todos os campos do template com os dados informados.</summary>
    /// <param name="template">Template com marcadores.</param>
    /// <param name="dados">Objeto de dados (pode ser nulo).</param>
    /// <param name="motor">Motor de template configurado.</param>
    /// <returns>E-mail pronto para envio.</returns>
    public static EmailResolvido Resolver(ITemplateEmail template, object? dados, MotorTemplateEmail motor)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(motor);

        string assunto = motor.Resolver(template.Assunto ?? string.Empty, dados).Texto;
        IReadOnlyList<string> para = Separar(motor.Resolver(template.Destinatarios ?? string.Empty, dados).Texto);
        IReadOnlyList<string> cc = Separar(motor.Resolver(template.Cc ?? string.Empty, dados).Texto);
        IReadOnlyList<string> cco = Separar(motor.Resolver(template.Cco ?? string.Empty, dados).Texto);
        string corpo = motor.Resolver(template.Corpo ?? string.Empty, dados).Texto;

        List<AnexoResolvido> anexos = [];

        foreach (IAnexoTemplate anexo in template.Anexos)
        {
            string nome = motor.Resolver(anexo.NomeArquivo ?? string.Empty, dados).Texto;
            string? caminho = anexo.PadraoCaminho is null
                ? null
                : motor.Resolver(anexo.PadraoCaminho, dados).Texto;

            // >>> SEGURANÇA (implementação do envio): 'caminho' contém valores vindos do objeto
            // de dados. Antes de ler o arquivo, o enviador DEVE sanitizar (remover '/', '\\', '..'
            // e caracteres inválidos) e validar que o caminho final permanece DENTRO da pasta base
            // configurada pelo dev — senão há risco de path traversal (ler arquivo de fora da pasta).
            anexos.Add(new AnexoResolvido(
                nome, anexo.ContentType, anexo.Conteudo, caminho, anexo.ContentId, anexo.ExcluirAposAnexar));
        }

        return new EmailResolvido(
            assunto, para, cc, cco, corpo, template.Classificacao, anexos, template.AcaoNaFalhaDeAnexo);
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
