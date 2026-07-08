// =============================================================================
//  BBiCore.cs  —  Biblioteca de componentes BBiCore (RCL, .NET 9) — CÓDIGO C#
// -----------------------------------------------------------------------------
//  Arquivo ÚNICO com todo o código C# da biblioteca, propositalmente consolidado
//  para facilitar a cópia manual. Os componentes visuais ficam nos .razro
//  (Razor não se funde com .cs). Três áreas, em namespaces distintos:
//
//    BBiCore.Importacao  -> importador tabular (CSV/Excel) + despivotador
//    BBiCore.Componentes -> tipos de apoio dos componentes (enums/records)
//    BBiCore.Email       -> motor de template, contrato e envio de e-mail
//
//  >>> NOTA PARA A INSTÂNCIA QUE FOR REORGANIZAR <<<
//  Cada tipo está isolado em uma #region nomeada pelo destino sugerido. Para
//  dividir no padrão do time, recorte região por região preservando os
//  namespaces. Nada aqui depende de estar no mesmo arquivo. As áreas de
//  Importacao e Email são independentes entre si e podem virar pastas separadas.
// =============================================================================

using System.Collections.Concurrent;
using System.Globalization;
using System.Linq.Expressions;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;

namespace BBiCore.Importacao
{

// #region Enums e opções  ->  destino sugerido: Enums/ e Options/ (ou Models/)

/// <summary>Define como o motor reage a falhas durante a importação.</summary>
public enum ModoImportacao
{
    /// <summary>Processa o arquivo inteiro e reporta todas as falhas encontradas.</summary>
    Completo,

    /// <summary>Encerra a importação assim que a primeira falha ocorre.</summary>
    PararNaPrimeiraFalha
}

/// <summary>Controla como linhas em branco na região de dados são tratadas.</summary>
public enum TratamentoLinhasEmBranco
{
    /// <summary>Ignora qualquer linha em branco, em qualquer posição (padrão).</summary>
    Todas,

    /// <summary>Ignora linhas em branco apenas antes da primeira linha de dados; a partir daí, uma linha em branco é tratada como dado.</summary>
    SomenteIniciais,

    /// <summary>Não ignora nenhuma linha em branco; toda linha vazia é tratada como dado.</summary>
    Nao
}

/// <summary>Situação de uma linha processada.</summary>
public enum StatusLinha
{
    /// <summary>Todas as colunas (e o callback de linha) converteram com sucesso.</summary>
    Ok,

    /// <summary>Pelo menos uma coluna ou o callback de linha falhou.</summary>
    Falha
}

/// <summary>Opções que valem para a importação inteira (não para uma coluna específica).</summary>
public sealed class OpcoesImportacao
{
    /// <summary>Cultura padrão das colunas que não definem a própria. Padrão: pt-BR.</summary>
    public CultureInfo CulturaGlobal { get; set; } = CultureInfo.GetCultureInfo("pt-BR");

    /// <summary>Comportamento diante de falhas. Padrão: <see cref="ModoImportacao.Completo"/>.</summary>
    public ModoImportacao Modo { get; set; } = ModoImportacao.Completo;

    /// <summary>
    /// Linha física do cabeçalho (1 = primeira linha). Nulo autodetecta: usa a primeira linha
    /// não-vazia. Ignorado quando o mapa é sem cabeçalho.
    /// </summary>
    public int? LinhaCabecalho { get; set; }

    /// <summary>
    /// Linha física onde os dados começam (1 = primeira linha). Nulo faz os dados começarem logo
    /// após o cabeçalho. Use para o caso "cabeçalho na linha 2, dados na linha 10".
    /// </summary>
    public int? LinhaInicioDados { get; set; }

    /// <summary>
    /// Coluna onde a tabela começa (1 = primeira coluna). Recorta as colunas à esquerda, tanto do
    /// cabeçalho quanto dos dados. As posições de <c>ColunaPosicao</c> passam a ser relativas a
    /// este início. Padrão: 1 (sem recorte).
    /// </summary>
    public int ColunaInicio { get; set; } = 1;

    /// <summary>Tratamento de linhas em branco na região de dados. Padrão: <see cref="TratamentoLinhasEmBranco.Todas"/>.</summary>
    public TratamentoLinhasEmBranco LinhasEmBranco { get; set; } = TratamentoLinhasEmBranco.Todas;
}

// #endregion

// #region Resultado  ->  destino sugerido: Models/ (ou Results/)

/// <summary>Descreve uma falha ocorrida em uma coluna (ou no callback de linha).</summary>
/// <param name="Origem">Coluna de origem no arquivo (nome ou posição). Nulo quando a falha é da linha inteira.</param>
/// <param name="Destino">Propriedade de destino no objeto.</param>
/// <param name="ValorBruto">Valor lido do arquivo que causou a falha, quando aplicável.</param>
/// <param name="Mensagem">Mensagem descritiva da falha.</param>
public sealed record FalhaImportacao(string? Origem, string Destino, string? ValorBruto, string Mensagem);

/// <summary>Resultado do processamento de uma única linha do arquivo.</summary>
/// <typeparam name="T">Tipo da entidade produzida.</typeparam>
public sealed class LinhaResultado<T>
{
    /// <summary>Nome do arquivo de origem.</summary>
    public required string Arquivo { get; init; }

    /// <summary>Número FÍSICO da linha no arquivo (linhas puladas e em branco também são contadas).</summary>
    public required int Linha { get; init; }

    /// <summary>Situação da linha.</summary>
    public StatusLinha Status { get; init; }

    /// <summary>Entidade produzida. Preenchida apenas quando <see cref="Status"/> é <see cref="StatusLinha.Ok"/>.</summary>
    public T? Item { get; init; }

    /// <summary>Falhas da linha. Vazia quando a linha é Ok.</summary>
    public List<FalhaImportacao> Falhas { get; init; } = [];

    /// <summary>Indica se a linha foi processada com sucesso.</summary>
    public bool Ok => Status == StatusLinha.Ok;
}

/// <summary>Resultado consolidado de uma importação: relatório por linha e lista tipada.</summary>
/// <typeparam name="T">Tipo da entidade produzida.</typeparam>
public sealed class ResultadoImportacao<T>
{
    /// <summary>Relatório com uma entrada por linha de dados lida (Ok ou Falha).</summary>
    public List<LinhaResultado<T>> Relatorio { get; init; } = [];

    /// <summary>Itens importados com sucesso — projeção das linhas Ok do relatório.</summary>
    public List<T> Itens => Relatorio.Where(l => l.Ok && l.Item is not null).Select(l => l.Item!).ToList();

    /// <summary>Total de linhas de dados processadas.</summary>
    public int Total => Relatorio.Count;

    /// <summary>Quantidade de linhas Ok.</summary>
    public int TotalOk => Relatorio.Count(l => l.Ok);

    /// <summary>Quantidade de linhas com falha.</summary>
    public int TotalFalha => Relatorio.Count(l => !l.Ok);

    /// <summary>Verdadeiro quando não houve nenhuma falha.</summary>
    public bool Sucesso => TotalFalha == 0;

    /// <summary>Permite desestruturar em (relatório, itens).</summary>
    /// <param name="relatorio">Recebe o relatório completo.</param>
    /// <param name="itens">Recebe a lista de itens Ok.</param>
    public void Deconstruct(out List<LinhaResultado<T>> relatorio, out List<T> itens)
    {
        relatorio = Relatorio;
        itens = Itens;
    }
}

// #endregion

// #region Contratos  ->  destino sugerido: Interfaces/ (ou Abstractions/)

/// <summary>Lê um arquivo tabular como sequência de linhas de células com valor nativo.</summary>
public interface ILeitorTabular
{
    /// <summary>Indica se este leitor suporta o arquivo informado (pela extensão).</summary>
    /// <param name="nomeArquivo">Nome do arquivo, com extensão.</param>
    /// <returns>Verdadeiro se o formato é suportado.</returns>
    bool Suporta(string nomeArquivo);

    /// <summary>Lê o conteúdo do stream, devolvendo cada linha como um vetor de células.</summary>
    /// <param name="stream">Fluxo do arquivo.</param>
    /// <returns>Sequência de linhas; cada linha é um vetor de valores nativos.</returns>
    IEnumerable<object?[]> LerLinhas(Stream stream);
}

/// <summary>Acesso às células cruas de uma linha, entregue ao callback de linha.</summary>
public interface IAcessorLinha
{
    /// <summary>Lê uma célula pelo nome da coluna (exige cabeçalho).</summary>
    /// <param name="nome">Nome da coluna.</param>
    /// <returns>Texto da célula, ou nulo se a coluna não existir.</returns>
    string? PorNome(string nome);

    /// <summary>Lê uma célula pela posição (0-based).</summary>
    /// <param name="indice">Índice da coluna.</param>
    /// <returns>Texto da célula, ou nulo se fora do intervalo.</returns>
    string? PorPosicao(int indice);
}

// #endregion

// #region Mapa  ->  destino sugerido: Mapping/ (ColunaMapeada é internal)

/// <summary>Definição interna de uma coluna do mapa. Construída via <see cref="MapaImportacao{T}"/>; não é exposta ao consumidor.</summary>
/// <typeparam name="T">Tipo da entidade de destino.</typeparam>
internal sealed class ColunaMapeada<T>
{
    /// <summary>Ação que grava o valor convertido na propriedade de destino.</summary>
    public required Action<T, object?> Setter { get; init; }

    /// <summary>Nome da propriedade de destino.</summary>
    public required string Destino { get; init; }

    /// <summary>Tipo da propriedade de destino.</summary>
    public required Type TipoDestino { get; init; }

    /// <summary>Nomes de coluna aceitos na origem (quando o mapeamento é por nome).</summary>
    public IReadOnlyList<string> NomesOrigem { get; init; } = [];

    /// <summary>Posição da coluna na origem (quando o mapeamento é por índice, 0-based).</summary>
    public int? PosicaoOrigem { get; init; }

    /// <summary>Indica se um valor vazio nesta coluna gera falha.</summary>
    public bool Obrigatoria { get; init; }

    /// <summary>Cultura específica desta coluna; sobrescreve a global quando informada.</summary>
    public CultureInfo? Cultura { get; init; }

    /// <summary>Formato explícito de conversão (ex.: data), quando informado.</summary>
    public string? Formato { get; init; }

    /// <summary>Indica se há um valor padrão para célula vazia (coluna não obrigatória).</summary>
    public bool TemValorPadrao { get; init; }

    /// <summary>Valor padrão aplicado quando a célula vem vazia.</summary>
    public object? ValorPadrao { get; init; }

    /// <summary>Conversor próprio da coluna; quando presente, substitui a conversão padrão.</summary>
    public Func<string, CultureInfo, object?>? Callback { get; init; }

    /// <summary>Descrição legível da origem (nomes aceitos ou posição), usada no relatório de falhas.</summary>
    public string Origem =>
        NomesOrigem.Count > 0 ? string.Join(" / ", NomesOrigem) : $"posição {PosicaoOrigem}";
}

/// <summary>Mapa de importação: descreve como cada coluna do arquivo alimenta uma propriedade de <typeparamref name="T"/>.</summary>
/// <typeparam name="T">Tipo da entidade produzida (precisa de construtor sem parâmetros).</typeparam>
public sealed class MapaImportacao<T> where T : new()
{
    /// <summary>Colunas configuradas, na ordem de declaração.</summary>
    private readonly List<ColunaMapeada<T>> _colunas = [];

    /// <summary>Indica se o arquivo possui linha de cabeçalho.</summary>
    public bool TemCabecalho { get; }

    /// <summary>Colunas mapeadas (uso interno do motor).</summary>
    internal IReadOnlyList<ColunaMapeada<T>> Colunas => _colunas;

    /// <summary>Callback executado ao final de cada linha Ok (uso interno do motor).</summary>
    internal Action<T, IAcessorLinha>? CallbackLinha { get; private set; }

    /// <summary>Construtor privado; use <see cref="ComCabecalho"/> ou <see cref="SemCabecalho"/>.</summary>
    /// <param name="temCabecalho">Se o arquivo tem cabeçalho.</param>
    private MapaImportacao(bool temCabecalho) => TemCabecalho = temCabecalho;

    /// <summary>Cria um mapa para arquivo COM linha de cabeçalho (colunas por nome).</summary>
    /// <returns>Novo mapa.</returns>
    public static MapaImportacao<T> ComCabecalho() => new(true);

    /// <summary>Cria um mapa para arquivo SEM cabeçalho (colunas por posição).</summary>
    /// <returns>Novo mapa.</returns>
    public static MapaImportacao<T> SemCabecalho() => new(false);

    /// <summary>Mapeia uma propriedade a partir de uma coluna identificada por nome.</summary>
    /// <typeparam name="TProp">Tipo da propriedade de destino.</typeparam>
    /// <param name="destino">Propriedade de destino (ex.: <c>x =&gt; x.Nome</c>).</param>
    /// <param name="nome">Nome da coluna no cabeçalho.</param>
    /// <param name="obrigatoria">Se verdadeiro, valor vazio gera falha.</param>
    /// <param name="cultura">Cultura desta coluna; sobrescreve a global.</param>
    /// <param name="formato">Formato explícito (ex.: <c>"dd/MM/yyyy"</c>).</param>
    /// <param name="valorPadrao">Valor usado quando a célula vem vazia (coluna não obrigatória).</param>
    /// <param name="callback">Conversor próprio desta coluna; quando informado, substitui a conversão padrão.</param>
    /// <returns>O próprio mapa, para encadeamento.</returns>
    public MapaImportacao<T> Coluna<TProp>(
        Expression<Func<T, TProp>> destino,
        string nome,
        bool obrigatoria = false,
        CultureInfo? cultura = null,
        string? formato = null,
        object? valorPadrao = null,
        Func<string, CultureInfo, object?>? callback = null)
        => AdicionarNome(destino, [nome], obrigatoria, cultura, formato, valorPadrao, callback);

    /// <summary>Mapeia uma propriedade aceitando vários nomes possíveis de coluna (cabeçalho variável).</summary>
    /// <typeparam name="TProp">Tipo da propriedade de destino.</typeparam>
    /// <param name="destino">Propriedade de destino.</param>
    /// <param name="nomesAceitos">Nomes aceitos, em ordem de preferência.</param>
    /// <param name="obrigatoria">Se verdadeiro, valor vazio gera falha.</param>
    /// <param name="cultura">Cultura desta coluna; sobrescreve a global.</param>
    /// <param name="formato">Formato explícito.</param>
    /// <param name="valorPadrao">Valor usado quando a célula vem vazia (coluna não obrigatória).</param>
    /// <param name="callback">Conversor próprio desta coluna.</param>
    /// <returns>O próprio mapa, para encadeamento.</returns>
    public MapaImportacao<T> Coluna<TProp>(
        Expression<Func<T, TProp>> destino,
        string[] nomesAceitos,
        bool obrigatoria = false,
        CultureInfo? cultura = null,
        string? formato = null,
        object? valorPadrao = null,
        Func<string, CultureInfo, object?>? callback = null)
        => AdicionarNome(destino, nomesAceitos, obrigatoria, cultura, formato, valorPadrao, callback);

    /// <summary>Mapeia uma propriedade a partir da posição da coluna (0-based, relativa a <see cref="OpcoesImportacao.ColunaInicio"/>).</summary>
    /// <typeparam name="TProp">Tipo da propriedade de destino.</typeparam>
    /// <param name="destino">Propriedade de destino.</param>
    /// <param name="posicao">Índice da coluna (0-based).</param>
    /// <param name="obrigatoria">Se verdadeiro, valor vazio gera falha.</param>
    /// <param name="cultura">Cultura desta coluna; sobrescreve a global.</param>
    /// <param name="formato">Formato explícito.</param>
    /// <param name="valorPadrao">Valor usado quando a célula vem vazia (coluna não obrigatória).</param>
    /// <param name="callback">Conversor próprio desta coluna.</param>
    /// <returns>O próprio mapa, para encadeamento.</returns>
    public MapaImportacao<T> ColunaPosicao<TProp>(
        Expression<Func<T, TProp>> destino,
        int posicao,
        bool obrigatoria = false,
        CultureInfo? cultura = null,
        string? formato = null,
        object? valorPadrao = null,
        Func<string, CultureInfo, object?>? callback = null)
    {
        (Action<T, object?> setter, string nome, Type tipo) = Compilar(destino);
        _colunas.Add(new ColunaMapeada<T>
        {
            Setter = setter, Destino = nome, TipoDestino = tipo,
            PosicaoOrigem = posicao, Obrigatoria = obrigatoria,
            Cultura = cultura, Formato = formato,
            TemValorPadrao = valorPadrao is not null, ValorPadrao = valorPadrao,
            Callback = callback
        });
        return this;
    }

    /// <summary>Define um callback executado após todas as colunas (apenas quando a linha está Ok). Recebe a entidade e um acessor às células cruas — resolve origem composta e ajustes finais.</summary>
    /// <param name="callback">Ação a executar sobre a entidade já preenchida.</param>
    /// <returns>O próprio mapa, para encadeamento.</returns>
    public MapaImportacao<T> ComCallbackLinha(Action<T, IAcessorLinha> callback)
    {
        CallbackLinha = callback;
        return this;
    }

    /// <summary>Adiciona uma coluna mapeada por nome (implementação comum das sobrecargas de <c>Coluna</c>).</summary>
    /// <typeparam name="TProp">Tipo da propriedade de destino.</typeparam>
    /// <param name="destino">Expressão da propriedade.</param>
    /// <param name="nomes">Nomes aceitos.</param>
    /// <param name="obrigatoria">Se é obrigatória.</param>
    /// <param name="cultura">Cultura específica.</param>
    /// <param name="formato">Formato explícito.</param>
    /// <param name="valorPadrao">Valor padrão.</param>
    /// <param name="callback">Conversor próprio.</param>
    /// <returns>O próprio mapa.</returns>
    private MapaImportacao<T> AdicionarNome<TProp>(
        Expression<Func<T, TProp>> destino, string[] nomes, bool obrigatoria,
        CultureInfo? cultura, string? formato, object? valorPadrao,
        Func<string, CultureInfo, object?>? callback)
    {
        if (!TemCabecalho)
            throw new InvalidOperationException(
                "Mapeamento por nome exige um mapa com cabeçalho (use ComCabecalho()).");

        (Action<T, object?> setter, string nome, Type tipo) = Compilar(destino);
        _colunas.Add(new ColunaMapeada<T>
        {
            Setter = setter, Destino = nome, TipoDestino = tipo,
            NomesOrigem = nomes, Obrigatoria = obrigatoria,
            Cultura = cultura, Formato = formato,
            TemValorPadrao = valorPadrao is not null, ValorPadrao = valorPadrao,
            Callback = callback
        });
        return this;
    }

    /// <summary>Compila uma expressão de propriedade em um setter fortemente tipado e extrai nome e tipo.</summary>
    /// <typeparam name="TProp">Tipo da propriedade.</typeparam>
    /// <param name="expr">Expressão que aponta para a propriedade.</param>
    /// <returns>Tupla com o setter, o nome e o tipo da propriedade.</returns>
    private static (Action<T, object?>, string, Type) Compilar<TProp>(Expression<Func<T, TProp>> expr)
    {
        MemberExpression? membro = expr.Body as MemberExpression
            ?? (expr.Body as UnaryExpression)?.Operand as MemberExpression;

        if (membro is null || membro.Member is not PropertyInfo prop)
            throw new ArgumentException("O destino deve apontar para uma propriedade.", nameof(expr));

        ParameterExpression alvo = Expression.Parameter(typeof(T), "alvo");
        ParameterExpression valor = Expression.Parameter(typeof(object), "valor");
        BinaryExpression corpo = Expression.Assign(
            Expression.Property(alvo, prop),
            Expression.Convert(valor, prop.PropertyType));

        Action<T, object?> setter = Expression
            .Lambda<Action<T, object?>>(corpo, alvo, valor)
            .Compile();

        return (setter, prop.Name, prop.PropertyType);
    }
}

// #endregion

// #region Conversão  ->  destino sugerido: Conversion/ (ConversorValor é internal)

/// <summary>Converte valores brutos (texto do CSV ou nativos do Excel) para o tipo de destino. Uso interno do motor.</summary>
internal sealed class ConversorValor
{
    /// <summary>Converte um valor bruto para o tipo de destino, respeitando cultura e formato.</summary>
    /// <param name="bruto">Valor lido do arquivo (string do CSV ou nativo do Excel).</param>
    /// <param name="tipoAlvo">Tipo de destino (pode ser anulável).</param>
    /// <param name="cultura">Cultura para parse de texto.</param>
    /// <param name="formato">Formato explícito (ex.: data), quando houver.</param>
    /// <returns>Valor convertido, ou nulo quando o bruto é vazio.</returns>
    public object? Converter(object? bruto, Type tipoAlvo, CultureInfo cultura, string? formato)
    {
        if (bruto is null)
            return null;

        Type tipo = Nullable.GetUnderlyingType(tipoAlvo) ?? tipoAlvo;

        // 1) Já é do tipo destino (valor tipado do Excel: DateTime, double, bool…).
        if (tipo.IsInstanceOfType(bruto))
            return bruto;

        // 2) Nativo não-string (Excel) → converte sem passar por texto/cultura.
        if (bruto is not string)
        {
            if (tipo == typeof(string))
                return Convert.ToString(bruto, cultura);

            if (bruto is IConvertible && (EhNumerico(tipo) || tipo == typeof(DateTime) || tipo == typeof(bool)))
                return Convert.ChangeType(bruto, tipo, CultureInfo.InvariantCulture);
        }

        // 3) Texto (CSV) → parseia com cultura + formato.
        string texto = (bruto as string ?? bruto.ToString() ?? string.Empty).Trim();
        if (texto.Length == 0)
            return null;

        return ConverterTexto(texto, tipo, cultura, formato);
    }

    /// <summary>Converte um texto já validado como não-vazio para o tipo de destino.</summary>
    /// <param name="texto">Texto de origem.</param>
    /// <param name="tipo">Tipo de destino (não anulável).</param>
    /// <param name="cultura">Cultura para parse.</param>
    /// <param name="formato">Formato explícito, quando houver.</param>
    /// <returns>Valor convertido.</returns>
    private static object? ConverterTexto(string texto, Type tipo, CultureInfo cultura, string? formato)
    {
        if (tipo == typeof(string)) return texto;
        if (tipo.IsEnum) return Enum.Parse(tipo, texto, ignoreCase: true);
        if (tipo == typeof(Guid)) return Guid.Parse(texto);
        if (tipo == typeof(bool)) return ConverterBool(texto);

        if (tipo == typeof(DateTime))
        {
            if (!string.IsNullOrEmpty(formato))
                return DateTime.ParseExact(texto, formato, cultura, DateTimeStyles.None);
            return DateTime.Parse(texto, cultura, DateTimeStyles.None);
        }

        switch (Type.GetTypeCode(tipo))
        {
            case TypeCode.Int16:
                return short.Parse(texto, NumberStyles.Integer, cultura);
            case TypeCode.Int32:
                return int.Parse(texto, NumberStyles.Integer, cultura);
            case TypeCode.Int64:
                return long.Parse(texto, NumberStyles.Integer, cultura);
            case TypeCode.Decimal:
                return decimal.Parse(texto, NumberStyles.Number, cultura);
            case TypeCode.Double:
                return double.Parse(texto, NumberStyles.Number, cultura);
            case TypeCode.Single:
                return float.Parse(texto, NumberStyles.Number, cultura);
            default:
                return Convert.ChangeType(texto, tipo, cultura);
        }
    }

    /// <summary>Converte texto em booleano aceitando formas em pt-BR (sim/não, s/n, 1/0, verdadeiro/falso).</summary>
    /// <param name="bruto">Texto de origem.</param>
    /// <returns>Booleano correspondente.</returns>
    private static bool ConverterBool(string bruto)
    {
        switch (bruto.Trim().ToLowerInvariant())
        {
            case "1":
            case "s":
            case "sim":
            case "true":
            case "verdadeiro":
                return true;
            case "0":
            case "n":
            case "nao":
            case "não":
            case "false":
            case "falso":
                return false;
            default:
                return bool.Parse(bruto);
        }
    }

    /// <summary>Indica se o tipo é numérico (inteiro ou de ponto flutuante).</summary>
    /// <param name="tipo">Tipo a testar.</param>
    /// <returns>Verdadeiro para tipos numéricos.</returns>
    private static bool EhNumerico(Type tipo)
    {
        switch (Type.GetTypeCode(tipo))
        {
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
}

// #endregion

// #region Motor  ->  destino sugerido: Services/ (ImportadorTabular)

/// <summary>Motor de importação: lê, mapeia e converte linhas em entidades tipadas. .NET puro, sem dependência de UI.</summary>
public static class ImportadorTabular
{
    /// <summary>Importa o conteúdo de um stream em uma lista tipada, com relatório por linha.</summary>
    /// <typeparam name="T">Tipo da entidade produzida.</typeparam>
    /// <param name="stream">Fluxo do arquivo.</param>
    /// <param name="nomeArquivo">Nome do arquivo (usado no relatório).</param>
    /// <param name="mapa">Mapa de colunas.</param>
    /// <param name="leitor">Leitor compatível com o formato.</param>
    /// <param name="opcoes">Opções da importação. Se nulo, usa os padrões.</param>
    /// <returns>Resultado consolidado com relatório e itens.</returns>
    public static ResultadoImportacao<T> Importar<T>(
        Stream stream,
        string nomeArquivo,
        MapaImportacao<T> mapa,
        ILeitorTabular leitor,
        OpcoesImportacao? opcoes = null)
        where T : new()
    {
        opcoes ??= new OpcoesImportacao();
        ConversorValor conversor = new();
        ResultadoImportacao<T> resultado = new();
        int colunaInicio = opcoes.ColunaInicio;

        using IEnumerator<object?[]> e = leitor.LerLinhas(stream).GetEnumerator();

        // Contador FÍSICO: incrementa em TODA linha crua. A "linha" do relatório é sempre a
        // posição real no arquivo — mesmo com âncoras de início ou linhas em branco no meio.
        int numeroLinha = 0;
        object?[] atual = [];

        // Avança o enumerador uma linha, atualizando o contador físico e a linha atual.
        bool Avancar()
        {
            if (e.MoveNext())
            {
                numeroLinha++;
                atual = e.Current;
                return true;
            }

            atual = [];
            return false;
        }

        Dictionary<string, int>? cabecalho = null;

        if (mapa.TemCabecalho)
        {
            object?[] linhaCabecalho = [];
            bool achou = false;

            if (opcoes.LinhaCabecalho is int lc)
            {
                // Âncora explícita: o cabeçalho está exatamente nesta linha física.
                while (numeroLinha < lc)
                    if (!Avancar())
                        return resultado;

                linhaCabecalho = Recortar(atual, colunaInicio);
                achou = true;
            }
            else
            {
                // Autodetect: primeira linha não-vazia (já recortada) é o cabeçalho.
                while (Avancar())
                {
                    object?[] rec = Recortar(atual, colunaInicio);
                    if (!LinhaVazia(rec))
                    {
                        linhaCabecalho = rec;
                        achou = true;
                        break;
                    }
                }
            }

            if (!achou)
                return resultado;

            cabecalho = MontarCabecalho(linhaCabecalho);

            List<FalhaImportacao> estruturais = mapa.Colunas
                .Where(c => c.Obrigatoria && c.PosicaoOrigem is null)
                .Where(c => !c.NomesOrigem.Any(n => cabecalho.ContainsKey(n.Trim())))
                .Select(c => new FalhaImportacao(c.Origem, c.Destino, null,
                    "Coluna obrigatória não encontrada no cabeçalho."))
                .ToList();

            if (estruturais.Count > 0)
            {
                resultado.Relatorio.Add(new LinhaResultado<T>
                {
                    Arquivo = nomeArquivo,
                    Linha = numeroLinha, // linha FÍSICA do cabeçalho
                    Status = StatusLinha.Falha,
                    Falhas = estruturais
                });
                return resultado;
            }

            // Posiciona nos dados, se houver âncora explícita.
            if (opcoes.LinhaInicioDados is int ld)
                while (numeroLinha < ld - 1)
                    if (!Avancar())
                        return resultado;
        }
        else if (opcoes.LinhaInicioDados is int ld)
        {
            // Sem cabeçalho: apenas posiciona o início dos dados.
            while (numeroLinha < ld - 1)
                if (!Avancar())
                    return resultado;
        }

        // Dados.
        bool jaViuDado = false;

        while (Avancar())
        {
            object?[] celulas = Recortar(atual, colunaInicio);

            if (LinhaVazia(celulas))
            {
                bool tratarComoDado = false;

                switch (opcoes.LinhasEmBranco)
                {
                    case TratamentoLinhasEmBranco.Todas:
                        continue;
                    case TratamentoLinhasEmBranco.SomenteIniciais:
                        if (!jaViuDado)
                            continue;
                        tratarComoDado = true;
                        break;
                    case TratamentoLinhasEmBranco.Nao:
                        tratarComoDado = true;
                        break;
                }

                if (!tratarComoDado)
                    continue;
            }
            else
            {
                jaViuDado = true;
            }

            LinhaResultado<T> linhaResultado = ProcessarLinha(
                celulas, numeroLinha, nomeArquivo, mapa, cabecalho, conversor, opcoes);

            resultado.Relatorio.Add(linhaResultado);

            if (opcoes.Modo == ModoImportacao.PararNaPrimeiraFalha && !linhaResultado.Ok)
                break;
        }

        return resultado;
    }

    /// <summary>Recorta as colunas à esquerda conforme <see cref="OpcoesImportacao.ColunaInicio"/>.</summary>
    /// <param name="celulas">Linha original.</param>
    /// <param name="colunaInicio">Coluna inicial (1-based).</param>
    /// <returns>Linha recortada (ou a original se o início é 1).</returns>
    private static object?[] Recortar(object?[] celulas, int colunaInicio)
    {
        if (colunaInicio <= 1)
            return celulas;

        int corte = colunaInicio - 1;

        if (corte >= celulas.Length)
            return [];

        return celulas[corte..];
    }

    /// <summary>Processa uma linha de dados, aplicando o mapa e produzindo a entidade ou as falhas.</summary>
    /// <typeparam name="T">Tipo da entidade.</typeparam>
    /// <param name="celulas">Células da linha (já recortadas).</param>
    /// <param name="numeroLinha">Número físico da linha.</param>
    /// <param name="nomeArquivo">Nome do arquivo.</param>
    /// <param name="mapa">Mapa de colunas.</param>
    /// <param name="cabecalho">Índice nome→posição, quando há cabeçalho.</param>
    /// <param name="conversor">Conversor de valores.</param>
    /// <param name="opcoes">Opções da importação.</param>
    /// <returns>Resultado da linha (Ok ou Falha).</returns>
    private static LinhaResultado<T> ProcessarLinha<T>(
        object?[] celulas, int numeroLinha, string nomeArquivo,
        MapaImportacao<T> mapa, Dictionary<string, int>? cabecalho,
        ConversorValor conversor, OpcoesImportacao opcoes)
        where T : new()
    {
        T entidade = new();
        List<FalhaImportacao> falhas = [];

        foreach (ColunaMapeada<T> col in mapa.Colunas)
        {
            CultureInfo cultura = col.Cultura ?? opcoes.CulturaGlobal;
            int indice = ResolverIndice(col, cabecalho);
            object? bruto = indice >= 0 && indice < celulas.Length ? celulas[indice] : null;

            bool vazio = bruto is null || (bruto is string s && string.IsNullOrWhiteSpace(s));

            if (vazio)
            {
                if (col.Obrigatoria)
                    falhas.Add(new FalhaImportacao(col.Origem, col.Destino, null, "Campo obrigatório vazio."));
                else if (col.TemValorPadrao)
                    col.Setter(entidade, col.ValorPadrao);

                continue;
            }

            try
            {
                object? valor;

                if (col.Callback is not null)
                {
                    string texto = Convert.ToString(bruto, cultura) ?? string.Empty;
                    valor = col.Callback(texto, cultura);
                }
                else
                {
                    valor = conversor.Converter(bruto, col.TipoDestino, cultura, col.Formato);
                }

                if (valor is not null)
                    col.Setter(entidade, valor);
            }
            catch (Exception ex)
            {
                string? brutoTexto = Convert.ToString(bruto, cultura);
                falhas.Add(new FalhaImportacao(col.Origem, col.Destino, brutoTexto, ex.Message));
            }
        }

        if (falhas.Count == 0 && mapa.CallbackLinha is not null)
        {
            try
            {
                IAcessorLinha acessor = new AcessorLinha(celulas, cabecalho, opcoes.CulturaGlobal);
                mapa.CallbackLinha(entidade, acessor);
            }
            catch (Exception ex)
            {
                falhas.Add(new FalhaImportacao(null, "(linha)", null, ex.Message));
            }
        }

        return falhas.Count == 0
            ? new LinhaResultado<T> { Arquivo = nomeArquivo, Linha = numeroLinha, Status = StatusLinha.Ok, Item = entidade }
            : new LinhaResultado<T> { Arquivo = nomeArquivo, Linha = numeroLinha, Status = StatusLinha.Falha, Falhas = falhas };
    }

    /// <summary>Resolve o índice físico de uma coluna, por posição fixa ou por nome no cabeçalho.</summary>
    /// <typeparam name="T">Tipo da entidade.</typeparam>
    /// <param name="col">Coluna mapeada.</param>
    /// <param name="cabecalho">Índice nome→posição, quando há cabeçalho.</param>
    /// <returns>Índice da coluna, ou -1 se não encontrada.</returns>
    private static int ResolverIndice<T>(ColunaMapeada<T> col, Dictionary<string, int>? cabecalho)
    {
        if (col.PosicaoOrigem is int pos)
            return pos;

        if (cabecalho is not null)
            foreach (string nome in col.NomesOrigem)
                if (cabecalho.TryGetValue(nome.Trim(), out int idx))
                    return idx;

        return -1;
    }

    /// <summary>Monta o índice nome→posição a partir da linha de cabeçalho.</summary>
    /// <param name="linha">Linha de cabeçalho.</param>
    /// <returns>Dicionário insensível a maiúsculas/minúsculas.</returns>
    private static Dictionary<string, int> MontarCabecalho(object?[] linha)
    {
        Dictionary<string, int> mapa = new(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < linha.Length; i++)
        {
            string nome = (linha[i]?.ToString() ?? string.Empty).Trim();
            if (nome.Length > 0 && !mapa.ContainsKey(nome))
                mapa[nome] = i;
        }

        return mapa;
    }

    /// <summary>Indica se todas as células da linha estão vazias.</summary>
    /// <param name="celulas">Células da linha.</param>
    /// <returns>Verdadeiro se a linha está em branco.</returns>
    private static bool LinhaVazia(object?[] celulas)
        => celulas.All(c => c is null || (c is string s && string.IsNullOrWhiteSpace(s)));

    /// <summary>Implementação de <see cref="IAcessorLinha"/> sobre as células cruas de uma linha.</summary>
    /// <param name="celulas">Células da linha.</param>
    /// <param name="cabecalho">Índice nome→posição, quando há cabeçalho.</param>
    /// <param name="cultura">Cultura usada para converter célula em texto.</param>
    private sealed class AcessorLinha(
        object?[] celulas, Dictionary<string, int>? cabecalho, CultureInfo cultura) : IAcessorLinha
    {
        /// <summary>Lê uma célula pela posição (0-based).</summary>
        /// <param name="indice">Índice da coluna.</param>
        /// <returns>Texto da célula, ou nulo se fora do intervalo.</returns>
        public string? PorPosicao(int indice)
            => indice >= 0 && indice < celulas.Length ? Convert.ToString(celulas[indice], cultura) : null;

        /// <summary>Lê uma célula pelo nome da coluna.</summary>
        /// <param name="nome">Nome da coluna.</param>
        /// <returns>Texto da célula, ou nulo se não existir.</returns>
        public string? PorNome(string nome)
            => cabecalho is not null && cabecalho.TryGetValue(nome.Trim(), out int idx) ? PorPosicao(idx) : null;
    }
}

// #endregion

// #region Leitores  ->  destino sugerido: Readers/ (um arquivo por leitor)

/// <summary>Leitor de arquivos CSV/TXT. Delimitador e codificação são opcionais; por padrão o delimitador é autodetectado.</summary>
public sealed class LeitorCsv : ILeitorTabular
{
    /// <summary>Delimitador fixo; nulo aciona autodetecção.</summary>
    private readonly char? _delimitador;

    /// <summary>Codificação usada na leitura do arquivo.</summary>
    private readonly Encoding _encoding;

    /// <summary>Cria o leitor de CSV.</summary>
    /// <param name="delimitador">Delimitador de campo. Se nulo, é autodetectado (';', ',' ou tabulação) pela primeira linha.</param>
    /// <param name="encoding">Codificação do arquivo. Se nulo, usa UTF-8 com detecção de BOM.</param>
    public LeitorCsv(char? delimitador = null, Encoding? encoding = null)
    {
        _delimitador = delimitador;
        _encoding = encoding ?? Encoding.UTF8;
    }

    /// <inheritdoc/>
    public bool Suporta(string nomeArquivo)
        => nomeArquivo.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
        || nomeArquivo.EndsWith(".txt", StringComparison.OrdinalIgnoreCase);

    /// <inheritdoc/>
    public IEnumerable<object?[]> LerLinhas(Stream stream)
    {
        string conteudo;
        using (StreamReader leitor = new(stream, _encoding, detectEncodingFromByteOrderMarks: true, 1024, leaveOpen: true))
            conteudo = leitor.ReadToEnd();

        if (string.IsNullOrEmpty(conteudo))
            yield break;

        char delim = _delimitador ?? DetectarDelimitador(conteudo);

        StringBuilder campo = new();
        List<string?> linha = [];
        bool emAspas = false;

        for (int i = 0; i < conteudo.Length; i++)
        {
            char c = conteudo[i];

            if (emAspas)
            {
                if (c == '"')
                {
                    if (i + 1 < conteudo.Length && conteudo[i + 1] == '"')
                    {
                        campo.Append('"');
                        i++;
                    }
                    else
                    {
                        emAspas = false;
                    }
                }
                else
                {
                    campo.Append(c);
                }

                continue;
            }

            switch (c)
            {
                case '"':
                    emAspas = true;
                    break;
                case '\r':
                    break;
                case '\n':
                    linha.Add(campo.ToString());
                    yield return linha.ToArray();
                    linha = [];
                    campo.Clear();
                    break;
                default:
                    if (c == delim)
                    {
                        linha.Add(campo.ToString());
                        campo.Clear();
                    }
                    else
                    {
                        campo.Append(c);
                    }
                    break;
            }
        }

        if (campo.Length > 0 || linha.Count > 0)
        {
            linha.Add(campo.ToString());
            yield return linha.ToArray();
        }
    }

    /// <summary>Detecta o delimitador mais provável a partir da primeira linha do conteúdo.</summary>
    /// <param name="conteudo">Conteúdo completo do arquivo.</param>
    /// <returns>O delimitador detectado (';', ',' ou tabulação).</returns>
    private static char DetectarDelimitador(string conteudo)
    {
        int fim = conteudo.IndexOf('\n');
        string primeira = fim >= 0 ? conteudo[..fim] : conteudo;

        int pontoVirgula = primeira.Count(c => c == ';');
        int virgula = primeira.Count(c => c == ',');
        int tab = primeira.Count(c => c == '\t');

        if (tab > pontoVirgula && tab > virgula) return '\t';
        return pontoVirgula >= virgula ? ';' : ',';
    }
}

/// <summary>Leitor de arquivos Excel (.xlsx/.xls) via ExcelDataReader. Lê a primeira planilha, devolvendo valores nativos.</summary>
public sealed class LeitorExcel : ILeitorTabular
{
    /// <summary>Registra o provedor de code pages exigido pelo ExcelDataReader para arquivos .xls legados.</summary>
    static LeitorExcel()
        => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <inheritdoc/>
    public bool Suporta(string nomeArquivo)
        => nomeArquivo.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
        || nomeArquivo.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);

    /// <inheritdoc/>
    public IEnumerable<object?[]> LerLinhas(Stream stream)
    {
        using IExcelDataReader leitor = ExcelReaderFactory.CreateReader(stream);

        // Primeira planilha. Valores saem TIPADOS (double/DateTime/bool); o motor converte
        // sem stringificar — evitando o erro de separador decimal.
        while (leitor.Read())
        {
            object?[] linha = new object?[leitor.FieldCount];

            for (int i = 0; i < leitor.FieldCount; i++)
                linha[i] = leitor.GetValue(i);

            yield return linha;
        }
    }
}

/// <summary>
/// Leitor que "despivota" (unpivot) uma matriz larga em registros longos. É uma PEÇA opcional:
/// envolve outro <see cref="ILeitorTabular"/> e não altera o motor nem o componente — o importador
/// continua enxergando "1 linha lida → 1 objeto".
/// </summary>
/// <remarks>
/// Exemplo: um relatório rede × cidade (colunas: <c>rede, SP, RJ, Fortaleza…</c>) vira registros
/// <c>{ rede, cidade, quantidade }</c>. Cada linha de origem, com N colunas de valor, gera N linhas.
/// </remarks>
public sealed class LeitorDespivotado : ILeitorTabular
{
    /// <summary>Leitor interno que fornece a matriz larga.</summary>
    private readonly ILeitorTabular _interno;

    /// <summary>Nomes das colunas que se repetem em cada registro gerado.</summary>
    private readonly string[] _identificadoras;

    /// <summary>Nome da coluna sintética que recebe o nome da coluna despivotada.</summary>
    private readonly string _colunaChave;

    /// <summary>Nome da coluna sintética que recebe o valor da célula.</summary>
    private readonly string _colunaValor;

    /// <summary>Colunas a despivotar; nulo significa "todas menos as identificadoras".</summary>
    private readonly string[]? _colunasValor;

    /// <summary>Quando verdadeiro, célula vazia não gera registro.</summary>
    private readonly bool _ignorarVazias;

    /// <summary>Nome da coluna sintética de linha de origem; nulo desativa a rastreabilidade.</summary>
    private readonly string? _colunaLinhaOrigem;

    /// <summary>Linha física do cabeçalho largo; nulo autodetecta.</summary>
    private readonly int? _linhaInicio;

    /// <summary>Coluna onde a matriz começa (1-based).</summary>
    private readonly int _colunaInicio;

    /// <summary>Cria o leitor despivotado.</summary>
    /// <param name="interno">Leitor que entrega a matriz larga (ex.: <see cref="LeitorCsv"/> ou <see cref="LeitorExcel"/>).</param>
    /// <param name="colunasIdentificadoras">Colunas que se repetem em cada registro gerado (ex.: <c>["rede"]</c>).</param>
    /// <param name="nomeColunaChave">Nome da coluna sintética que recebe o nome da coluna despivotada (ex.: <c>"cidade"</c>).</param>
    /// <param name="nomeColunaValor">Nome da coluna sintética que recebe o valor da célula (ex.: <c>"quantidade"</c>).</param>
    /// <param name="colunasValor">Colunas a despivotar. Se nulo, tudo que não for identificador vira valor.</param>
    /// <param name="ignorarCelulasVazias">Quando verdadeiro, célula vazia não gera registro. Padrão: verdadeiro.</param>
    /// <param name="nomeColunaLinhaOrigem">Se informado, emite uma coluna sintética com a linha física da origem, para rastreabilidade.</param>
    /// <param name="linhaInicio">Linha física do cabeçalho largo (1 = primeira linha). Nulo autodetecta a primeira linha não-vazia.</param>
    /// <param name="colunaInicio">Coluna onde a matriz começa (1 = primeira coluna). Recorta as colunas à esquerda. Padrão: 1.</param>
    public LeitorDespivotado(
        ILeitorTabular interno,
        string[] colunasIdentificadoras,
        string nomeColunaChave,
        string nomeColunaValor,
        string[]? colunasValor = null,
        bool ignorarCelulasVazias = true,
        string? nomeColunaLinhaOrigem = null,
        int? linhaInicio = null,
        int colunaInicio = 1)
    {
        if (colunasIdentificadoras.Length == 0)
            throw new ArgumentException("Informe ao menos uma coluna identificadora.", nameof(colunasIdentificadoras));

        _interno = interno;
        _identificadoras = colunasIdentificadoras;
        _colunaChave = nomeColunaChave;
        _colunaValor = nomeColunaValor;
        _colunasValor = colunasValor;
        _ignorarVazias = ignorarCelulasVazias;
        _colunaLinhaOrigem = nomeColunaLinhaOrigem;
        _linhaInicio = linhaInicio;
        _colunaInicio = colunaInicio;
    }

    /// <inheritdoc/>
    public bool Suporta(string nomeArquivo) => _interno.Suporta(nomeArquivo);

    /// <inheritdoc/>
    public IEnumerable<object?[]> LerLinhas(Stream stream)
    {
        int linhaFisica = 0;
        bool emiteLinhaOrigem = _colunaLinhaOrigem is not null;

        List<(string Nome, int Indice)>? colunas = null;
        int[] idxIdentificadoras = [];
        List<(int Indice, string Nome)> idxValores = [];

        foreach (object?[] bruta in _interno.LerLinhas(stream))
        {
            linhaFisica++;

            // Antes da linha de início: ignora tudo (título, metadados…).
            if (_linhaInicio is int li && linhaFisica < li)
                continue;

            object?[] linha = Recortar(bruta, _colunaInicio);

            // Primeira linha útil = cabeçalho largo.
            if (colunas is null)
            {
                // No autodetect, pula brancos até achar o cabeçalho. Com âncora, confia na linha.
                if (_linhaInicio is null && LinhaVazia(linha))
                    continue;

                colunas = MontarColunas(linha);
                idxIdentificadoras = ResolverIdentificadoras(colunas);
                idxValores = ResolverValores(colunas, idxIdentificadoras);

                yield return MontarCabecalhoSintetico(emiteLinhaOrigem);
                continue;
            }

            if (LinhaVazia(linha))
                continue;

            foreach ((int indice, string nome) in idxValores)
            {
                object? valor = indice < linha.Length ? linha[indice] : null;

                if (_ignorarVazias && EhVazio(valor))
                    continue;

                yield return MontarRegistro(linha, idxIdentificadoras, nome, valor, linhaFisica, emiteLinhaOrigem);
            }
        }
    }

    /// <summary>Recorta as colunas à esquerda conforme a coluna de início.</summary>
    /// <param name="celulas">Linha original.</param>
    /// <param name="colunaInicio">Coluna inicial (1-based).</param>
    /// <returns>Linha recortada (ou a original se o início é 1).</returns>
    private static object?[] Recortar(object?[] celulas, int colunaInicio)
    {
        if (colunaInicio <= 1)
            return celulas;

        int corte = colunaInicio - 1;

        if (corte >= celulas.Length)
            return [];

        return celulas[corte..];
    }

    /// <summary>Monta a lista ordenada de (nome, índice) das colunas não-vazias do cabeçalho largo.</summary>
    /// <param name="linha">Linha de cabeçalho.</param>
    /// <returns>Colunas na ordem física.</returns>
    private static List<(string Nome, int Indice)> MontarColunas(object?[] linha)
    {
        List<(string Nome, int Indice)> colunas = [];

        for (int i = 0; i < linha.Length; i++)
        {
            string nome = (linha[i]?.ToString() ?? string.Empty).Trim();
            if (nome.Length > 0)
                colunas.Add((nome, i));
        }

        return colunas;
    }

    /// <summary>Resolve os índices das colunas identificadoras no cabeçalho.</summary>
    /// <param name="colunas">Colunas do cabeçalho.</param>
    /// <returns>Índices na mesma ordem das identificadoras configuradas.</returns>
    private int[] ResolverIdentificadoras(List<(string Nome, int Indice)> colunas)
    {
        int[] indices = new int[_identificadoras.Length];

        for (int i = 0; i < _identificadoras.Length; i++)
        {
            string alvo = _identificadoras[i].Trim();
            int achado = -1;

            foreach ((string nome, int indice) in colunas)
                if (string.Equals(nome, alvo, StringComparison.OrdinalIgnoreCase))
                {
                    achado = indice;
                    break;
                }

            if (achado < 0)
                throw new InvalidOperationException($"Coluna identificadora '{alvo}' não encontrada no cabeçalho.");

            indices[i] = achado;
        }

        return indices;
    }

    /// <summary>Resolve as colunas de valor a despivotar (lista explícita ou "todas menos identificadoras").</summary>
    /// <param name="colunas">Colunas do cabeçalho.</param>
    /// <param name="idxIdentificadoras">Índices das identificadoras.</param>
    /// <returns>Pares (índice, nome) das colunas de valor.</returns>
    private List<(int Indice, string Nome)> ResolverValores(
        List<(string Nome, int Indice)> colunas, int[] idxIdentificadoras)
    {
        List<(int Indice, string Nome)> valores = [];

        if (_colunasValor is not null)
        {
            foreach (string alvo in _colunasValor)
            {
                string nomeAlvo = alvo.Trim();

                foreach ((string nome, int indice) in colunas)
                    if (string.Equals(nome, nomeAlvo, StringComparison.OrdinalIgnoreCase))
                    {
                        valores.Add((indice, nome));
                        break;
                    }
            }

            return valores;
        }

        foreach ((string nome, int indice) in colunas)
            if (Array.IndexOf(idxIdentificadoras, indice) < 0)
                valores.Add((indice, nome));

        return valores;
    }

    /// <summary>Monta o cabeçalho sintético emitido ao motor: identificadoras + chave + valor (+ linha de origem).</summary>
    /// <param name="emiteLinhaOrigem">Se deve incluir a coluna de linha de origem.</param>
    /// <returns>Vetor de nomes de coluna.</returns>
    private object?[] MontarCabecalhoSintetico(bool emiteLinhaOrigem)
    {
        int total = _identificadoras.Length + 2 + (emiteLinhaOrigem ? 1 : 0);
        object?[] cabecalho = new object?[total];

        int pos = 0;
        foreach (string nome in _identificadoras)
            cabecalho[pos++] = nome;

        cabecalho[pos++] = _colunaChave;
        cabecalho[pos++] = _colunaValor;

        if (emiteLinhaOrigem)
            cabecalho[pos] = _colunaLinhaOrigem;

        return cabecalho;
    }

    /// <summary>Monta um registro longo para uma célula de valor específica.</summary>
    /// <param name="linha">Linha larga de origem.</param>
    /// <param name="idxIdentificadoras">Índices das colunas identificadoras.</param>
    /// <param name="chave">Nome da coluna despivotada (valor da coluna-chave).</param>
    /// <param name="valor">Valor da célula.</param>
    /// <param name="linhaFisica">Linha física de origem.</param>
    /// <param name="emiteLinhaOrigem">Se deve incluir a linha de origem.</param>
    /// <returns>Vetor de células do registro longo.</returns>
    private object?[] MontarRegistro(
        object?[] linha, int[] idxIdentificadoras, string chave, object? valor,
        int linhaFisica, bool emiteLinhaOrigem)
    {
        int total = _identificadoras.Length + 2 + (emiteLinhaOrigem ? 1 : 0);
        object?[] registro = new object?[total];

        int pos = 0;
        foreach (int idx in idxIdentificadoras)
            registro[pos++] = idx < linha.Length ? linha[idx] : null;

        registro[pos++] = chave;
        registro[pos++] = valor;

        if (emiteLinhaOrigem)
            registro[pos] = linhaFisica;

        return registro;
    }

    /// <summary>Indica se todas as células da linha estão vazias.</summary>
    /// <param name="celulas">Células da linha.</param>
    /// <returns>Verdadeiro se a linha está em branco.</returns>
    private static bool LinhaVazia(object?[] celulas)
        => celulas.All(EhVazio);

    /// <summary>Indica se uma célula está vazia (nula ou texto em branco).</summary>
    /// <param name="valor">Valor da célula.</param>
    /// <returns>Verdadeiro se vazia.</returns>
    private static bool EhVazio(object? valor)
        => valor is null || (valor is string s && string.IsNullOrWhiteSpace(s));
}

// #endregion

} // fim namespace BBiCore.Importacao

namespace BBiCore.Componentes
{

// #region Apoio do componente  ->  destino sugerido: Componentes/ (junto do .razor)

/// <summary>Define quando a importação é disparada após a seleção de arquivos.</summary>
public enum ModoDisparo
{
    /// <summary>A importação começa automaticamente ao selecionar os arquivos.</summary>
    Automatico,

    /// <summary>A seleção apenas guarda os arquivos; a importação é disparada por botão (interno ou via <c>ImportarAsync</c>).</summary>
    Manual
}

/// <summary>Descreve um arquivo grande depositado em pasta para processamento assíncrono.</summary>
/// <param name="NomeOriginal">Nome original do arquivo enviado.</param>
/// <param name="CaminhoFinal">Caminho final do arquivo já gravado na pasta de destino.</param>
/// <param name="Tamanho">Tamanho do arquivo em bytes.</param>
public sealed record ArquivoDepositado(string NomeOriginal, string CaminhoFinal, long Tamanho);

// #endregion

} // fim namespace BBiCore.Componentes

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
}

/// <summary>Implementação default de <see cref="ITemplateEmail"/>. Um sistema pode usá-la ou mapear sua própria entidade.</summary>
public sealed class TemplateEmailDto : ITemplateEmail
{
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
    public const string PrefixoSistema = "Sistema.";

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
    public static IReadOnlyList<CampoTemplate> ListarDoObjeto(Type tipoDados)
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

/// <summary>Implementação temporária de <see cref="IEnviadorEmail"/> até o envio via Exchange ser implementado.</summary>
/// <remarks>
/// >>> PARA A INSTÂNCIA QUE FOR IMPLEMENTAR O ENVIO: substitua por uma implementação real que fale com o
/// Exchange (materializando os anexos de caminho, aplicando a política de falha e a exclusão pós-anexo).
/// Registre-a no lugar deste stub na injeção de dependência de cada sistema.
/// </remarks>
public sealed class EnviadorEmailNaoImplementado : IEnviadorEmail
{
    /// <inheritdoc/>
    public Task<ResultadoEnvio> EnviarAsync(EmailResolvido email, CancellationToken cancelamento = default)
        => Task.FromResult(new ResultadoEnvio(
            false,
            "Envio de e-mail ainda não implementado. Registre um IEnviadorEmail real (Exchange) na injeção de dependência."));
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
