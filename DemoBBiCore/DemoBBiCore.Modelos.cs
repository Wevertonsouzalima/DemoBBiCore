// =============================================================================
//  DemoBBiCore.Modelos.cs  —  Modelos e mapas do app de demonstração
// -----------------------------------------------------------------------------
//  Reúne, em um arquivo só, os tipos de domínio e os mapas de importação usados
//  pelas páginas de teste. Consome a RCL BBiCore.
//
//  >>> NOTA PARA REORGANIZAÇÃO: na separação, cada tipo vira um arquivo em
//  Modelos/ (Anime, StatusAnime, FatoMercado) e os mapas em Modelos/ ou
//  Mapping/ (MapaAnimes, MapaMercado). Namespace atual: DemoBBiCore.Modelos.
// =============================================================================

using System.Globalization;
using BBiCore.Importacao;

namespace DemoBBiCore.Modelos;

// #region Domínio: animes  ->  destino sugerido: Modelos/

/// <summary>Situação de exibição de um anime no catálogo.</summary>
public enum StatusAnime
{
    /// <summary>Status não reconhecido no arquivo de origem.</summary>
    Desconhecido,

    /// <summary>Série finalizada.</summary>
    Concluido,

    /// <summary>Série em exibição.</summary>
    EmLancamento,

    /// <summary>Série pausada.</summary>
    Hiato
}

/// <summary>Entidade de anime produzida pela importação do catálogo.</summary>
public sealed class Anime
{
    /// <summary>Identificador do anime.</summary>
    public int Id { get; set; }

    /// <summary>Título do anime.</summary>
    public string Titulo { get; set; } = string.Empty;

    /// <summary>Gênero (texto livre).</summary>
    public string Genero { get; set; } = string.Empty;

    /// <summary>Nota média (usa ponto decimal na origem).</summary>
    public decimal NotaMedia { get; set; }

    /// <summary>Quantidade de episódios.</summary>
    public int Episodios { get; set; }

    /// <summary>Ano de lançamento.</summary>
    public int AnoLancamento { get; set; }

    /// <summary>Situação de exibição.</summary>
    public StatusAnime Status { get; set; }
}

// #endregion

// #region Domínio: mercado (despivot)  ->  destino sugerido: Modelos/

/// <summary>Registro despivotado de um relatório rede × cidade.</summary>
public sealed class FatoMercado
{
    /// <summary>Nome da rede de supermercado.</summary>
    public string Rede { get; set; } = string.Empty;

    /// <summary>Cidade (coluna despivotada).</summary>
    public string Cidade { get; set; } = string.Empty;

    /// <summary>Quantidade registrada na célula rede × cidade.</summary>
    public int Quantidade { get; set; }

    /// <summary>Linha física da rede no arquivo original (rastreabilidade do despivot).</summary>
    public int LinhaOrigem { get; set; }
}

// #endregion

// #region Mapas  ->  destino sugerido: Mapping/ (ou Modelos/)

/// <summary>Mapa de importação do catálogo de animes (cabeçalho: id, titulo, genero, nota_media, episodios, ano_lancamento, status).</summary>
public static class MapaAnimes
{
    /// <summary>Cria o mapa padrão para o CSV/Excel de animes.</summary>
    /// <returns>Mapa configurado.</returns>
    public static MapaImportacao<Anime> Padrao()
        => MapaImportacao<Anime>.ComCabecalho()
            .Coluna(a => a.Id, "id", obrigatoria: true)
            .Coluna(a => a.Titulo, "titulo", obrigatoria: true)
            .Coluna(a => a.Genero, "genero")
            // nota_media usa PONTO decimal → cultura invariante nesta coluna.
            .Coluna(a => a.NotaMedia, "nota_media", cultura: CultureInfo.InvariantCulture)
            .Coluna(a => a.Episodios, "episodios")
            .Coluna(a => a.AnoLancamento, "ano_lancamento")
            // status vira enum via callback; valor desconhecido cai no padrão (não é falha).
            .Coluna(a => a.Status, "status", valorPadrao: StatusAnime.Desconhecido,
                    callback: (bruto, _) => MapearStatus(bruto));

    /// <summary>Converte o texto de status do arquivo no enum correspondente.</summary>
    /// <param name="bruto">Texto de status lido do arquivo.</param>
    /// <returns>Status correspondente, ou <see cref="StatusAnime.Desconhecido"/>.</returns>
    private static StatusAnime MapearStatus(string bruto)
    {
        switch (bruto.Trim().ToLowerInvariant())
        {
            case "concluído":
            case "concluido":
                return StatusAnime.Concluido;
            case "em lançamento":
            case "em lancamento":
                return StatusAnime.EmLancamento;
            case "hiato":
                return StatusAnime.Hiato;
            default:
                return StatusAnime.Desconhecido;
        }
    }
}

/// <summary>Mapa para o relatório rede × cidade já despivotado em { rede, cidade, quantidade }.</summary>
public static class MapaMercado
{
    /// <summary>Cria o mapa longo. Combine com um <see cref="LeitorDespivotado"/> no consumo.</summary>
    /// <returns>Mapa configurado.</returns>
    public static MapaImportacao<FatoMercado> Padrao()
        => MapaImportacao<FatoMercado>.ComCabecalho()
            .Coluna(x => x.Rede, "rede", obrigatoria: true)
            .Coluna(x => x.Cidade, "cidade", obrigatoria: true)
            .Coluna(x => x.Quantidade, "quantidade")
            .Coluna(x => x.LinhaOrigem, "linha_origem");
}

// #endregion

// #region Domínio: e-mail  ->  destino sugerido: Modelos/

/// <summary>Objeto de dados de exemplo para os templates de e-mail (flat, um nível).</summary>
public sealed class PedidoCliente
{
    /// <summary>Nome do cliente.</summary>
    public string NomeCliente { get; set; } = string.Empty;

    /// <summary>E-mail do cliente (destinatário).</summary>
    public string EmailCliente { get; set; } = string.Empty;

    /// <summary>Número do pedido.</summary>
    public int NumeroPedido { get; set; }

    /// <summary>Valor total do pedido.</summary>
    public decimal Total { get; set; }

    /// <summary>Data do pedido.</summary>
    public DateTime DataPedido { get; set; }

    /// <summary>Situação do pedido.</summary>
    public string Status { get; set; } = string.Empty;
}

// #endregion
