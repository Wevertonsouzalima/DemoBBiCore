// =============================================================================
//  Exportacao.cs  —  Módulo de exportação de dados (BBiCore.Exportacao)
// -----------------------------------------------------------------------------
//  Exporta uma lista de objetos T para CSV (RFC 4180) ou Excel (.xlsx). As
//  colunas são descobertas por reflexão (menos as excluídas pelo dev); a
//  seleção final de quais exportar é feita na UI. Exibição sempre em padrão
//  brasileiro; o Excel guarda valores nativos (reimportáveis).
//
//  >>> NOTA PARA REORGANIZAÇÃO: cada tipo em sua #region, nomeada pelo destino.
//  Dependência: ClosedXML (escrita de .xlsx). O CSV não depende de nada.
// =============================================================================

using System.Globalization;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;

namespace BBiCore.Exportacao
{
    // #region Enums  ->  destino sugerido: Models/

    /// <summary>Formato de saída da exportação.</summary>
    public enum FormatoExportacao
    {
        /// <summary>Texto separado por delimitador (RFC 4180).</summary>
        Csv,

        /// <summary>Planilha Excel (.xlsx).</summary>
        Excel
    }

    // #endregion

    // #region Coluna  ->  destino sugerido: Models/

    /// <summary>Descreve uma coluna exportável de <typeparamref name="T"/>.</summary>
    /// <typeparam name="T">Tipo do item exportado.</typeparam>
    public sealed class ColunaExportacao<T>
    {
        /// <summary>Nome da propriedade de origem (usado como chave/seleção).</summary>
        public string Nome { get; }

        /// <summary>Rótulo exibido no cabeçalho do arquivo.</summary>
        public string Cabecalho { get; set; }

        /// <summary>Tipo (efetivo, sem <see cref="Nullable{T}"/>) do valor da coluna.</summary>
        public Type TipoValor { get; }

        /// <summary>Função que extrai o valor bruto do item.</summary>
        public Func<T, object?> ObterValor { get; }

        /// <summary>Formato .NET opcional para o valor (ex.: "N2", "dd/MM/yyyy").</summary>
        public string? Formato { get; set; }

        /// <summary>Cria a descrição da coluna.</summary>
        /// <param name="nome">Nome da propriedade.</param>
        /// <param name="cabecalho">Rótulo do cabeçalho.</param>
        /// <param name="tipoValor">Tipo efetivo do valor.</param>
        /// <param name="obterValor">Extrator do valor.</param>
        public ColunaExportacao(string nome, string cabecalho, Type tipoValor, Func<T, object?> obterValor)
        {
            Nome = nome;
            Cabecalho = cabecalho;
            TipoValor = tipoValor;
            ObterValor = obterValor;
        }
    }

    // #endregion

    // #region Opções  ->  destino sugerido: Models/

    /// <summary>Opções de exportação para CSV.</summary>
    public sealed class OpcoesCsv
    {
        /// <summary>Separador de campo (o usuário pode escolher; padrão ';').</summary>
        public char Separador { get; set; } = ';';

        /// <summary>Cultura de formatação dos valores (padrão pt-BR).</summary>
        public CultureInfo Cultura { get; set; } = CultureInfo.GetCultureInfo("pt-BR");

        /// <summary>Se inclui a linha de cabeçalho.</summary>
        public bool IncluirCabecalho { get; set; } = true;

        /// <summary>Se grava BOM UTF-8 (ajuda o Excel a abrir acentos corretamente).</summary>
        public bool UsarBom { get; set; } = true;
    }

    // #endregion

    // #region Motor  ->  destino sugerido: Services/

    /// <summary>Descobre colunas por reflexão e exporta listas de objetos para CSV ou Excel.</summary>
    public static class ExportadorTabular
    {
        /// <summary>Descobre as colunas exportáveis de <typeparamref name="T"/> (propriedades simples), exceto as excluídas.</summary>
        /// <typeparam name="T">Tipo do item.</typeparam>
        /// <param name="propriedadesRemover">Nomes de propriedades a NÃO oferecer (comparação sem caixa).</param>
        /// <returns>Colunas descobertas, na ordem de declaração.</returns>
        public static IReadOnlyList<ColunaExportacao<T>> DescobrirColunas<T>(IEnumerable<string>? propriedadesRemover = null)
        {
            HashSet<string> remover = new(propriedadesRemover ?? [], StringComparer.OrdinalIgnoreCase);
            List<ColunaExportacao<T>> colunas = [];

            foreach (PropertyInfo prop in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!prop.CanRead || prop.GetIndexParameters().Length > 0)
                    continue;

                if (remover.Contains(prop.Name))
                    continue;

                Type efetivo = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

                if (!EhTipoSimples(efetivo))
                    continue;

                PropertyInfo capturada = prop;
                colunas.Add(new ColunaExportacao<T>(
                    prop.Name,
                    prop.Name,
                    efetivo,
                    item => capturada.GetValue(item)));
            }

            return colunas;
        }

        /// <summary>Exporta os itens para CSV (RFC 4180), aplicando aspas quando o valor exige.</summary>
        /// <typeparam name="T">Tipo do item.</typeparam>
        /// <param name="itens">Itens a exportar.</param>
        /// <param name="colunas">Colunas selecionadas (na ordem desejada).</param>
        /// <param name="opcoes">Opções de CSV.</param>
        /// <returns>Bytes do arquivo CSV.</returns>
        public static byte[] ParaCsv<T>(IEnumerable<T> itens, IReadOnlyList<ColunaExportacao<T>> colunas, OpcoesCsv opcoes)
        {
            ArgumentNullException.ThrowIfNull(itens);
            ArgumentNullException.ThrowIfNull(colunas);
            ArgumentNullException.ThrowIfNull(opcoes);

            StringBuilder sb = new();

            if (opcoes.IncluirCabecalho)
                sb.Append(string.Join(opcoes.Separador, colunas.Select(c => EscaparCsv(c.Cabecalho, opcoes.Separador)))).Append("\r\n");

            foreach (T item in itens)
            {
                IEnumerable<string> campos = colunas.Select(c =>
                    EscaparCsv(FormatarTexto(c.ObterValor(item), c.Formato, opcoes.Cultura), opcoes.Separador));

                sb.Append(string.Join(opcoes.Separador, campos)).Append("\r\n");
            }

            byte[] conteudo = new UTF8Encoding(false).GetBytes(sb.ToString());

            if (!opcoes.UsarBom)
                return conteudo;

            byte[] bom = [0xEF, 0xBB, 0xBF];
            byte[] comBom = new byte[bom.Length + conteudo.Length];
            Buffer.BlockCopy(bom, 0, comBom, 0, bom.Length);
            Buffer.BlockCopy(conteudo, 0, comBom, bom.Length, conteudo.Length);
            return comBom;
        }

        /// <summary>Exporta os itens para Excel (.xlsx), gravando valores NATIVOS (reimportáveis) com formato BR de exibição.</summary>
        /// <typeparam name="T">Tipo do item.</typeparam>
        /// <param name="itens">Itens a exportar.</param>
        /// <param name="colunas">Colunas selecionadas (na ordem desejada).</param>
        /// <param name="nomeAba">Nome da planilha.</param>
        /// <returns>Bytes do arquivo .xlsx.</returns>
        public static byte[] ParaExcel<T>(IEnumerable<T> itens, IReadOnlyList<ColunaExportacao<T>> colunas, string nomeAba = "Dados")
        {
            ArgumentNullException.ThrowIfNull(itens);
            ArgumentNullException.ThrowIfNull(colunas);

            using XLWorkbook livro = new();
            IXLWorksheet aba = livro.Worksheets.Add(string.IsNullOrWhiteSpace(nomeAba) ? "Dados" : nomeAba);

            for (int c = 0; c < colunas.Count; c++)
            {
                IXLCell celula = aba.Cell(1, c + 1);
                celula.Value = colunas[c].Cabecalho;
                celula.Style.Font.Bold = true;
            }

            int linha = 2;

            foreach (T item in itens)
            {
                for (int c = 0; c < colunas.Count; c++)
                {
                    IXLCell celula = aba.Cell(linha, c + 1);
                    EscreverCelulaExcel(celula, colunas[c].ObterValor(item), colunas[c].TipoValor, colunas[c].Formato);
                }

                linha++;
            }

            aba.SheetView.FreezeRows(1);
            aba.Columns().AdjustToContents();

            using MemoryStream ms = new();
            livro.SaveAs(ms);
            return ms.ToArray();
        }

        /// <summary>Grava uma célula do Excel preservando o tipo nativo e aplicando um formato de exibição.</summary>
        /// <param name="celula">Célula alvo.</param>
        /// <param name="valor">Valor bruto.</param>
        /// <param name="tipoValor">Tipo efetivo da coluna.</param>
        /// <param name="formato">Formato .NET opcional (tem prioridade para texto).</param>
        private static void EscreverCelulaExcel(IXLCell celula, object? valor, Type tipoValor, string? formato)
        {
            if (valor is null)
            {
                celula.Value = string.Empty;
                return;
            }

            switch (valor)
            {
                case DateTime data:
                    celula.Value = data;
                    celula.Style.DateFormat.Format = "dd/mm/yyyy";
                    break;
                case bool logico:
                    celula.Value = logico ? "Sim" : "Não";
                    break;
                case byte or sbyte or short or ushort or int or uint or long or ulong:
                    celula.Value = Convert.ToDouble(valor, CultureInfo.InvariantCulture);
                    celula.Style.NumberFormat.Format = "#,##0";
                    break;
                case decimal or double or float:
                    celula.Value = Convert.ToDouble(valor, CultureInfo.InvariantCulture);
                    celula.Style.NumberFormat.Format = "#,##0.00";
                    break;
                default:
                    celula.Value = valor.ToString() ?? string.Empty;
                    break;
            }
        }

        /// <summary>Formata um valor como texto no padrão brasileiro (para CSV).</summary>
        /// <param name="valor">Valor bruto.</param>
        /// <param name="formato">Formato .NET opcional.</param>
        /// <param name="cultura">Cultura de formatação.</param>
        /// <returns>Texto formatado.</returns>
        private static string FormatarTexto(object? valor, string? formato, CultureInfo cultura)
        {
            if (valor is null)
                return string.Empty;

            if (!string.IsNullOrEmpty(formato))
                return string.Format(cultura, "{0:" + formato + "}", valor);

            switch (valor)
            {
                case DateTime data:
                    return data.ToString("dd/MM/yyyy", cultura);
                case bool logico:
                    return logico ? "Sim" : "Não";
                case decimal dec:
                    return dec.ToString("0.00", cultura);
                case double dbl:
                    return dbl.ToString("0.00", cultura);
                case float flt:
                    return flt.ToString("0.00", cultura);
                default:
                    return Convert.ToString(valor, cultura) ?? string.Empty;
            }
        }

        /// <summary>Aplica as regras de aspas do RFC 4180 quando o campo contém separador, aspas ou quebra de linha.</summary>
        /// <param name="campo">Texto do campo.</param>
        /// <param name="separador">Separador em uso.</param>
        /// <returns>Campo pronto para o CSV.</returns>
        private static string EscaparCsv(string campo, char separador)
        {
            bool precisa = campo.Contains(separador)
                || campo.Contains('"')
                || campo.Contains('\n')
                || campo.Contains('\r');

            if (!precisa)
                return campo;

            return "\"" + campo.Replace("\"", "\"\"") + "\"";
        }

        /// <summary>Indica se o tipo é "simples" (exportável como uma coluna).</summary>
        /// <param name="tipo">Tipo efetivo.</param>
        /// <returns>Verdadeiro para string, número, data, bool, enum, Guid.</returns>
        private static bool EhTipoSimples(Type tipo)
        {
            if (tipo.IsEnum)
                return true;

            return tipo == typeof(string)
                || tipo == typeof(bool)
                || tipo == typeof(byte) || tipo == typeof(sbyte)
                || tipo == typeof(short) || tipo == typeof(ushort)
                || tipo == typeof(int) || tipo == typeof(uint)
                || tipo == typeof(long) || tipo == typeof(ulong)
                || tipo == typeof(decimal) || tipo == typeof(double) || tipo == typeof(float)
                || tipo == typeof(DateTime) || tipo == typeof(DateTimeOffset)
                || tipo == typeof(Guid);
        }
    }

    // #endregion
}
