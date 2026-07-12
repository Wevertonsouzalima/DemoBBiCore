// =============================================================================
//  ExportadorPdf.cs  —  Exportação tabular para PDF (BBiCore.Exportacao)
// -----------------------------------------------------------------------------
//  ARQUIVO ISOLADO DE PROPÓSITO: é o único que depende do gerador de PDF.
//  Se um dia trocar de biblioteca, mexe-se só aqui.
//
//  PACOTE:
//      <PackageReference Include="PDFsharp-MigraDoc" Version="6.1.1" />
//  Licença MIT — uso comercial livre, sem restrição de faturamento. (O QuestPDF,
//  mais conhecido, é gratuito apenas para empresas com receita abaixo de US$ 1M;
//  o iText é AGPL. Por isso a escolha recaiu no PDFsharp/MigraDoc.)
//
//  DIFERENÇA PARA CSV/EXCEL: a página é FINITA. Por isso aqui existem decisões
//  que os outros formatos não têm — orientação, tamanho de fonte, quebra de
//  texto na célula e paginação com o cabeçalho repetido.
//
//  A formatação dos valores é a MESMA do CSV (padrão brasileiro), reaproveitada
//  de ExportadorTabular.FormatarTexto.
// =============================================================================

using System.Globalization;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;

namespace BBiCore.Exportacao
{
    // #region Opções  ->  destino sugerido: Models/

    /// <summary>Orientação da página do PDF.</summary>
    public enum OrientacaoPdf
    {
        /// <summary>Decide sozinho: paisagem quando há muitas colunas, retrato caso contrário.</summary>
        Automatica,

        /// <summary>Página em pé.</summary>
        Retrato,

        /// <summary>Página deitada (cabe mais coluna).</summary>
        Paisagem
    }

    /// <summary>Opções de exportação para PDF.</summary>
    public sealed class OpcoesPdf
    {
        /// <summary>Título impresso no topo da primeira página. Vazio omite o título.</summary>
        public string Titulo { get; set; } = string.Empty;

        /// <summary>Subtítulo opcional, abaixo do título (ex.: período do relatório).</summary>
        public string? Subtitulo { get; set; }

        /// <summary>Orientação da página. Padrão: automática.</summary>
        public OrientacaoPdf Orientacao { get; set; } = OrientacaoPdf.Automatica;

        /// <summary>A partir de quantas colunas a orientação automática vira paisagem.</summary>
        public int ColunasParaPaisagem { get; set; } = 6;

        /// <summary>Tamanho de fonte desejado para os dados (em pontos). É reduzido automaticamente se não couber.</summary>
        public double TamanhoFonte { get; set; } = 9;

        /// <summary>Piso do tamanho de fonte: abaixo disso a leitura fica ruim e o componente avisa em vez de encolher mais.</summary>
        public double TamanhoFonteMinimo { get; set; } = 6;

        /// <summary>Se imprime o rodapé com "página X de Y".</summary>
        public bool NumerarPaginas { get; set; } = true;

        /// <summary>Se imprime a data/hora de geração no rodapé.</summary>
        public bool MostrarDataGeracao { get; set; } = true;

        /// <summary>Cultura de formatação dos valores (padrão pt-BR).</summary>
        public CultureInfo Cultura { get; set; } = CultureInfo.GetCultureInfo("pt-BR");
    }

    // #endregion

    // #region Motor  ->  destino sugerido: Services/

    /// <summary>Gera o PDF paginado de uma lista de objetos, com cabeçalho repetido e rodapé numerado.</summary>
    public static class ExportadorPdf
    {
        /// <summary>Cor de fundo do cabeçalho da tabela.</summary>
        private static readonly Color CorCabecalho = new(230, 235, 245);

        /// <summary>Cor das linhas de grade.</summary>
        private static readonly Color CorGrade = new(180, 190, 205);

        /// <summary>Exporta os itens para PDF.</summary>
        /// <typeparam name="T">Tipo do item.</typeparam>
        /// <param name="itens">Itens a exportar.</param>
        /// <param name="colunas">Colunas selecionadas (na ordem desejada).</param>
        /// <param name="opcoes">Opções do PDF.</param>
        /// <returns>Bytes do arquivo .pdf.</returns>
        public static byte[] Gerar<T>(IEnumerable<T> itens, IReadOnlyList<ColunaExportacao<T>> colunas, OpcoesPdf opcoes)
        {
            ArgumentNullException.ThrowIfNull(itens);
            ArgumentNullException.ThrowIfNull(colunas);
            ArgumentNullException.ThrowIfNull(opcoes);

            if (colunas.Count == 0)
                throw new ArgumentException("Selecione ao menos uma coluna para exportar.", nameof(colunas));

            Document documento = new();
            documento.Info.Title = string.IsNullOrWhiteSpace(opcoes.Titulo) ? "Exportação" : opcoes.Titulo;

            ConfigurarEstilos(documento, CalcularTamanhoFonte(colunas.Count, opcoes));

            Section secao = documento.AddSection();
            ConfigurarPagina(secao, colunas.Count, opcoes);

            EscreverTitulo(secao, opcoes);
            EscreverRodape(secao, opcoes);

            Table tabela = MontarTabela(secao, colunas, opcoes);
            PreencherLinhas(tabela, itens, colunas, opcoes);

            return Renderizar(documento);
        }

        /// <summary>Define os estilos do documento (fonte base e do cabeçalho).</summary>
        /// <param name="documento">Documento em construção.</param>
        /// <param name="tamanhoFonte">Tamanho de fonte calculado para os dados.</param>
        private static void ConfigurarEstilos(Document documento, double tamanhoFonte)
        {
            Style normal = documento.Styles["Normal"]!;
            normal.Font.Name = "Arial";
            normal.Font.Size = tamanhoFonte;

            Style titulo = documento.Styles.AddStyle("TituloRelatorio", "Normal");
            titulo.Font.Size = tamanhoFonte + 5;
            titulo.Font.Bold = true;
            titulo.ParagraphFormat.SpaceAfter = Unit.FromPoint(2);

            Style subtitulo = documento.Styles.AddStyle("SubtituloRelatorio", "Normal");
            subtitulo.Font.Size = tamanhoFonte + 1;
            subtitulo.Font.Color = new Color(90, 100, 115);
            subtitulo.ParagraphFormat.SpaceAfter = Unit.FromPoint(8);

            Style rodape = documento.Styles["Footer"]!;
            rodape.Font.Size = tamanhoFonte - 1;
            rodape.Font.Color = new Color(110, 120, 135);
        }

        /// <summary>Configura tamanho, orientação e margens da página.</summary>
        /// <param name="secao">Seção do documento.</param>
        /// <param name="quantidadeColunas">Quantas colunas serão impressas.</param>
        /// <param name="opcoes">Opções do PDF.</param>
        private static void ConfigurarPagina(Section secao, int quantidadeColunas, OpcoesPdf opcoes)
        {
            secao.PageSetup = secao.Document!.DefaultPageSetup.Clone();
            secao.PageSetup.PageFormat = PageFormat.A4;

            secao.PageSetup.Orientation = DecidirOrientacao(quantidadeColunas, opcoes) == OrientacaoPdf.Paisagem
                ? Orientation.Landscape
                : Orientation.Portrait;

            secao.PageSetup.TopMargin = Unit.FromCentimeter(1.6);
            secao.PageSetup.BottomMargin = Unit.FromCentimeter(1.6);
            secao.PageSetup.LeftMargin = Unit.FromCentimeter(1.2);
            secao.PageSetup.RightMargin = Unit.FromCentimeter(1.2);
        }

        /// <summary>Resolve a orientação: a automática vira paisagem quando há colunas demais para o retrato.</summary>
        /// <param name="quantidadeColunas">Quantas colunas serão impressas.</param>
        /// <param name="opcoes">Opções do PDF.</param>
        /// <returns>Orientação efetiva.</returns>
        internal static OrientacaoPdf DecidirOrientacao(int quantidadeColunas, OpcoesPdf opcoes)
        {
            switch (opcoes.Orientacao)
            {
                case OrientacaoPdf.Retrato:
                    return OrientacaoPdf.Retrato;
                case OrientacaoPdf.Paisagem:
                    return OrientacaoPdf.Paisagem;
                case OrientacaoPdf.Automatica:
                default:
                    return quantidadeColunas > opcoes.ColunasParaPaisagem
                        ? OrientacaoPdf.Paisagem
                        : OrientacaoPdf.Retrato;
            }
        }

        /// <summary>
        /// Calcula o tamanho da fonte: encolhe conforme aumenta o número de colunas, respeitando o piso
        /// configurado (abaixo dele a leitura fica ruim — a saída passa a ser a quebra de texto na célula).
        /// </summary>
        /// <param name="quantidadeColunas">Quantas colunas serão impressas.</param>
        /// <param name="opcoes">Opções do PDF.</param>
        /// <returns>Tamanho de fonte, em pontos.</returns>
        internal static double CalcularTamanhoFonte(int quantidadeColunas, OpcoesPdf opcoes)
        {
            double tamanho = opcoes.TamanhoFonte;

            // A partir de 8 colunas, cada coluna extra tira 0,4 pt — até o piso.
            if (quantidadeColunas > 8)
                tamanho -= (quantidadeColunas - 8) * 0.4;

            return Math.Max(opcoes.TamanhoFonteMinimo, tamanho);
        }

        /// <summary>Escreve o título e o subtítulo no topo da primeira página.</summary>
        /// <param name="secao">Seção do documento.</param>
        /// <param name="opcoes">Opções do PDF.</param>
        private static void EscreverTitulo(Section secao, OpcoesPdf opcoes)
        {
            if (!string.IsNullOrWhiteSpace(opcoes.Titulo))
                secao.AddParagraph(opcoes.Titulo).Style = "TituloRelatorio";

            if (!string.IsNullOrWhiteSpace(opcoes.Subtitulo))
                secao.AddParagraph(opcoes.Subtitulo).Style = "SubtituloRelatorio";
        }

        /// <summary>Escreve o rodapé com a data de geração e a numeração "página X de Y".</summary>
        /// <param name="secao">Seção do documento.</param>
        /// <param name="opcoes">Opções do PDF.</param>
        private static void EscreverRodape(Section secao, OpcoesPdf opcoes)
        {
            if (!opcoes.NumerarPaginas && !opcoes.MostrarDataGeracao)
                return;

            Paragraph rodape = secao.Footers.Primary.AddParagraph();
            rodape.Format.Alignment = ParagraphAlignment.Center;

            if (opcoes.MostrarDataGeracao)
            {
                rodape.AddText($"Gerado em {DateTime.Now.ToString("dd/MM/yyyy HH:mm", opcoes.Cultura)}");

                if (opcoes.NumerarPaginas)
                    rodape.AddText("   —   ");
            }

            if (opcoes.NumerarPaginas)
            {
                rodape.AddText("Página ");
                rodape.AddPageField();
                rodape.AddText(" de ");
                rodape.AddNumPagesField();
            }
        }

        /// <summary>Cria a tabela e a linha de cabeçalho, marcada para se repetir em toda página.</summary>
        /// <typeparam name="T">Tipo do item.</typeparam>
        /// <param name="secao">Seção do documento.</param>
        /// <param name="colunas">Colunas selecionadas.</param>
        /// <param name="opcoes">Opções do PDF.</param>
        /// <returns>Tabela pronta para receber as linhas.</returns>
        private static Table MontarTabela<T>(Section secao, IReadOnlyList<ColunaExportacao<T>> colunas, OpcoesPdf opcoes)
        {
            Table tabela = secao.AddTable();
            tabela.Borders.Width = 0.5;
            tabela.Borders.Color = CorGrade;
            tabela.Rows.LeftIndent = 0;

            // Largura útil dividida igualmente; o texto quebra dentro da célula quando não cabe.
            double larguraUtil = secao.PageSetup.PageWidth.Point
                - secao.PageSetup.LeftMargin.Point
                - secao.PageSetup.RightMargin.Point;

            double larguraColuna = larguraUtil / colunas.Count;

            foreach (ColunaExportacao<T> coluna in colunas)
            {
                Column c = tabela.AddColumn(Unit.FromPoint(larguraColuna));

                c.Format.Alignment = ExportadorTabular.EhNumerico(coluna.TipoValor)
                    ? ParagraphAlignment.Right
                    : ParagraphAlignment.Left;
            }

            Row cabecalho = tabela.AddRow();
            cabecalho.HeadingFormat = true;          // repete o cabeçalho em cada página
            cabecalho.Format.Font.Bold = true;
            cabecalho.Shading.Color = CorCabecalho;
            cabecalho.VerticalAlignment = VerticalAlignment.Center;

            for (int i = 0; i < colunas.Count; i++)
            {
                Paragraph p = cabecalho.Cells[i].AddParagraph(colunas[i].Cabecalho);
                p.Format.Alignment = ParagraphAlignment.Center;
            }

            return tabela;
        }

        /// <summary>Preenche as linhas de dados, com os valores no padrão brasileiro.</summary>
        /// <typeparam name="T">Tipo do item.</typeparam>
        /// <param name="tabela">Tabela do documento.</param>
        /// <param name="itens">Itens a imprimir.</param>
        /// <param name="colunas">Colunas selecionadas.</param>
        /// <param name="opcoes">Opções do PDF.</param>
        private static void PreencherLinhas<T>(
            Table tabela,
            IEnumerable<T> itens,
            IReadOnlyList<ColunaExportacao<T>> colunas,
            OpcoesPdf opcoes)
        {
            bool alternada = false;

            foreach (T item in itens)
            {
                Row linha = tabela.AddRow();
                linha.VerticalAlignment = VerticalAlignment.Center;

                // Faixa zebrada: ajuda a acompanhar a linha em tabelas largas.
                if (alternada)
                    linha.Shading.Color = new Color(246, 248, 251);

                alternada = !alternada;

                for (int i = 0; i < colunas.Count; i++)
                {
                    string texto = ExportadorTabular.FormatarTexto(
                        colunas[i].ObterValor(item),
                        colunas[i].Formato,
                        opcoes.Cultura);

                    linha.Cells[i].AddParagraph(texto);
                }
            }
        }

        /// <summary>Renderiza o documento em bytes de PDF.</summary>
        /// <param name="documento">Documento montado.</param>
        /// <returns>Bytes do arquivo .pdf.</returns>
        private static byte[] Renderizar(Document documento)
        {
            PdfDocumentRenderer renderizador = new()
            {
                Document = documento
            };

            renderizador.RenderDocument();

            using MemoryStream ms = new();
            renderizador.PdfDocument.Save(ms, false);
            return ms.ToArray();
        }
    }

    // #endregion
}
