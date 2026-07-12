// =============================================================================
//  Anexos.cs  —  Acervo de anexos e vínculo com templates (BBiCore.Email)
// -----------------------------------------------------------------------------
//  TRÊS EIXOS INDEPENDENTES, que antes estavam misturados num "tipo" só:
//
//   1. O QUE O ARQUIVO É        -> ContentType ("image/png", "application/pdf")
//   2. COMO O CONTEÚDO É OBTIDO -> ModoObtencao (bytes no banco, caminho fixo,
//                                  caminho dinâmico). Vale para QUALQUER arquivo.
//   3. QUAL O PAPEL NO E-MAIL   -> Papel (cabeçalho, rodapé, inline, anexo).
//                                  Existe só para IMAGEM: apenas ela tem quatro
//                                  destinos. PDF, CSV e afins só sabem ser anexo.
//
//  DUAS COISAS DIFERENTES que só coincidem no mecanismo MIME:
//    · RECURSO DO CORPO (cabeçalho, rodapé, inline) — vai embutido por cid:,
//      o destinatário NÃO o vê no clipe do Outlook. É parte da renderização.
//    · ANEXO — o destinatário baixa. Aparece no clipe.
//
//  DUAS TABELAS:
//    · Anexo         — o ACERVO: o arquivo em si, reutilizável entre templates.
//    · TemplateAnexo — o VÍNCULO: como um template usa um arquivo do acervo
//                      (com que papel, em que ordem, exclusivo ou não).
//
//  O papel mora no VÍNCULO, não no acervo: a mesma imagem pode ser cabeçalho num
//  template e inline em outro.
//
//  >>> NOTA PARA REORGANIZAÇÃO: cada tipo em sua #region, nomeada pelo destino.
// =============================================================================

namespace BBiCore.Email
{
    // #region Enums  ->  destino sugerido: Models/

    /// <summary>Como o conteúdo do arquivo é obtido. Vale para qualquer arquivo, em qualquer papel.</summary>
    public enum ModoObtencaoAnexo
    {
        /// <summary>Bytes gravados no acervo, junto com o cadastro do arquivo.</summary>
        BytesNoBanco,

        /// <summary>Caminho fixo no servidor: sempre o mesmo arquivo (ex.: <c>templates/logo.png</c>), lido no envio.</summary>
        CaminhoFixo,

        /// <summary>Caminho com marcadores (ex.: <c>relatorios/pedido{{NumeroPedido}}.csv</c>), resolvido a cada envio.</summary>
        CaminhoDinamico
    }

    /// <summary>
    /// Papel do arquivo dentro do e-mail. Os três primeiros são RECURSOS DO CORPO (vão embutidos por
    /// <c>cid:</c> e não aparecem no clipe) e só valem para IMAGEM — apenas ela tem quatro destinos.
    /// Qualquer outro arquivo (PDF, CSV, Excel) só pode ser <see cref="Anexo"/>.
    /// </summary>
    public enum PapelAnexo
    {
        /// <summary>Imagem do topo do e-mail (modo normal). No máximo uma por template.</summary>
        Cabecalho,

        /// <summary>Imagem do rodapé do e-mail (modo normal). No máximo uma por template.</summary>
        Rodape,

        /// <summary>Imagem posicionada pelo usuário dentro do corpo (modo avançado). Quantas quiser.</summary>
        Inline,

        /// <summary>Arquivo que o destinatário baixa. Único papel possível para quem não é imagem.</summary>
        Anexo
    }

    // #endregion

    // #region Acervo  ->  destino sugerido: Models/

    /// <summary>
    /// Um arquivo do ACERVO: cadastrado uma vez e reutilizável por vários templates (a logo
    /// institucional, o PDF de termos). Não sabe nada sobre e-mail — só sobre o arquivo.
    /// </summary>
    public interface IAnexoAcervo
    {
        /// <summary>Identificador do arquivo no acervo.</summary>
        int Id { get; set; }

        /// <summary>Nome do arquivo apresentado ao usuário e ao destinatário.</summary>
        string NomeArquivo { get; set; }

        /// <summary>Content-type (ex.: <c>image/png</c>). É ele que define se o arquivo pode ser recurso de corpo.</summary>
        string? ContentType { get; set; }

        /// <summary>Como o conteúdo é obtido.</summary>
        ModoObtencaoAnexo ModoObtencao { get; set; }

        /// <summary>Bytes do arquivo. Preenchido apenas quando o modo é <see cref="ModoObtencaoAnexo.BytesNoBanco"/>.</summary>
        byte[]? Conteudo { get; set; }

        /// <summary>Caminho do arquivo (fixo ou com marcadores). Preenchido nos modos de caminho.</summary>
        string? Caminho { get; set; }

        /// <summary>Descrição livre, para o usuário reconhecer o arquivo na lista de seleção.</summary>
        string? Descricao { get; set; }
    }

    /// <summary>Implementação pronta de <see cref="IAnexoAcervo"/>, útil quando o sistema não tem entidade própria.</summary>
    public sealed class AnexoAcervoDto : IAnexoAcervo
    {
        /// <inheritdoc/>
        public int Id { get; set; }

        /// <inheritdoc/>
        public string NomeArquivo { get; set; } = string.Empty;

        /// <inheritdoc/>
        public string? ContentType { get; set; }

        /// <inheritdoc/>
        public ModoObtencaoAnexo ModoObtencao { get; set; } = ModoObtencaoAnexo.BytesNoBanco;

        /// <inheritdoc/>
        public byte[]? Conteudo { get; set; }

        /// <inheritdoc/>
        public string? Caminho { get; set; }

        /// <inheritdoc/>
        public string? Descricao { get; set; }
    }

    // #endregion

    // #region Vínculo  ->  destino sugerido: Models/

    /// <summary>
    /// Liga um template a um arquivo do acervo e diz COMO aquele template usa aquele arquivo.
    /// A mesma imagem pode ser cabeçalho num template e inline em outro — por isso o papel mora aqui.
    /// </summary>
    public interface IVinculoAnexo
    {
        /// <summary>Identificador do vínculo.</summary>
        int Id { get; set; }

        /// <summary>Template ao qual o arquivo está vinculado.</summary>
        int IdTemplate { get; set; }

        /// <summary>Arquivo do acervo.</summary>
        int IdAnexo { get; set; }

        /// <summary>Papel do arquivo NESTE template.</summary>
        PapelAnexo Papel { get; set; }

        /// <summary>
        /// Quando verdadeiro, este template REIVINDICA o arquivo: ele deixa de ser oferecido aos demais.
        /// A marca fica no vínculo (e não no acervo) porque a regra é verificada na CRIAÇÃO do vínculo —
        /// que é o momento em que dá para barrar sem afetar templates já existentes.
        /// </summary>
        bool Exclusivo { get; set; }

        /// <summary>
        /// Identificador usado no HTML como <c>cid:{ContentId}</c>. Existe SE E SOMENTE SE o papel for de
        /// corpo (cabeçalho, rodapé ou inline). É GERADO pela biblioteca — nunca digitado pelo usuário.
        /// </summary>
        string? ContentId { get; set; }

        /// <summary>Ordem de exibição na tela e de anexação no e-mail.</summary>
        int Ordem { get; set; }

        /// <summary>Se o arquivo de origem deve ser excluído após o envio. Só faz sentido nos modos de caminho.</summary>
        bool ExcluirAposAnexar { get; set; }
    }

    /// <summary>Implementação pronta de <see cref="IVinculoAnexo"/>.</summary>
    public sealed class VinculoAnexoDto : IVinculoAnexo
    {
        /// <inheritdoc/>
        public int Id { get; set; }

        /// <inheritdoc/>
        public int IdTemplate { get; set; }

        /// <inheritdoc/>
        public int IdAnexo { get; set; }

        /// <inheritdoc/>
        public PapelAnexo Papel { get; set; } = PapelAnexo.Anexo;

        /// <inheritdoc/>
        public bool Exclusivo { get; set; }

        /// <inheritdoc/>
        public string? ContentId { get; set; }

        /// <inheritdoc/>
        public int Ordem { get; set; }

        /// <inheritdoc/>
        public bool ExcluirAposAnexar { get; set; }
    }

    /// <summary>Um arquivo do acervo junto com o papel que ele cumpre num template. É o que a tela e o envio consomem.</summary>
    /// <param name="Anexo">Arquivo do acervo.</param>
    /// <param name="Vinculo">Como este template usa o arquivo.</param>
    public sealed record AnexoVinculado(IAnexoAcervo Anexo, IVinculoAnexo Vinculo)
    {
        /// <summary>Indica que o arquivo vai EMBUTIDO no corpo (cabeçalho, rodapé ou inline), e não no clipe.</summary>
        public bool EhRecursoDeCorpo => RegrasAnexo.EhRecursoDeCorpo(Vinculo.Papel);
    }

    // #endregion

    // #region Regras  ->  destino sugerido: Services/

    /// <summary>Regras que a biblioteca garante sobre anexos — não são opcionais nem configuráveis.</summary>
    public static class RegrasAnexo
    {
        /// <summary>Prefixo dos ContentId gerados, para não colidir com nada que o usuário escreva.</summary>
        private const string PrefixoContentId = "bbi";

        /// <summary>Indica se o papel é um recurso do corpo (vai embutido por cid:, fora do clipe).</summary>
        /// <param name="papel">Papel do arquivo.</param>
        /// <returns>Verdadeiro para cabeçalho, rodapé e inline.</returns>
        public static bool EhRecursoDeCorpo(PapelAnexo papel)
        {
            switch (papel)
            {
                case PapelAnexo.Cabecalho:
                case PapelAnexo.Rodape:
                case PapelAnexo.Inline:
                    return true;
                case PapelAnexo.Anexo:
                default:
                    return false;
            }
        }

        /// <summary>Indica se o content-type é de imagem (só imagem pode ser recurso de corpo).</summary>
        /// <param name="contentType">Content-type do arquivo.</param>
        /// <returns>Verdadeiro quando é imagem.</returns>
        public static bool EhImagem(string? contentType)
            => !string.IsNullOrWhiteSpace(contentType)
               && contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);

        /// <summary>Indica se um papel é válido para um arquivo — cabeçalho, rodapé e inline exigem imagem.</summary>
        /// <param name="papel">Papel pretendido.</param>
        /// <param name="contentType">Content-type do arquivo.</param>
        /// <returns>Verdadeiro quando a combinação é permitida.</returns>
        public static bool PapelPermitido(PapelAnexo papel, string? contentType)
        {
            if (!EhRecursoDeCorpo(papel))
                return true;

            return EhImagem(contentType);
        }

        /// <summary>Gera o ContentId de um recurso de corpo. Único e sem colisão — nunca vem do usuário.</summary>
        /// <param name="papel">Papel do recurso.</param>
        /// <returns>ContentId pronto para o <c>cid:</c>; nulo quando o papel não é de corpo.</returns>
        public static string? GerarContentId(PapelAnexo papel)
        {
            if (!EhRecursoDeCorpo(papel))
                return null;

            switch (papel)
            {
                case PapelAnexo.Cabecalho:
                    return $"{PrefixoContentId}-cabecalho";
                case PapelAnexo.Rodape:
                    return $"{PrefixoContentId}-rodape";
                case PapelAnexo.Inline:
                default:
                    // Imagens do corpo podem ser várias: cada uma ganha um identificador próprio.
                    return $"{PrefixoContentId}-img-{Guid.NewGuid():N}"[..24];
            }
        }

        /// <summary>Indica se o papel admite apenas UM vínculo por template (cabeçalho e rodapé).</summary>
        /// <param name="papel">Papel do arquivo.</param>
        /// <returns>Verdadeiro para cabeçalho e rodapé.</returns>
        public static bool EhPapelUnico(PapelAnexo papel)
        {
            switch (papel)
            {
                case PapelAnexo.Cabecalho:
                case PapelAnexo.Rodape:
                    return true;
                case PapelAnexo.Inline:
                case PapelAnexo.Anexo:
                default:
                    return false;
            }
        }

        /// <summary>Rótulo do papel em português, para a tela.</summary>
        /// <param name="papel">Papel do arquivo.</param>
        /// <returns>Texto exibível.</returns>
        public static string Rotulo(PapelAnexo papel)
        {
            switch (papel)
            {
                case PapelAnexo.Cabecalho:
                    return "Cabeçalho";
                case PapelAnexo.Rodape:
                    return "Rodapé";
                case PapelAnexo.Inline:
                    return "Imagem no corpo";
                case PapelAnexo.Anexo:
                default:
                    return "Anexo";
            }
        }
    }

    // #endregion

    // #region Repositório  ->  destino sugerido: Services/

    /// <summary>
    /// Acesso ao acervo de anexos e aos vínculos. Cada sistema implementa contra as suas tabelas
    /// (o mínimo de colunas está em <see cref="IAnexoAcervo"/> e <see cref="IVinculoAnexo"/>).
    /// </summary>
    public interface IRepositorioAnexos
    {
        /// <summary>
        /// Lista os arquivos do acervo DISPONÍVEIS para um template: todos, menos os reivindicados
        /// com exclusividade por OUTROS templates.
        /// </summary>
        /// <param name="idTemplate">Template que está montando a lista.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Arquivos que este template pode usar.</returns>
        Task<IReadOnlyList<IAnexoAcervo>> ListarDisponiveisAsync(int idTemplate, CancellationToken cancelamento = default);

        /// <summary>Lista os arquivos já vinculados a um template, com o papel de cada um.</summary>
        /// <param name="idTemplate">Template.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Arquivos vinculados.</returns>
        Task<IReadOnlyList<AnexoVinculado>> ListarDoTemplateAsync(int idTemplate, CancellationToken cancelamento = default);

        /// <summary>Indica se um arquivo do acervo já foi reivindicado com exclusividade por outro template.</summary>
        /// <param name="idAnexo">Arquivo do acervo.</param>
        /// <param name="idTemplateAtual">Template que está tentando usá-lo (é ignorado na checagem).</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Verdadeiro quando outro template o reivindicou.</returns>
        Task<bool> EstaReivindicadoPorOutroAsync(int idAnexo, int idTemplateAtual, CancellationToken cancelamento = default);

        /// <summary>Indica se um arquivo do acervo está vinculado a algum template além do informado.</summary>
        /// <param name="idAnexo">Arquivo do acervo.</param>
        /// <param name="idTemplateAtual">Template que está tentando reivindicá-lo.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Verdadeiro quando há outro template usando o arquivo.</returns>
        Task<bool> EstaEmUsoPorOutroAsync(int idAnexo, int idTemplateAtual, CancellationToken cancelamento = default);

        /// <summary>Grava um arquivo novo no acervo.</summary>
        /// <param name="anexo">Arquivo a cadastrar.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Identificador gerado.</returns>
        Task<int> SalvarAnexoAsync(IAnexoAcervo anexo, CancellationToken cancelamento = default);

        /// <summary>Grava (ou atualiza) o vínculo entre um template e um arquivo.</summary>
        /// <param name="vinculo">Vínculo a gravar.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        Task SalvarVinculoAsync(IVinculoAnexo vinculo, CancellationToken cancelamento = default);

        /// <summary>Remove um vínculo (o arquivo continua no acervo).</summary>
        /// <param name="idVinculo">Vínculo a remover.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        Task RemoverVinculoAsync(int idVinculo, CancellationToken cancelamento = default);
    }

    // #endregion
}
