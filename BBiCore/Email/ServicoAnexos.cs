// =============================================================================
//  ServicoAnexos.cs  —  Regras de anexo do template (BBiCore.Email)
// -----------------------------------------------------------------------------
//  As regras aqui NÃO são opcionais: a biblioteca as garante em toda gravação,
//  sem parâmetro para desligar e sem depender de o dev de cada aplicação
//  lembrar de validar.
//
//   · Papel de corpo (cabeçalho, rodapé, inline) exige IMAGEM. Subir um PDF como
//     cabeçalho é recusado.
//   · Cabeçalho e rodapé: no máximo UM por template. Vincular um novo SUBSTITUI
//     o anterior (é o que a tela sugere: um campo, não uma lista).
//   · ContentId é GERADO pela biblioteca para todo recurso de corpo — nunca vem
//     do usuário, então não há colisão nem erro de digitação.
//   · Exclusividade: um template pode REIVINDICAR um arquivo do acervo. A regra
//     é verificada na CRIAÇÃO do vínculo, quando ainda dá para barrar sem afetar
//     quem já usava.
// =============================================================================

namespace BBiCore.Email
{
    /// <summary>Resultado de uma tentativa de vincular um arquivo do acervo a um template.</summary>
    /// <param name="Sucesso">Se o vínculo foi criado.</param>
    /// <param name="Mensagem">Motivo da recusa, quando houver.</param>
    /// <param name="Vinculo">Vínculo criado, em caso de sucesso.</param>
    public sealed record ResultadoVinculo(bool Sucesso, string? Mensagem, IVinculoAnexo? Vinculo = null);

    /// <summary>Aplica as regras de anexo e conversa com o repositório. É por aqui que a tela vincula arquivos.</summary>
    public sealed class ServicoAnexos
    {
        /// <summary>Acesso ao acervo e aos vínculos.</summary>
        private readonly IRepositorioAnexos _repositorio;

        /// <summary>Cria o serviço de anexos.</summary>
        /// <param name="repositorio">Repositório de anexos do sistema.</param>
        public ServicoAnexos(IRepositorioAnexos repositorio)
            => _repositorio = repositorio ?? throw new ArgumentNullException(nameof(repositorio));

        /// <summary>Lista os arquivos do acervo que este template pode usar, filtrados pelo papel pretendido.</summary>
        /// <param name="idTemplate">Template que está montando a lista.</param>
        /// <param name="papel">Papel pretendido: cabeçalho/rodapé/inline mostram só imagens.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Arquivos oferecíveis ao usuário naquele contexto.</returns>
        public async Task<IReadOnlyList<IAnexoAcervo>> ListarParaEscolhaAsync(
            int idTemplate,
            PapelAnexo papel,
            CancellationToken cancelamento = default)
        {
            IReadOnlyList<IAnexoAcervo> disponiveis = await _repositorio.ListarDisponiveisAsync(idTemplate, cancelamento);

            // A lista é filtrada pelo CONTEXTO: escolhendo um cabeçalho, o usuário só vê imagens.
            List<IAnexoAcervo> filtrados = [];

            foreach (IAnexoAcervo anexo in disponiveis)
                if (RegrasAnexo.PapelPermitido(papel, anexo.ContentType))
                    filtrados.Add(anexo);

            return filtrados;
        }

        /// <summary>Lista os arquivos já vinculados ao template, com o papel de cada um.</summary>
        /// <param name="idTemplate">Template.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Arquivos vinculados.</returns>
        public Task<IReadOnlyList<AnexoVinculado>> ListarDoTemplateAsync(int idTemplate, CancellationToken cancelamento = default)
            => _repositorio.ListarDoTemplateAsync(idTemplate, cancelamento);

        /// <summary>
        /// Vincula um arquivo do acervo a um template, aplicando TODAS as regras. Cabeçalho e rodapé
        /// substituem o anterior; o ContentId é gerado aqui.
        /// </summary>
        /// <param name="idTemplate">Template.</param>
        /// <param name="anexo">Arquivo do acervo.</param>
        /// <param name="papel">Papel pretendido.</param>
        /// <param name="exclusivo">Se este template reivindica o arquivo para si.</param>
        /// <param name="excluirAposAnexar">Se o arquivo de origem deve ser excluído após o envio (só nos modos de caminho).</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado do vínculo.</returns>
        public async Task<ResultadoVinculo> VincularAsync(
            int idTemplate,
            IAnexoAcervo anexo,
            PapelAnexo papel,
            bool exclusivo = false,
            bool excluirAposAnexar = false,
            CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(anexo);

            // REGRA 1 — papel de corpo exige imagem.
            if (!RegrasAnexo.PapelPermitido(papel, anexo.ContentType))
                return new ResultadoVinculo(
                    false,
                    $"'{anexo.NomeArquivo}' não é uma imagem — só imagens podem ser {RegrasAnexo.Rotulo(papel).ToLowerInvariant()}.");

            // REGRA 2 — o arquivo não pode estar reivindicado por outro template.
            if (await _repositorio.EstaReivindicadoPorOutroAsync(anexo.Id, idTemplate, cancelamento))
                return new ResultadoVinculo(
                    false,
                    $"'{anexo.NomeArquivo}' foi reservado com exclusividade por outro template.");

            // REGRA 3 — para reivindicar, o arquivo não pode estar em uso por mais ninguém.
            if (exclusivo && await _repositorio.EstaEmUsoPorOutroAsync(anexo.Id, idTemplate, cancelamento))
                return new ResultadoVinculo(
                    false,
                    $"'{anexo.NomeArquivo}' já é usado por outros templates e não pode ser reservado com exclusividade.");

            // REGRA 4 — cabeçalho e rodapé são únicos: o novo substitui o anterior.
            if (RegrasAnexo.EhPapelUnico(papel))
                await RemoverPapelAsync(idTemplate, papel, cancelamento);

            // REGRA 5 — excluir-após-anexar só faz sentido quando o arquivo vem de disco.
            bool excluir = excluirAposAnexar && anexo.ModoObtencao != ModoObtencaoAnexo.BytesNoBanco;

            IReadOnlyList<AnexoVinculado> atuais = await _repositorio.ListarDoTemplateAsync(idTemplate, cancelamento);

            VinculoAnexoDto vinculo = new()
            {
                IdTemplate = idTemplate,
                IdAnexo = anexo.Id,
                Papel = papel,
                Exclusivo = exclusivo,
                ContentId = RegrasAnexo.GerarContentId(papel),   // gerado pela lib, nunca digitado
                Ordem = atuais.Count,
                ExcluirAposAnexar = excluir
            };

            await _repositorio.SalvarVinculoAsync(vinculo, cancelamento);
            return new ResultadoVinculo(true, null, vinculo);
        }

        /// <summary>Remove o vínculo de um arquivo com o template (o arquivo continua no acervo).</summary>
        /// <param name="idVinculo">Vínculo a remover.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        public Task DesvincularAsync(int idVinculo, CancellationToken cancelamento = default)
            => _repositorio.RemoverVinculoAsync(idVinculo, cancelamento);

        /// <summary>Cadastra um arquivo novo no acervo e já o vincula ao template.</summary>
        /// <param name="idTemplate">Template.</param>
        /// <param name="anexo">Arquivo a cadastrar.</param>
        /// <param name="papel">Papel pretendido.</param>
        /// <param name="exclusivo">Se o template reivindica o arquivo.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        /// <returns>Resultado do vínculo.</returns>
        public async Task<ResultadoVinculo> CadastrarEVincularAsync(
            int idTemplate,
            IAnexoAcervo anexo,
            PapelAnexo papel,
            bool exclusivo = false,
            CancellationToken cancelamento = default)
        {
            ArgumentNullException.ThrowIfNull(anexo);

            if (!RegrasAnexo.PapelPermitido(papel, anexo.ContentType))
                return new ResultadoVinculo(
                    false,
                    $"'{anexo.NomeArquivo}' não é uma imagem — só imagens podem ser {RegrasAnexo.Rotulo(papel).ToLowerInvariant()}.");

            anexo.Id = await _repositorio.SalvarAnexoAsync(anexo, cancelamento);
            return await VincularAsync(idTemplate, anexo, papel, exclusivo, false, cancelamento);
        }

        /// <summary>Remove o vínculo do papel único (cabeçalho ou rodapé) que porventura já exista no template.</summary>
        /// <param name="idTemplate">Template.</param>
        /// <param name="papel">Papel único a liberar.</param>
        /// <param name="cancelamento">Token de cancelamento.</param>
        private async Task RemoverPapelAsync(int idTemplate, PapelAnexo papel, CancellationToken cancelamento)
        {
            IReadOnlyList<AnexoVinculado> atuais = await _repositorio.ListarDoTemplateAsync(idTemplate, cancelamento);

            foreach (AnexoVinculado atual in atuais)
                if (atual.Vinculo.Papel == papel)
                    await _repositorio.RemoverVinculoAsync(atual.Vinculo.Id, cancelamento);
        }
    }
}
