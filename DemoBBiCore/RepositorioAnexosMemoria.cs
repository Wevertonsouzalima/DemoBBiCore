// Implementação em memória do acervo de anexos — só para o projeto de demonstração
// funcionar sem banco. Em produção, cada sistema implementa contra as suas tabelas
// (ver BBiCore/Email/BBiCore-Email.sql).

using BBiCore.Email;

namespace DemoBBiCore;

/// <summary>Acervo de anexos e vínculos guardados em memória (apenas para a demonstração).</summary>
public sealed class RepositorioAnexosMemoria : IRepositorioAnexos
{
    /// <summary>Arquivos do acervo, por id.</summary>
    private readonly Dictionary<int, IAnexoAcervo> _acervo = [];

    /// <summary>Vínculos entre templates e arquivos.</summary>
    private readonly List<IVinculoAnexo> _vinculos = [];

    /// <summary>Próximo id de anexo.</summary>
    private int _proximoAnexo = 1;

    /// <summary>Próximo id de vínculo.</summary>
    private int _proximoVinculo = 1;

    /// <inheritdoc/>
    public Task<IReadOnlyList<IAnexoAcervo>> ListarDisponiveisAsync(int idTemplate, CancellationToken cancelamento = default)
    {
        // Todos do acervo, menos os reivindicados com exclusividade por OUTROS templates.
        HashSet<int> reivindicados = [.. _vinculos
            .Where(v => v.Exclusivo && v.IdTemplate != idTemplate)
            .Select(v => v.IdAnexo)];

        IReadOnlyList<IAnexoAcervo> disponiveis = [.. _acervo.Values.Where(a => !reivindicados.Contains(a.Id))];
        return Task.FromResult(disponiveis);
    }

    /// <inheritdoc/>
    public Task<IReadOnlyList<AnexoVinculado>> ListarDoTemplateAsync(int idTemplate, CancellationToken cancelamento = default)
    {
        IReadOnlyList<AnexoVinculado> itens = [.. _vinculos
            .Where(v => v.IdTemplate == idTemplate)
            .Where(v => _acervo.ContainsKey(v.IdAnexo))
            .Select(v => new AnexoVinculado(_acervo[v.IdAnexo], v))];

        return Task.FromResult(itens);
    }

    /// <inheritdoc/>
    public Task<bool> EstaReivindicadoPorOutroAsync(int idAnexo, int idTemplateAtual, CancellationToken cancelamento = default)
        => Task.FromResult(_vinculos.Any(v => v.IdAnexo == idAnexo && v.Exclusivo && v.IdTemplate != idTemplateAtual));

    /// <inheritdoc/>
    public Task<bool> EstaEmUsoPorOutroAsync(int idAnexo, int idTemplateAtual, CancellationToken cancelamento = default)
        => Task.FromResult(_vinculos.Any(v => v.IdAnexo == idAnexo && v.IdTemplate != idTemplateAtual));

    /// <inheritdoc/>
    public Task<int> SalvarAnexoAsync(IAnexoAcervo anexo, CancellationToken cancelamento = default)
    {
        if (anexo.Id == 0)
            anexo.Id = _proximoAnexo++;

        _acervo[anexo.Id] = anexo;
        return Task.FromResult(anexo.Id);
    }

    /// <inheritdoc/>
    public Task SalvarVinculoAsync(IVinculoAnexo vinculo, CancellationToken cancelamento = default)
    {
        if (vinculo.Id == 0)
        {
            vinculo.Id = _proximoVinculo++;
            _vinculos.Add(vinculo);
        }

        return Task.CompletedTask;
    }

    /// <inheritdoc/>
    public Task RemoverVinculoAsync(int idVinculo, CancellationToken cancelamento = default)
    {
        _vinculos.RemoveAll(v => v.Id == idVinculo);
        return Task.CompletedTask;
    }
}
