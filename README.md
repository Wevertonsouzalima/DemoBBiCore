# DemoBBiCore — solução de demonstração da BBiCore

Solução .NET 9 com dois projetos:

- **BBiCore** — a biblioteca (RCL) completa: importação, despivot e e-mail.
- **DemoBBiCore** — app Blazor Server que consome a BBiCore e exercita os três
  recursos, um por página.

## Rodar

```bash
dotnet restore
dotnet run --project DemoBBiCore
```
Acesse `http://localhost:5090`.

## As três demonstrações

1. **Importar** (`/`) — lê um CSV de catálogo de animes e converte em objetos
   `Anime`, com relatório de linhas ok/falha. Exemplos em
   `wwwroot/exemplos/catalogo_animes.csv` (limpo) e `catalogo_animes_sujo.csv`.

2. **Despivot** (`/cruzado`) — lê uma matriz larga rede × cidade e a transforma
   em registros `{ rede, cidade, quantidade }`. Arquivo de teste em
   `wwwroot/exemplos/mercados_cidades.csv` (e uma versão deslocada para exercitar
   as âncoras de linha/coluna).

3. **E-mail** (`/email`) — o componente `BBiEmail` com o objeto `PedidoCliente`:
   envelope (assunto/destinatários/cc/cco com marcadores), classificação, corpo
   nos modos normal/avançado, anexos e "Enviar teste". O envio usa o
   `IEnviadorEmail` registrado no `Program.cs` (stub nesta demo).

## Configuração relevante (Program.cs)

```csharp
builder.Services.AddScoped<IEnviadorEmail, EnviadorEmailNaoImplementado>();
```
Troque o stub pela implementação real (Exchange) quando existir — sem tocar nos
componentes nem nas páginas.

## Observação

Não compilado no ambiente de geração (sem SDK .NET / NuGet). O primeiro
`dotnet build`/`run` no seu ambiente é o teste real.
