# BBiCore — biblioteca de componentes (RCL, .NET 9)


## Organização (uma pasta por módulo, autocontida)

Cada módulo fica numa pasta com TUDO junto — componentes, CSS, JS colocado e
classes — para copiar módulo a módulo sem caçar arquivos soltos:

```
Importacao/   BBiImportadorArquivo.razor (+.css)   Importacao.cs   (motor + despivot)
Email/        BBiEmail / BBiComporEmail (+.css)    BBiComporEmail.razor.js   Email.cs
Exportacao/   BBiExportador (+.css)                BBiExportador.razor.js    Exportacao.cs
Mascaras/     BBiCampoMascarado/Moeda/Data (+.css) BBiCampoMascarado.razor.js Mascaras.cs
_Imports.razor   BBiCore.csproj
```

O JS é *colocado* (`{Componente}.razor.js`, ao lado do `.razor`) e servido em
`_content/BBiCore/{pasta}/{Componente}.razor.js`. Isso mantém o JS na pasta do
módulo (bom para a cópia) e continua funcionando por referência de projeto ou NuGet.

Biblioteca única com três áreas, consolidada no menor número de arquivos possível
para facilitar a cópia manual e a divisão posterior no padrão do time.

## Conteúdo

Todo o **código C#** está em `BBiCore.cs` (um arquivo), dividido em três namespaces:

- **`BBiCore.Importacao`** — importador tabular (CSV/Excel) e despivotador:
  motor `ImportadorTabular`, `MapaImportacao<T>`, `OpcoesImportacao` (âncoras de
  início, tratamento de brancos), leitores `LeitorCsv`/`LeitorExcel`/`LeitorDespivotado`,
  tipos de resultado e contratos.
- **`BBiCore.Componentes`** — tipos de apoio dos componentes (`ModoDisparo`,
  `ArquivoDepositado`).
- **`BBiCore.Email`** — motor de template multi-fonte (`MotorTemplateEmail`),
  contrato (`ITemplateEmail`/`IAnexoTemplate` + enums + DTOs), e envio
  (`IEnviadorEmail`, `EmailResolvido`, `ResolvedorEmail`, stub `EnviadorEmailNaoImplementado`).

- **`BBiCore.Exportacao`** — exporta listas para CSV (RFC 4180) e Excel (.xlsx),
  com descoberta de colunas por reflexão e seleção na UI (`BBiExportador`).
- **`BBiCore.Mascaras`** — máscaras/validações brasileiras (CPF, CNPJ, telefone, CEP,
  moeda, data) como utilitários puros + campos (`BBiCampoMascarado`, `BBiCampoMoeda`,
  `BBiCampoData`): exibem em BR, entregam valor nativo/cru (sem conflito ao salvar).

Os **componentes visuais** ficam nos `.razor` (Razor não se funde com `.cs`):

- `Componentes/BBiImportadorArquivo.razor` — importação/despivot de arquivos.
- `Componentes/BBiComporEmail.razor` — composição do corpo do e-mail (modo normal/avançado).
- `Componentes/BBiEmail.razor` — componente completo de e-mail (envelope + corpo + anexos + envio).


## Configuração no sistema consumidor

O envio de e-mail é centralizado. Cada sistema tem o SEU grupo de configurações
(`OpcoesEmail`: remetente, usuário sistêmico, senha, servidor) e registra UMA
implementação de `IEnviadorEmail`:

```csharp
// Configurações do sistema (idealmente vindas do appsettings; senha de um cofre).
OpcoesEmail opcoes = builder.Configuration.GetSection("Email").Get<OpcoesEmail>()!;
builder.Services.AddSingleton(opcoes);

// Produção:
builder.Services.AddScoped<IEnviadorEmail, EnviadorEmailExchange>();
// Testes (grava .eml em disco, sem servidor):
// builder.Services.AddScoped<IEnviadorEmail, EnviadorEmailSimulado>();
```

O envio monta o MIME (corpo, imagens inline por cid:, anexos das formas 1 e 3),
aplica a política de falha de anexo, sanitiza caminhos (path traversal) e trata a
exclusão pós-anexo. Usa `SmtpClient` nativo; para o caminho mais moderno, troque a
entrega por MailKit/Graph implementando a mesma `IEnviadorEmail`.

A persistência de templates é do consumidor: implemente seu repositório contra a
tabela `Email.Templates` (local por sistema); o componente apenas dispara `OnSalvar`.

## Divisão futura

Cada tipo em `BBiCore.cs` está isolado em uma `#region` nomeada pelo destino
sugerido (Models/, Services/, Interfaces/, Readers/, Mapping/…). As áreas
Importacao e Email são independentes e podem virar projetos/pastas separados.
Ao dividir, preserve os namespaces e a API pública.

## Observação

Este pacote não foi compilado no ambiente de geração (sem SDK .NET / NuGet).
Balanceamentos e lógica foram conferidos; o primeiro `dotnet build`/`run` no seu
ambiente é o teste real, e os `.razor` (JS interop + iframe) são a parte mais
sujeita a ajuste fino.
