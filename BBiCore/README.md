# BBiCore — biblioteca de componentes (RCL, .NET 9)

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

Os **componentes visuais** ficam nos `.razor` (Razor não se funde com `.cs`):

- `Componentes/BBiImportadorArquivo.razor` — importação/despivot de arquivos.
- `Componentes/BBiComporEmail.razor` — composição do corpo do e-mail (modo normal/avançado).
- `Componentes/BBiEmail.razor` — componente completo de e-mail (envelope + corpo + anexos + envio).

## Arquivos (mínimo para a biblioteca)

```
BBiCore.csproj
BBiCore.cs                              <- todo o código C#
_Imports.razor
Componentes/BBiImportadorArquivo.razor (+ .razor.css)
Componentes/BBiComporEmail.razor        (+ .razor.css)
Componentes/BBiEmail.razor              (+ .razor.css)
wwwroot/bbiComporEmail.js
```

## Configuração no sistema consumidor

O envio de e-mail é centralizado: registre UMA implementação de `IEnviadorEmail`
na injeção de dependência (o stub serve de placeholder até o Exchange entrar):

```csharp
builder.Services.AddScoped<IEnviadorEmail, EnviadorEmailNaoImplementado>();
```

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
