// =============================================================================
//  Program.cs  —  Host Blazor Server do app de demonstração (DemoBBiCore)
//  Registra os componentes interativos, as configurações de e-mail do sistema
//  e o enviador. Harness para exercitar a RCL BBiCore; não faz parte da lib.
// =============================================================================

using BBiCore.Email;
using DemoBBiCore.Components;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// -----------------------------------------------------------------------------
// Configurações de e-mail: cada sistema tem o SEU grupo (remetente, usuário
// sistêmico, senha, servidor). Aqui vêm do appsettings ("Email"); em produção,
// a senha deve vir de um cofre/segredo, não do arquivo.
// -----------------------------------------------------------------------------
OpcoesEmail opcoesEmail = builder.Configuration.GetSection("Email").Get<OpcoesEmail>() ?? new OpcoesEmail();

// Caminhos relativos do appsettings resolvidos a partir da raiz do app.
if (!string.IsNullOrWhiteSpace(opcoesEmail.PastaBaseAnexos) && !Path.IsPathRooted(opcoesEmail.PastaBaseAnexos))
    opcoesEmail.PastaBaseAnexos = Path.Combine(builder.Environment.ContentRootPath, opcoesEmail.PastaBaseAnexos);

// Pasta onde o enviador simulado grava os .eml (para inspeção nesta demo).
opcoesEmail.PastaSimulacao = Path.Combine(builder.Environment.ContentRootPath, "emails-simulados");

builder.Services.AddSingleton(opcoesEmail);

// Envio centralizado. Nesta demo, o SIMULADO grava o .eml em disco (sem servidor).
// Em produção, troque por EnviadorEmailExchange — mesma interface, mesma configuração.
builder.Services.AddScoped<IEnviadorEmail, EnviadorEmailSimulado>();
// builder.Services.AddScoped<IEnviadorEmail, EnviadorEmailExchange>();

// Persistência de templates. Nesta demo, em memória (some ao reiniciar o app).
// Em produção, implemente IRepositorioTemplateEmail com EF Core na tabela do sistema.
builder.Services.AddSingleton<IRepositorioTemplateEmail, RepositorioTemplateEmailMemoria>();

var app = builder.Build();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
