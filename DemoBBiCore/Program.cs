// =============================================================================
//  Program.cs  —  Host Blazor Server do app de demonstração (DemoBBiCore)
//  Registra os componentes interativos de servidor e mapeia o App raiz.
//  Harness para exercitar a RCL BBiCore; não faz parte da biblioteca.
// =============================================================================

using DemoBBiCore.Components;
using BBiCore.Email;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Envio de e-mail centralizado: registre UMA implementação de IEnviadorEmail.
// Troque o stub pela implementação real (Exchange) quando existir.
builder.Services.AddScoped<IEnviadorEmail, EnviadorEmailNaoImplementado>();

var app = builder.Build();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
