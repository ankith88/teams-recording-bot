using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using TeamsBot.Core.Interfaces;
using TeamsBot.Server.Services;
using TeamsBot.Signaling.Controllers;
using TeamsBot.Storage.Services;
using TeamsBot.Transcription.Services;

var builder = WebApplication.CreateBuilder(args);

builder.WebHost.UseUrls("http://localhost:5001", "http://localhost:5000");

builder.Services.AddHttpClient();
builder.Services.AddSingleton<ITranscriptionService, LocalWhisperTranscriptionService>();
builder.Services.AddSingleton<IStorageService, StorageService>();
builder.Services.AddSingleton<CallSignalingService>();

builder.Services.AddSingleton<IEmailService, MsGraphEmailService>();
builder.Services.AddSingleton<IAuthService, AuthService>();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

builder.Services.AddControllers();

var app = builder.Build();

app.UseCors("AllowAll");

app.MapGet("/", () => "Teams Recording Bot Server is running!");
app.MapControllers();

app.Run();
