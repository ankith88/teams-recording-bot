using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using TeamsBot.Core.Interfaces;
using TeamsBot.Signaling.Controllers;
using TeamsBot.Storage.Services;
using TeamsBot.Transcription.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton<ITranscriptionService, LocalWhisperTranscriptionService>();
builder.Services.AddSingleton<IStorageService, StorageService>();
builder.Services.AddSingleton<CallSignalingService>();
builder.Services.AddControllers();

var app = builder.Build();

app.MapGet("/", () => "Teams Recording Bot Server is running!");
app.MapControllers();

app.Run();
