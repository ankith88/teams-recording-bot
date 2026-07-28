# Build Stage
FROM mcr.microsoft.com/dotnet/sdk:10.0-preview AS build
WORKDIR /src

# Copy project files and restore dependencies
COPY ["src/TeamsBot.Core/TeamsBot.Core.csproj", "src/TeamsBot.Core/"]
COPY ["src/TeamsBot.Storage/TeamsBot.Storage.csproj", "src/TeamsBot.Storage/"]
COPY ["src/TeamsBot.Transcription/TeamsBot.Transcription.csproj", "src/TeamsBot.Transcription/"]
COPY ["src/TeamsBot.Signaling/TeamsBot.Signaling.csproj", "src/TeamsBot.Signaling/"]
COPY ["src/TeamsBot.Server/TeamsBot.Server.csproj", "src/TeamsBot.Server/"]

RUN dotnet restore "src/TeamsBot.Server/TeamsBot.Server.csproj"

# Copy full source and publish
COPY . .
WORKDIR "/src/src/TeamsBot.Server"
RUN dotnet publish "TeamsBot.Server.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Runtime Stage
FROM mcr.microsoft.com/dotnet/aspnet:10.0-preview AS final
WORKDIR /app
COPY --from=build /app/publish .

EXPOSE 8080
ENV ASPNETCORE_URLS=http://+:8080

ENTRYPOINT ["dotnet", "TeamsBot.Server.dll"]
