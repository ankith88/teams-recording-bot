# Microsoft Teams Meeting Recording & Transcription Bot

A production-quality application that joins Microsoft Teams meetings as an automated bot, records audio in real time, and produces timestamped, speaker-attributed transcripts stored directly on the user's laptop.

Designed for **Mail Plus Pty Ltd** team members, integrated with **ProspectPlus**, supporting **macOS** and **Windows** without requiring end-users to install bot software locally.

---

## 🌟 Key Features

- **No Dependency on Host**: Bot joins meetings independently via a Teams Join URL or Meeting ID.
- **Zero Local Downloads for Team Members**: End-users access the transcriber via the ProspectPlus Web UI on macOS Finder or Windows File Explorer.
- **100% Accurate Speaker Attribution**: Utilizes per-participant audio stream sockets from the Microsoft Graph Calling SDK to accurately attribute spoken text to specific Microsoft account users (`[00:01:15 - Jane Doe]: ...`).
- **Local-First Storage**: Transcripts (`.docx`, `.srt`, `.json`, `.txt`) write directly to the user's laptop using the browser's native **File System Access API** (`window.showDirectoryPicker()`). Zero transcripts are stored in the cloud.
- **Compliance & Recording Notification**: Plays an automated audio notification (*"This meeting is being recorded and transcribed for team notes"*) upon joining.
- **Sanitized & Non-Clobbering Filenames**: Automatically cleans meeting titles into valid OS filenames (`Weekly Sync.docx`) and appends numbers (`Weekly Sync (1).docx`) if a file already exists.
- **Engine Abstraction**: Default local offline transcription via **`Whisper.net`** (whisper.cpp), with optional configuration switching to Azure Speech API or OpenAI.

---

## 🏗 System Architecture

```
+-----------------------------------------------------------------------------------------+
| CENTRALIZED BOT SERVER (Hosted on Windows Cloud VM / Server)                            |
|                                                                                         |
|  - C# .NET 8 Web API + Microsoft.Graph.Communications.Calls.Media                       |
|  - Listens on TCP/UDP Audio Socket for Real-Time PCM Frames                             |
|  - Local Whisper.net Engine for Offline Transcription & Speaker Channel Attribution     |
+--------------------------------------------+--------------------------------------------+
                                             | WebSockets / SignalR
                                             v
+-----------------------------------------------------------------------------------------+
| END-USER LAPTOP (Mac / Windows Browser via ProspectPlus Web App)                         |
|                                                                                         |
|  - Route: https://prospectplus.com.au/teams-recorder                                    |
|  - Authentication: Microsoft SSO (Entra ID)                                             |
|  - Local Storage: Browser HTML5 File System Access API (macOS Finder / Windows Explorer) |
|  - Saves: .docx, .srt, .json, .txt directly into chosen laptop directory               |
+-----------------------------------------------------------------------------------------+
```

---

## 🔐 Azure Entra ID (Azure AD) Setup Guide

### 1. App Registration
1. In [Azure Portal](https://portal.azure.com), go to **Microsoft Entra ID** -> **App registrations** -> **New registration**.
2. **Name**: `Teams Meeting Recording Bot`
3. **Supported account types**: Select **`Single tenant only - Mail Plus Pty Ltd`**.
4. **Redirect URI** (Select **Web**):
   - Local Dev: `http://localhost:3000/api/auth/callback/azure-ad`
   - Production: `https://prospectplus.com.au/api/auth/callback/azure-ad`
5. Click **Register**. Copy the **Application (Client) ID** and **Directory (Tenant) ID**.

### 2. Client Secret
1. Go to **Certificates & secrets** -> **New client secret**.
2. Description: `ProspectPlusBotSecret`, Expiration: `730 days`.
3. Copy the **Secret Value**.

### 3. API Permissions & Admin Consent
1. Go to **API permissions** -> **Add a permission** -> **Microsoft Graph**.
2. **Delegated Permissions**: `User.Read`, `Calendars.Read`, `openid`, `profile`, `email`.
3. **Application Permissions**:
   - `Calls.AccessMedia.All`
   - `Calls.JoinGroupCall.All`
   - `Calls.JoinGroupCallAsGuest.All`
   - `OnlineMeetings.Read.All`
4. Click **"Grant admin consent for Mail Plus Pty Ltd"** and confirm (**Yes**).

---

## 🚀 Quick Start (Running Locally)

### 1. Configure Environment Variables
Copy `.env.example` to `.env` and fill in your Azure credentials:
```bash
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_ID=your_client_id_here
AZURE_CLIENT_SECRET=your_client_secret_here
BOT_BASE_URL=https://your-dev-tunnel.ngrok-free.app
```

### 2. Start Dev Tunnel (For Local Graph Webhooks)
```bash
devtunnel host -p 5001 --protocol https
```

### 3. Run the Bot Backend Service
```bash
cd src/TeamsBot.Server
dotnet run
```

### 4. Trigger a Recording from ProspectPlus
1. Open ProspectPlus in browser (`http://localhost:3000/teams-recorder`).
2. Sign in with your `@mailplus.com.au` account.
3. Paste a Teams meeting URL and click **Pick Destination Folder** to select your laptop directory.
4. Click **Start Recording**. When finished, click **Stop & Save Transcript**.
