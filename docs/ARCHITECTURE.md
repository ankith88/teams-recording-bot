# Technical Architecture & Platform Specifications

## 1. Overview

The Teams Meeting Recording & Transcription system provides an automated meeting assistant that joins Microsoft Teams calls, records audio, attributes speakers in real time, and exports timestamped transcripts directly to the user's laptop.

---

## 2. Platform Constraints & Technical Justification

### A. Windows Server Dependency for Real-Time Media Engine
- The official Microsoft Graph Real-Time Media SDK (`Microsoft.Graph.Communications.Calls.Media`) relies on native C++ Win32 audio processing libraries (`MediaPlatform.dll`).
- **Impact**: The bot backend engine must run on a Windows OS host (Windows Server 2022, Windows 10/11, or Windows Server Containers).
- **Solution**: The backend engine is deployed to a central Windows Cloud VM / Server. End-users (whether on macOS or Windows) interact with the bot through the **ProspectPlus Web Dashboard**, eliminating any requirement for end-users to install Windows software locally.

### B. Speaker Attribution Mechanism
- Standard audio recorders receive a single mixed audio track, requiring complex machine-learning diarization models to guess speaker changes.
- In contrast, the Microsoft Graph Calling SDK provides **individual per-participant audio streams** (`AudioMediaStream`).
- Each stream is tagged with the speaker's Microsoft account identity (`DisplayName`, `User.Id`).
- **Result**: Speaker attribution is 100% accurate per audio stream channel without ML voice recognition errors.

---

## 3. Data Flow Architecture

```
[Microsoft Teams Call]
       |
       | Real-Time PCM Audio Streams & Graph Events
       v
[Central Bot Backend - C# .NET 8]
       |
       |-- 1. Graph Webhook Signal Controller (/api/calls)
       |-- 2. Audio Socket Media Processor (16kHz 16-bit PCM)
       |-- 3. Local Whisper.net Transcription Engine
       |
       | Live WebSocket / SignalR Transcript Stream
       v
[ProspectPlus Web Application - Next.js]
       |
       |-- 1. Microsoft Entra ID Single Sign-On
       |-- 2. Live Transcript Display Component
       |-- 3. Browser File System Access API (showDirectoryPicker)
       v
[User's Laptop Hard Drive] (Mac / Windows)
  ├── Weekly Sync.docx
  ├── Weekly Sync.srt
  ├── Weekly Sync.json
  └── Weekly Sync.txt
```

---

## 4. Local-First Security & Data Egress Guarantee

1. **Zero Cloud Transcript Retention**: Transcripts are streamed over encrypted WebSockets directly to the authenticated user's browser session and written to disk on their laptop. No transcripts are persisted in cloud databases.
2. **Local Transcription Engine**: Default transcription runs locally on the bot server using `Whisper.net` (whisper.cpp), ensuring audio never leaves your private infrastructure to third-party APIs.
3. **Non-Clobbering Storage**: The `StorageService` sanitizes meeting titles (removing `\ / : * ? " < > |`) and appends incremental counters (`Weekly Sync (1).docx`) to avoid overwriting existing files.
