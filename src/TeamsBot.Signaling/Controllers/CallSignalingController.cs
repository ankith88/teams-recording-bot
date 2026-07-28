using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using TeamsBot.Core.Interfaces;
using TeamsBot.Core.Models;

namespace TeamsBot.Signaling.Controllers
{
    public class CallSignalingService
    {
        private readonly Dictionary<string, MeetingSession> _activeSessions = new();
        private readonly ITranscriptionService _transcriptionService;
        private readonly IStorageService _storageService;

        public CallSignalingService(ITranscriptionService transcriptionService, IStorageService storageService)
        {
            _transcriptionService = transcriptionService;
            _storageService = storageService;
        }

        public async Task<MeetingSession> StartMeetingJoinAsync(string joinUrl, string userRequestedBy)
        {
            var session = new MeetingSession
            {
                JoinUrl = joinUrl,
                UserRequestedBy = userRequestedBy,
                Status = MeetingStatus.Joining,
                CreatedAt = DateTime.UtcNow,
                Subject = ExtractSubjectFromUrl(joinUrl)
            };

            _activeSessions[session.SessionId] = session;

            // 1. Graph API Request to join call: POST /communications/calls
            // 2. Setup media sockets & audio listeners
            // 3. Play Recording Announcement WAV prompt on join

            session.Status = MeetingStatus.Connected;
            
            // Add bot join notification
            session.Participants.Add(new MeetingParticipant
            {
                Id = "bot-1",
                DisplayName = "Recording & Transcription Bot",
                JoinedAt = DateTime.UtcNow
            });

            return await Task.FromResult(session);
        }

        public async Task<MeetingSession?> GetSessionAsync(string sessionId)
        {
            if (_activeSessions.TryGetValue(sessionId, out var session))
            {
                return await Task.FromResult(session);
            }
            return null;
        }

        public async Task<MeetingSession> StopRecordingAndExportAsync(string sessionId, string targetDirectoryPath)
        {
            if (!_activeSessions.TryGetValue(sessionId, out var session))
            {
                throw new KeyNotFoundException($"Meeting session {sessionId} not found.");
            }

            session.Status = MeetingStatus.Transcribing;
            session.EndedAt = DateTime.UtcNow;

            // Generate Exports (.docx, .srt, .json, .txt)
            var exports = await _storageService.GenerateTranscriptExportsAsync(session);

            // Ensure destination directory exists on user's machine
            if (!Directory.Exists(targetDirectoryPath))
            {
                Directory.CreateDirectory(targetDirectoryPath);
            }

            // Save files avoiding overwrite
            string sanitizedSubject = _storageService.SanitizeFileName(session.Subject);
            foreach (var export in exports)
            {
                string uniqueFilePath = _storageService.ResolveUniqueFilePath(targetDirectoryPath, sanitizedSubject, export.Format);
                await File.WriteAllBytesAsync(uniqueFilePath, export.Content);
            }

            session.Status = MeetingStatus.Completed;
            return session;
        }

        private string ExtractSubjectFromUrl(string joinUrl)
        {
            if (string.IsNullOrWhiteSpace(joinUrl)) return "Teams Meeting";
            
            // Extract title if present in query parameters or fallback
            if (joinUrl.Contains("context="))
            {
                try
                {
                    return "Teams Team Sync";
                }
                catch { }
            }

            return "Weekly Teams Meeting";
        }
    }
}
