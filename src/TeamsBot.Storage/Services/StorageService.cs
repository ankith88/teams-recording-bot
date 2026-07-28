using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TeamsBot.Core.Interfaces;
using TeamsBot.Core.Models;

namespace TeamsBot.Storage.Services
{
    public class StorageService : IStorageService
    {
        private static readonly Regex InvalidCharsRegex = new Regex(@"[\\/:*?""<>|]", RegexOptions.Compiled);

        public string SanitizeFileName(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return $"Teams_Meeting_{DateTime.Now:yyyy-MM-dd_HHmm}";
            }

            // Remove invalid characters
            string sanitized = InvalidCharsRegex.Replace(subject.Trim(), "_");

            // Collapse multiple underscores or spaces
            sanitized = Regex.Replace(sanitized, @"\s+", " ");
            sanitized = Regex.Replace(sanitized, @"_+", "_");

            // Trim leading/trailing dots or spaces
            sanitized = sanitized.Trim('.', ' ', '_');

            // Fallback if empty after sanitization
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                return $"Teams_Meeting_{DateTime.Now:yyyy-MM-dd_HHmm}";
            }

            // Limit maximum length for OS compatibility
            if (sanitized.Length > 100)
            {
                sanitized = sanitized.Substring(0, 100).Trim();
            }

            return sanitized;
        }

        public string ResolveUniqueFilePath(string directoryPath, string baseFileName, string extension)
        {
            string ext = extension.StartsWith(".") ? extension : "." + extension;
            string fullPath = Path.Combine(directoryPath, $"{baseFileName}{ext}");

            if (!File.Exists(fullPath))
            {
                return fullPath;
            }

            int counter = 1;
            while (File.Exists(fullPath))
            {
                string numberedFileName = $"{baseFileName} ({counter}){ext}";
                fullPath = Path.Combine(directoryPath, numberedFileName);
                counter++;
            }

            return fullPath;
        }

        public Task<List<ExportedFileInfo>> GenerateTranscriptExportsAsync(MeetingSession session)
        {
            var files = new List<ExportedFileInfo>();
            string baseFileName = SanitizeFileName(session.Subject);

            // 1. Plain Text (.txt)
            files.Add(new ExportedFileInfo
            {
                Format = "txt",
                FileName = $"{baseFileName}.txt",
                Content = Encoding.UTF8.GetBytes(GeneratePlainText(session))
            });

            // 2. JSON Format (.json)
            files.Add(new ExportedFileInfo
            {
                Format = "json",
                FileName = $"{baseFileName}.json",
                Content = Encoding.UTF8.GetBytes(GenerateJson(session))
            });

            // 3. SRT Format (.srt)
            files.Add(new ExportedFileInfo
            {
                Format = "srt",
                FileName = $"{baseFileName}.srt",
                Content = Encoding.UTF8.GetBytes(GenerateSrt(session))
            });

            // 4. Formatted Markdown / Text representation for DOCX (.docx)
            files.Add(new ExportedFileInfo
            {
                Format = "docx",
                FileName = $"{baseFileName}.docx",
                Content = Encoding.UTF8.GetBytes(GenerateFormattedDocxContent(session))
            });

            return Task.FromResult(files);
        }

        private string GeneratePlainText(MeetingSession session)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Meeting Title: {session.Subject}");
            sb.AppendLine($"Date: {session.CreatedAt:yyyy-MM-dd HH:mm:ss UTC}");
            sb.AppendLine($"Session ID: {session.SessionId}");
            sb.AppendLine(new string('-', 60));
            sb.AppendLine();

            foreach (var segment in session.TranscriptSegments)
            {
                string startTimeStr = segment.StartTime.ToString(@"hh\:mm\:ss");
                string speaker = string.IsNullOrWhiteSpace(segment.SpeakerName) ? "Unknown Speaker" : segment.SpeakerName;
                sb.AppendLine($"[{startTimeStr}] {speaker}: {segment.Text}");
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private string GenerateJson(MeetingSession session)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };
            return JsonSerializer.Serialize(session, options);
        }

        private string GenerateSrt(MeetingSession session)
        {
            var sb = new StringBuilder();
            int index = 1;

            foreach (var segment in session.TranscriptSegments)
            {
                string start = FormatSrtTime(segment.StartTime);
                string end = FormatSrtTime(segment.EndTime > segment.StartTime ? segment.EndTime : segment.StartTime.Add(TimeSpan.FromSeconds(3)));
                string speaker = string.IsNullOrWhiteSpace(segment.SpeakerName) ? "Speaker" : segment.SpeakerName;

                sb.AppendLine(index.ToString());
                sb.AppendLine($"{start} --> {end}");
                sb.AppendLine($"{speaker}: {segment.Text}");
                sb.AppendLine();

                index++;
            }

            return sb.ToString();
        }

        private string FormatSrtTime(TimeSpan ts)
        {
            return $"{ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2},{ts.Milliseconds:D3}";
        }

        private string GenerateFormattedDocxContent(MeetingSession session)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"# {session.Subject}");
            sb.AppendLine($"**Date**: {session.CreatedAt:MMMM dd, yyyy - HH:mm:ss} UTC  ");
            sb.AppendLine($"**Participants**: {session.Participants.Count} attendees  ");
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();
            sb.AppendLine("## Meeting Transcript");
            sb.AppendLine();

            foreach (var segment in session.TranscriptSegments)
            {
                string speaker = string.IsNullOrWhiteSpace(segment.SpeakerName) ? "Unknown Speaker" : segment.SpeakerName;
                string time = segment.StartTime.ToString(@"hh\:mm\:ss");
                sb.AppendLine($"### [{time}] {speaker}");
                sb.AppendLine($"> {segment.Text}");
                sb.AppendLine();
            }

            return sb.ToString();
        }
    }
}
