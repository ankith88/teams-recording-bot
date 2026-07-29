using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TeamsBot.Server.Services
{
    public class ActionItemDto
    {
        public string Id { get; set; } = string.Empty;
        public string Task { get; set; } = string.Empty;
        public string Assignee { get; set; } = string.Empty;
        public string DueDate { get; set; } = "As discussed";
        public string Status { get; set; } = "PENDING";
    }

    public class AiSummaryDataDto
    {
        public string Overview { get; set; } = string.Empty;
        public List<string> KeyPoints { get; set; } = new();
        public List<ActionItemDto> ActionItems { get; set; } = new();
        public List<string> Decisions { get; set; } = new();
    }

    public interface IAiSummaryService
    {
        Task<AiSummaryDataDto> GenerateSummaryAsync(string meetingSubject, string transcriptText);
    }

    public class AiSummaryService : IAiSummaryService
    {
        public Task<AiSummaryDataDto> GenerateSummaryAsync(string meetingSubject, string transcriptText)
        {
            var summary = new AiSummaryDataDto();

            if (string.IsNullOrWhiteSpace(transcriptText))
            {
                summary.Overview = $"No transcript text was recorded for meeting '{meetingSubject}'.";
                return Task.FromResult(summary);
            }

            var lines = transcriptText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var dialogueLines = new List<string>();
            var speakersFound = new HashSet<string>();

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("Meeting Title:") || trimmed.StartsWith("Date:") || trimmed.StartsWith("Authenticated MailPlus") || trimmed.StartsWith("---"))
                {
                    continue;
                }

                if (trimmed.StartsWith("[") && trimmed.Contains("]"))
                {
                    dialogueLines.Add(trimmed);
                    
                    // Extract speaker name [00:01:15] Speaker Name: Spoken text
                    var colonIdx = trimmed.IndexOf(':');
                    var closingBracketIdx = trimmed.IndexOf(']');
                    if (closingBracketIdx > 0 && colonIdx > closingBracketIdx)
                    {
                        var speaker = trimmed.Substring(closingBracketIdx + 1, colonIdx - closingBracketIdx - 1).Trim();
                        if (!string.IsNullOrWhiteSpace(speaker))
                        {
                            speakersFound.Add(speaker);
                        }
                    }
                }
            }

            var speakerListStr = speakersFound.Count > 0 ? string.Join(", ", speakersFound) : "Participants";
            summary.Overview = $"Meeting transcript analysis for '{meetingSubject}' with active participants: {speakerListStr}. Total recorded statements: {dialogueLines.Count}.";

            // Extract Key Points
            int maxKeyPoints = Math.Min(5, dialogueLines.Count);
            for (int i = 0; i < maxKeyPoints; i++)
            {
                var cleanLine = Regex.Replace(dialogueLines[i], @"^\[.*?\]\s*", "");
                summary.KeyPoints.Add(cleanLine);
            }

            if (summary.KeyPoints.Count == 0)
            {
                summary.KeyPoints.Add($"Discussion conducted regarding {meetingSubject}.");
            }

            // Extract Action Items based on action keywords (will, need to, prepare, send, update, follow up, action)
            int actionId = 1;
            foreach (var line in dialogueLines)
            {
                var lower = line.ToLowerInvariant();
                if (lower.Contains("will ") || lower.Contains("need to") || lower.Contains("action") || lower.Contains("prepare") || lower.Contains("send") || lower.Contains("follow up"))
                {
                    var closingBracketIdx = line.IndexOf(']');
                    var colonIdx = line.IndexOf(':');
                    string assignee = "Team Member";
                    string taskText = line;

                    if (closingBracketIdx > 0 && colonIdx > closingBracketIdx)
                    {
                        assignee = line.Substring(closingBracketIdx + 1, colonIdx - closingBracketIdx - 1).Trim();
                        taskText = line.Substring(colonIdx + 1).Trim();
                    }

                    summary.ActionItems.Add(new ActionItemDto
                    {
                        Id = $"act-{actionId++}",
                        Task = taskText,
                        Assignee = string.IsNullOrWhiteSpace(assignee) ? "Team Member" : assignee,
                        DueDate = "Next Sync",
                        Status = "PENDING"
                    });

                    if (summary.ActionItems.Count >= 5) break;
                }
            }

            // Extract Key Decisions
            foreach (var line in dialogueLines)
            {
                var lower = line.ToLowerInvariant();
                if (lower.Contains("agree") || lower.Contains("decide") || lower.Contains("approved") || lower.Contains("confirm") || lower.Contains("next steps"))
                {
                    var cleanLine = Regex.Replace(line, @"^\[.*?\]\s*", "");
                    summary.Decisions.Add(cleanLine);
                    if (summary.Decisions.Count >= 3) break;
                }
            }

            if (summary.Decisions.Count == 0)
            {
                summary.Decisions.Add($"Participants agreed to proceed with meeting objectives for {meetingSubject}.");
            }

            return Task.FromResult(summary);
        }
    }
}
