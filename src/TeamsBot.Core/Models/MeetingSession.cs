using System;
using System.Collections.Generic;

namespace TeamsBot.Core.Models
{
    public enum MeetingStatus
    {
        Idle,
        Joining,
        Connected,
        Recording,
        Paused,
        Transcribing,
        Enhancing,
        Completed,
        Failed
    }

    public enum SpeakerType
    {
        User,
        Participant,
        Unknown
    }

    public class MeetingParticipant
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string DisplayName { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Role { get; set; } = "Participant";
        public uint AudioSourceId { get; set; }
        public DateTime JoinedAt { get; set; } = DateTime.UtcNow;
        public bool IsMuted { get; set; }
    }

    public class TranscriptSegment
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string MeetingId { get; set; } = string.Empty;
        public string SpeakerId { get; set; } = string.Empty;
        public string SpeakerName { get; set; } = "Participant";
        public SpeakerType SpeakerType { get; set; } = SpeakerType.Participant;
        public TimeSpan StartTime { get; set; }
        public TimeSpan EndTime { get; set; }
        public string TimestampFormatted { get; set; } = "00:00:00";
        public string Text { get; set; } = string.Empty;
        public double Confidence { get; set; } = 1.0;
    }

    public class ActionItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string MeetingId { get; set; } = string.Empty;
        public string Task { get; set; } = string.Empty;
        public string Assignee { get; set; } = "Team Member";
        public string DueDate { get; set; } = "Next Sync";
        public string Status { get; set; } = "PENDING"; // PENDING, IN_PROGRESS, COMPLETED
    }

    public class MeetingNotes
    {
        public string MeetingId { get; set; } = string.Empty;
        public string RawHumanNotes { get; set; } = string.Empty;
        public string EnhancedNotesMarkdown { get; set; } = string.Empty;
        public string Overview { get; set; } = string.Empty;
        public List<string> KeyPoints { get; set; } = new();
        public List<string> Decisions { get; set; } = new();
        public List<ActionItem> ActionItems { get; set; } = new();
        public string TemplatePreset { get; set; } = "General";
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
        public int Version { get; set; } = 1;
    }

    public class MeetingChatMessage
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string MeetingId { get; set; } = string.Empty;
        public string Role { get; set; } = "user"; // "user" or "assistant"
        public string Content { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    public class MeetingSession
    {
        public string SessionId { get; set; } = Guid.NewGuid().ToString();
        public string MeetingId { get; set; } = string.Empty;
        public string Subject { get; set; } = "Untitled Meeting";
        public string JoinUrl { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? EndedAt { get; set; }
        public int DurationSeconds { get; set; } = 0;
        public MeetingStatus Status { get; set; } = MeetingStatus.Idle;
        public string UserRequestedBy { get; set; } = string.Empty;
        public string TemplatePreset { get; set; } = "General";

        public List<MeetingParticipant> Participants { get; set; } = new();
        public List<TranscriptSegment> TranscriptSegments { get; set; } = new();
        public MeetingNotes Notes { get; set; } = new();
        public List<MeetingChatMessage> ChatMessages { get; set; } = new();

        public string RawAudioPath { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
    }

    public class SemanticSearchResult
    {
        public string MeetingId { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
        public double RelevanceScore { get; set; }
        public string MatchedSnippet { get; set; } = string.Empty;
        public string MatchSource { get; set; } = string.Empty; // "Notes" | "Transcript" | "Decisions" | "ActionItems"
    }
}
