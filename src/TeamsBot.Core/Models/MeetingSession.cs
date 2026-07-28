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
        Transcribing,
        Completed,
        Failed
    }

    public class MeetingParticipant
    {
        public string Id { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public uint AudioSourceId { get; set; }
        public DateTime JoinedAt { get; set; } = DateTime.UtcNow;
        public bool IsMuted { get; set; }
    }

    public class TranscriptSegment
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string SpeakerId { get; set; } = string.Empty;
        public string SpeakerName { get; set; } = string.Empty;
        public TimeSpan StartTime { get; set; }
        public TimeSpan EndTime { get; set; }
        public string Text { get; set; } = string.Empty;
        public double Confidence { get; set; } = 1.0;
    }

    public class MeetingSession
    {
        public string SessionId { get; set; } = Guid.NewGuid().ToString();
        public string MeetingId { get; set; } = string.Empty;
        public string Subject { get; set; } = "Untitled Meeting";
        public string JoinUrl { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? EndedAt { get; set; }
        public MeetingStatus Status { get; set; } = MeetingStatus.Idle;
        public string UserRequestedBy { get; set; } = string.Empty;
        
        public List<MeetingParticipant> Participants { get; set; } = new();
        public List<TranscriptSegment> TranscriptSegments { get; set; } = new();
        public string RawAudioPath { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
