using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using TeamsBot.Core.Models;
using TeamsBot.Server.Services;

namespace TeamsBot.Server.Controllers
{
    [ApiController]
    [Route("api/meetings")]
    public class MeetingController : ControllerBase
    {
        private static readonly ConcurrentDictionary<string, MeetingSession> _meetings = new();
        private static readonly object _saveLock = new();
        private readonly IAiSummaryService _aiSummaryService;
        private readonly IConfiguration _configuration;
        private static readonly string StorageFilePath = Path.Combine(AppContext.BaseDirectory, "granola_meetings_store.json");

        static MeetingController()
        {
            LoadMeetingsFromDisk();
        }

        public MeetingController(IAiSummaryService aiSummaryService, IConfiguration configuration)
        {
            _aiSummaryService = aiSummaryService;
            _configuration = configuration;
        }

        public class CreateOrUpdateMeetingRequest
        {
            public string? MeetingId { get; set; }
            public string Subject { get; set; } = "Untitled Meeting";
            public string UserEmail { get; set; } = "user@mailplus.com.au";
            public string TemplatePreset { get; set; } = "General";
            public string RawHumanNotes { get; set; } = string.Empty;
            public int DurationSeconds { get; set; } = 0;
            public MeetingStatus Status { get; set; } = MeetingStatus.Completed;
            public List<string>? Attendees { get; set; }
            public List<TranscriptSegment>? TranscriptSegments { get; set; }
        }

        public class UpdateNotesRequest
        {
            public string RawHumanNotes { get; set; } = string.Empty;
            public string? TemplatePreset { get; set; }
        }

        public class EnhanceNotesRequest
        {
            public string? MeetingSubject { get; set; }
            public string? RawHumanNotes { get; set; }
            public string? TemplatePreset { get; set; }
            public List<string>? Attendees { get; set; }
            public List<TranscriptSegment>? TranscriptSegments { get; set; }
            public string? TranscriptText { get; set; }
        }

        public class AppendChunkRequest
        {
            public string SpeakerName { get; set; } = "Participant";
            public SpeakerType SpeakerType { get; set; } = SpeakerType.Participant;
            public string TimestampFormatted { get; set; } = "00:00:00";
            public string Text { get; set; } = string.Empty;
            public double Confidence { get; set; } = 1.0;
        }

        public class MeetingChatRequest
        {
            public string Prompt { get; set; } = string.Empty;
        }

        public class SemanticSearchRequest
        {
            public string Query { get; set; } = string.Empty;
            public string? UserEmail { get; set; }
        }

        [HttpGet]
        public IActionResult GetAllMeetings([FromQuery] string? userEmail)
        {
            var list = _meetings.Values
                .Where(m => string.IsNullOrWhiteSpace(userEmail) || m.UserRequestedBy.Equals(userEmail.Trim(), StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(m => m.CreatedAt)
                .ToList();

            return Ok(new { success = true, count = list.Count, meetings = list });
        }

        [HttpGet("{id}")]
        public IActionResult GetMeetingById(string id)
        {
            if (_meetings.TryGetValue(id, out var meeting))
            {
                return Ok(new { success = true, meeting });
            }

            return NotFound(new { success = false, message = $"Meeting with ID '{id}' not found." });
        }

        [HttpPost]
        public IActionResult CreateOrUpdateMeeting([FromBody] CreateOrUpdateMeetingRequest request)
        {
            string id = string.IsNullOrWhiteSpace(request.MeetingId) ? Guid.NewGuid().ToString() : request.MeetingId;

            var meeting = _meetings.GetOrAdd(id, k => new MeetingSession
            {
                MeetingId = k,
                SessionId = k,
                CreatedAt = DateTime.UtcNow
            });

            meeting.Subject = string.IsNullOrWhiteSpace(request.Subject) ? "Untitled Meeting" : request.Subject.Trim();
            meeting.UserRequestedBy = string.IsNullOrWhiteSpace(request.UserEmail) ? "user@mailplus.com.au" : request.UserEmail.Trim().ToLowerInvariant();
            meeting.TemplatePreset = string.IsNullOrWhiteSpace(request.TemplatePreset) ? "General" : request.TemplatePreset;
            meeting.Status = request.Status;
            meeting.DurationSeconds = request.DurationSeconds;

            if (request.Attendees != null)
            {
                meeting.Participants = request.Attendees.Select(a => new MeetingParticipant { DisplayName = a }).ToList();
            }

            if (request.TranscriptSegments != null && request.TranscriptSegments.Count > 0)
            {
                meeting.TranscriptSegments = request.TranscriptSegments;
            }

            if (meeting.Notes == null) meeting.Notes = new MeetingNotes { MeetingId = id };
            if (!string.IsNullOrEmpty(request.RawHumanNotes))
            {
                meeting.Notes.RawHumanNotes = request.RawHumanNotes;
                meeting.Notes.UpdatedAt = DateTime.UtcNow;
            }

            PersistMeetingsToDisk();

            return Ok(new { success = true, meetingId = id, meeting });
        }

        [HttpPut("{id}/notes")]
        public IActionResult UpdateNotes(string id, [FromBody] UpdateNotesRequest request)
        {
            if (!_meetings.TryGetValue(id, out var meeting))
            {
                meeting = new MeetingSession
                {
                    MeetingId = id,
                    SessionId = id,
                    Subject = "Meeting Notes",
                    CreatedAt = DateTime.UtcNow
                };
                _meetings[id] = meeting;
            }

            if (meeting.Notes == null) meeting.Notes = new MeetingNotes { MeetingId = id };
            meeting.Notes.RawHumanNotes = request.RawHumanNotes ?? "";
            meeting.Notes.UpdatedAt = DateTime.UtcNow;
            meeting.Notes.Version++;
            if (!string.IsNullOrWhiteSpace(request.TemplatePreset))
            {
                meeting.TemplatePreset = request.TemplatePreset;
                meeting.Notes.TemplatePreset = request.TemplatePreset;
            }

            PersistMeetingsToDisk();
            return Ok(new { success = true, meetingId = id, version = meeting.Notes.Version });
        }

        [HttpPost("{id}/transcript/chunk")]
        public IActionResult AppendTranscriptChunk(string id, [FromBody] AppendChunkRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Text))
            {
                return BadRequest(new { success = false, message = "Transcript chunk text cannot be empty." });
            }

            var meeting = _meetings.GetOrAdd(id, k => new MeetingSession
            {
                MeetingId = k,
                SessionId = k,
                CreatedAt = DateTime.UtcNow
            });

            var segment = new TranscriptSegment
            {
                MeetingId = id,
                SpeakerName = string.IsNullOrWhiteSpace(request.SpeakerName) ? "Participant" : request.SpeakerName.Trim(),
                SpeakerType = request.SpeakerType,
                TimestampFormatted = request.TimestampFormatted,
                Text = request.Text.Trim(),
                Confidence = request.Confidence
            };

            meeting.TranscriptSegments.Add(segment);
            PersistMeetingsToDisk();

            return Ok(new { success = true, segmentId = segment.Id, count = meeting.TranscriptSegments.Count });
        }

        [HttpPost("{id}/enhance")]
        public async Task<IActionResult> EnhanceMeetingNotes(string id, [FromBody] EnhanceNotesRequest? request)
        {
            MeetingSession? meeting = null;
            _meetings.TryGetValue(id, out meeting);

            string subject = request?.MeetingSubject ?? meeting?.Subject ?? "Meeting Notes";
            string rawNotes = request?.RawHumanNotes ?? meeting?.Notes?.RawHumanNotes ?? "";
            string template = request?.TemplatePreset ?? meeting?.TemplatePreset ?? "General";
            var attendees = request?.Attendees ?? meeting?.Participants?.Select(p => p.DisplayName).ToList() ?? new List<string>();

            var segments = request?.TranscriptSegments;
            if ((segments == null || segments.Count == 0) && meeting?.TranscriptSegments != null)
            {
                segments = meeting.TranscriptSegments;
            }

            if ((segments == null || segments.Count == 0) && !string.IsNullOrWhiteSpace(request?.TranscriptText))
            {
                var synth = await _aiSummaryService.GenerateSummaryAsync(subject, request.TranscriptText);
                return Ok(new { success = true, enhancedSummary = synth });
            }

            var enhanced = await _aiSummaryService.EnhanceNotesAsync(subject, rawNotes, segments ?? new List<TranscriptSegment>(), template, attendees);

            if (meeting != null)
            {
                if (meeting.Notes == null) meeting.Notes = new MeetingNotes { MeetingId = id };
                meeting.Notes.EnhancedNotesMarkdown = enhanced.EnhancedMarkdown;
                meeting.Notes.Overview = enhanced.Overview;
                meeting.Notes.KeyPoints = enhanced.KeyPoints;
                meeting.Notes.Decisions = enhanced.Decisions;
                meeting.Notes.ActionItems = enhanced.ActionItems.Select(a => new ActionItem
                {
                    MeetingId = id,
                    Task = a.Task,
                    Assignee = a.Assignee,
                    DueDate = a.DueDate,
                    Status = a.Status
                }).ToList();
                meeting.Notes.TemplatePreset = template;
                meeting.Status = MeetingStatus.Completed;

                PersistMeetingsToDisk();
            }

            return Ok(new
            {
                success = true,
                meetingId = id,
                enhancedSummary = enhanced
            });
        }

        [HttpPost("enhance")]
        public async Task<IActionResult> EnhanceDirect([FromBody] EnhanceNotesRequest request)
        {
            string subject = string.IsNullOrWhiteSpace(request.MeetingSubject) ? "Direct Meeting Notes" : request.MeetingSubject.Trim();
            string rawNotes = request.RawHumanNotes ?? "";
            string template = string.IsNullOrWhiteSpace(request.TemplatePreset) ? "General" : request.TemplatePreset.Trim();
            var attendees = request.Attendees ?? new List<string>();
            var segments = request.TranscriptSegments ?? new List<TranscriptSegment>();

            var enhanced = await _aiSummaryService.EnhanceNotesAsync(subject, rawNotes, segments, template, attendees);
            return Ok(new { success = true, enhancedSummary = enhanced });
        }

        [HttpPost("{id}/chat")]
        public async Task<IActionResult> ChatWithMeeting(string id, [FromBody] MeetingChatRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Prompt))
            {
                return BadRequest(new { success = false, message = "Prompt cannot be empty." });
            }

            _meetings.TryGetValue(id, out var meeting);
            string subject = meeting?.Subject ?? "Meeting";
            string rawNotes = meeting?.Notes?.RawHumanNotes ?? "";
            string enhancedNotes = meeting?.Notes?.EnhancedNotesMarkdown ?? "";
            string transcriptText = meeting?.TranscriptSegments != null
                ? string.Join("\n", meeting.TranscriptSegments.Select(s => $"[{s.TimestampFormatted}] {s.SpeakerName}: {s.Text}"))
                : "";

            var history = meeting?.ChatMessages ?? new List<MeetingChatMessage>();

            var answer = await _aiSummaryService.ChatWithMeetingAsync(subject, request.Prompt, rawNotes, transcriptText, enhancedNotes, history);

            if (meeting != null)
            {
                meeting.ChatMessages.Add(new MeetingChatMessage { Role = "user", Content = request.Prompt, MeetingId = id });
                meeting.ChatMessages.Add(new MeetingChatMessage { Role = "assistant", Content = answer, MeetingId = id });
                PersistMeetingsToDisk();
            }

            return Ok(new
            {
                success = true,
                meetingId = id,
                role = "assistant",
                response = answer
            });
        }

        [HttpPost("search")]
        public async Task<IActionResult> SearchMeetings([FromBody] SemanticSearchRequest request)
        {
            var userMeetings = _meetings.Values
                .Where(m => string.IsNullOrWhiteSpace(request.UserEmail) || m.UserRequestedBy.Equals(request.UserEmail.Trim(), StringComparison.OrdinalIgnoreCase))
                .ToList();

            var searchResults = await _aiSummaryService.SemanticSearchAsync(request.Query, userMeetings);
            return Ok(new { success = true, count = searchResults.Count, results = searchResults });
        }

        [HttpPost("{id}/expand-notes")]
        public async Task<IActionResult> ExpandMeetingNotes(string id, [FromBody] EnhanceNotesRequest? request)
        {
            _meetings.TryGetValue(id, out var meeting);
            string subject = request?.MeetingSubject ?? meeting?.Subject ?? "Meeting Notes";
            string rawNotes = request?.RawHumanNotes ?? meeting?.Notes?.RawHumanNotes ?? "";
            string template = request?.TemplatePreset ?? meeting?.TemplatePreset ?? "General";
            var segments = request?.TranscriptSegments ?? meeting?.TranscriptSegments ?? new List<TranscriptSegment>();

            var expanded = await _aiSummaryService.ExpandNotesAsync(subject, rawNotes, segments, template);
            return Ok(new { success = true, meetingId = id, expandedNotes = expanded });
        }

        [HttpPost("expand-notes")]
        public async Task<IActionResult> ExpandNotesDirect([FromBody] EnhanceNotesRequest request)
        {
            string subject = string.IsNullOrWhiteSpace(request.MeetingSubject) ? "Meeting Notes" : request.MeetingSubject;
            string rawNotes = request.RawHumanNotes ?? "";
            string template = string.IsNullOrWhiteSpace(request.TemplatePreset) ? "General" : request.TemplatePreset;
            var segments = request.TranscriptSegments ?? new List<TranscriptSegment>();

            var expanded = await _aiSummaryService.ExpandNotesAsync(subject, rawNotes, segments, template);
            return Ok(new { success = true, expandedNotes = expanded });
        }

        [HttpPost("{id}/live-signals")]
        public async Task<IActionResult> GetLiveSignals(string id, [FromBody] List<TranscriptSegment>? recentSegments)
        {
            _meetings.TryGetValue(id, out var meeting);
            var segments = recentSegments ?? meeting?.TranscriptSegments ?? new List<TranscriptSegment>();

            var signals = await _aiSummaryService.ExtractLiveSignalsAsync(segments, meeting?.Subject ?? "");
            return Ok(new { success = true, meetingId = id, signals });
        }

        [HttpPost("live-signals")]
        public async Task<IActionResult> GetLiveSignalsDirect([FromBody] List<TranscriptSegment> recentSegments)
        {
            var signals = await _aiSummaryService.ExtractLiveSignalsAsync(recentSegments ?? new List<TranscriptSegment>(), "Live Call");
            return Ok(new { success = true, signals });
        }

        [HttpPost("{id}/tags")]
        public IActionResult UpdateMeetingTags(string id, [FromBody] List<string> tags)
        {
            if (_meetings.TryGetValue(id, out var meeting))
            {
                meeting.Tags = tags ?? new List<string>();
                PersistMeetingsToDisk();
                return Ok(new { success = true, meetingId = id, tags = meeting.Tags });
            }
            return NotFound(new { success = false, message = "Meeting not found." });
        }

        private static readonly ConcurrentDictionary<string, CustomTemplateDto> _customTemplates = new();
        private static readonly string TemplatesStorageFilePath = Path.Combine(AppContext.BaseDirectory, "granola_custom_templates.json");

        [HttpGet("/api/templates")]
        public IActionResult GetTemplates()
        {
            return Ok(new { success = true, templates = _customTemplates.Values.ToList() });
        }

        [HttpPost("/api/templates")]
        public IActionResult SaveTemplate([FromBody] CustomTemplateDto template)
        {
            if (string.IsNullOrWhiteSpace(template.Label))
            {
                return BadRequest(new { success = false, message = "Template label cannot be empty." });
            }

            string id = string.IsNullOrWhiteSpace(template.Id) ? "tmpl-" + Guid.NewGuid().ToString("N").Substring(0, 8) : template.Id;
            template.Id = id;
            template.IsCustom = true;
            template.CreatedAt = DateTime.UtcNow;

            _customTemplates[id] = template;
            PersistTemplatesToDisk();

            return Ok(new { success = true, template });
        }

        [HttpDelete("/api/templates/{id}")]
        public IActionResult DeleteTemplate(string id)
        {
            bool removed = _customTemplates.TryRemove(id, out _);
            if (removed)
            {
                PersistTemplatesToDisk();
                return Ok(new { success = true, message = "Template deleted." });
            }
            return NotFound(new { success = false, message = "Template not found." });
        }

        private static void PersistTemplatesToDisk()
        {
            try
            {
                var json = JsonSerializer.Serialize(_customTemplates.Values, new JsonSerializerOptions { WriteIndented = true });
                System.IO.File.WriteAllText(TemplatesStorageFilePath, json);
            }
            catch {}
        }

        [HttpDelete("{id}")]
        public IActionResult DeleteMeeting(string id)
        {
            bool removed = _meetings.TryRemove(id, out _);
            if (removed)
            {
                PersistMeetingsToDisk();
                return Ok(new { success = true, message = $"Meeting {id} deleted successfully." });
            }
            return NotFound(new { success = false, message = "Meeting not found." });
        }

        private static void PersistMeetingsToDisk()
        {
            lock (_saveLock)
            {
                try
                {
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    var json = JsonSerializer.Serialize(_meetings.Values, options);
                    System.IO.File.WriteAllText(StorageFilePath, json);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MeetingController] Warning saving meetings store: {ex.Message}");
                }
            }
        }

        private static void LoadMeetingsFromDisk()
        {
            lock (_saveLock)
            {
                try
                {
                    if (System.IO.File.Exists(StorageFilePath))
                    {
                        var json = System.IO.File.ReadAllText(StorageFilePath);
                        var loaded = JsonSerializer.Deserialize<List<MeetingSession>>(json);
                        if (loaded != null)
                        {
                            foreach (var m in loaded)
                            {
                                if (!string.IsNullOrWhiteSpace(m.MeetingId))
                                {
                                    _meetings[m.MeetingId] = m;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MeetingController] Warning loading meetings store: {ex.Message}");
                }
            }
        }
    }
}
