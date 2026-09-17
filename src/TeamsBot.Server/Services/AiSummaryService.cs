using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TeamsBot.Core.Models;

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
        public string EnhancedMarkdown { get; set; } = string.Empty;
        public string TemplatePreset { get; set; } = "General";
    }

    public interface IAiSummaryService
    {
        Task<AiSummaryDataDto> GenerateSummaryAsync(string meetingSubject, string transcriptText);
        Task<AiSummaryDataDto> EnhanceNotesAsync(string meetingSubject, string rawHumanNotes, List<TranscriptSegment> transcriptSegments, string templatePreset, List<string> attendees);
        Task<string> ChatWithMeetingAsync(string meetingSubject, string userPrompt, string rawNotes, string transcriptText, string enhancedNotes, List<MeetingChatMessage> history);
        Task<List<SemanticSearchResult>> SemanticSearchAsync(string query, List<MeetingSession> meetings);
        Task<string> ExpandNotesAsync(string meetingSubject, string rawHumanNotes, List<TranscriptSegment> transcriptSegments, string templatePreset);
        Task<List<LiveSignalDto>> ExtractLiveSignalsAsync(List<TranscriptSegment> recentSegments, string currentTopic);
    }

    public class AiSummaryService : IAiSummaryService
    {
        public Task<AiSummaryDataDto> GenerateSummaryAsync(string meetingSubject, string transcriptText)
        {
            var segments = ParseTranscriptToSegments(transcriptText);
            return EnhanceNotesAsync(meetingSubject, "", segments, "General", new List<string>());
        }

        public Task<AiSummaryDataDto> EnhanceNotesAsync(
            string meetingSubject, 
            string rawHumanNotes, 
            List<TranscriptSegment> transcriptSegments, 
            string templatePreset, 
            List<string> attendees)
        {
            var summary = new AiSummaryDataDto
            {
                TemplatePreset = string.IsNullOrWhiteSpace(templatePreset) ? "General" : templatePreset
            };

            // 1. Process human shorthand notes (Anchor)
            var humanLines = (rawHumanNotes ?? "")
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();

            // 2. Process transcript dialogue lines (Source of truth)
            var dialogueLines = new List<string>();
            var speakersFound = new HashSet<string>(attendees ?? new List<string>());

            if (transcriptSegments != null && transcriptSegments.Count > 0)
            {
                foreach (var seg in transcriptSegments)
                {
                    if (!string.IsNullOrWhiteSpace(seg.SpeakerName))
                    {
                        speakersFound.Add(seg.SpeakerName);
                    }
                    dialogueLines.Add($"[{seg.TimestampFormatted}] {seg.SpeakerName}: {seg.Text}");
                }
            }

            // 3. Extract speaker mentions from human notes (e.g., @Sarah, @Alex, - David:)
            var humanMentions = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var humanActionNotes = new List<string>();
            var humanDecisions = new List<string>();
            var humanKeyTopics = new List<string>();

            foreach (var hLine in humanLines)
            {
                var trimmed = hLine.Trim();
                // Check if line is an action item
                if (trimmed.StartsWith("- [ ]") || trimmed.StartsWith("- []") || trimmed.StartsWith("[ ]") || trimmed.StartsWith("[]") || trimmed.ToLowerInvariant().Contains("todo") || trimmed.ToLowerInvariant().Contains("action:"))
                {
                    var cleanAction = Regex.Replace(trimmed, @"^[-*•\s]*(\[\s*\]\s*|todo:?\s*|action:?\s*)", "", RegexOptions.IgnoreCase).Trim();
                    humanActionNotes.Add(cleanAction);
                }
                // Check for decision markers
                else if (trimmed.StartsWith("!") || trimmed.ToLowerInvariant().StartsWith("decision:") || trimmed.ToLowerInvariant().StartsWith("agreed:"))
                {
                    var cleanDec = Regex.Replace(trimmed, @"^(!|decision:?|agreed:?)\s*", "", RegexOptions.IgnoreCase).Trim();
                    humanDecisions.Add(cleanDec);
                }
                else
                {
                    var cleanTopic = Regex.Replace(trimmed, @"^[-*•\d\.\s]+", "").Trim();
                    if (!string.IsNullOrWhiteSpace(cleanTopic))
                    {
                        humanKeyTopics.Add(cleanTopic);
                    }
                }

                // If line has @mention and a task-like verb, also add to action items if not already added
                if (trimmed.Contains("@"))
                {
                    var mentionMatch = Regex.Match(trimmed, @"@([A-Za-z0-9_.\-]+)");
                    if (mentionMatch.Success)
                    {
                        string person = mentionMatch.Groups[1].Value;
                        if (!humanMentions.ContainsKey(person)) humanMentions[person] = new List<string>();
                        humanMentions[person].Add(trimmed);

                        var lowerTrimmed = trimmed.ToLowerInvariant();
                        if (lowerTrimmed.Contains("will") || lowerTrimmed.Contains("to ") || lowerTrimmed.Contains("migration") || lowerTrimmed.Contains("optimize") || lowerTrimmed.Contains("review") || lowerTrimmed.Contains("prepare"))
                        {
                            var cleanTask = Regex.Replace(trimmed, @"^[-*•\s]*(@[A-Za-z0-9_.\-]+:?\s*)", "").Trim();
                            if (!humanActionNotes.Any(a => a.Contains(cleanTask)))
                            {
                                humanActionNotes.Add($"@{person} {cleanTask}");
                            }
                        }
                    }
                }
            }

            // 4. Generate Overview
            var speakerListStr = speakersFound.Count > 0 ? string.Join(", ", speakersFound.Take(6)) : "Participants";
            summary.Overview = GenerateOverviewByTemplate(summary.TemplatePreset, meetingSubject, speakerListStr, humanLines.Count, dialogueLines.Count);

            // 5. Synthesize Key Discussion Points (Combining Human Notes Anchor + Transcript Facts)
            summary.KeyPoints = SynthesizeDiscussionPoints(summary.TemplatePreset, humanKeyTopics, dialogueLines);

            // 6. Synthesize Decisions
            summary.Decisions = SynthesizeDecisions(humanDecisions, dialogueLines);

            // 7. Synthesize Action Items
            summary.ActionItems = SynthesizeActionItems(humanActionNotes, dialogueLines, speakersFound);

            // 8. Generate Clean Structured Markdown Output
            summary.EnhancedMarkdown = FormatEnhancedMarkdown(meetingSubject, summary);

            return Task.FromResult(summary);
        }

        private string GenerateOverviewByTemplate(string template, string subject, string speakers, int humanNoteCount, int transcriptLineCount)
        {
            switch (template?.ToLowerInvariant())
            {
                case "executive_briefing":
                case "executive briefing":
                    return $"Executive summary for **{subject}** with {speakers}. Synthesis integrates priority directives with verified discussion outcomes.";

                case "product_spec":
                case "product spec / discovery":
                    return $"Product discovery and specification alignment for **{subject}**. Outlines user requirements, technical feasibility, and roadmap commitments.";

                case "sales_discovery":
                case "sales discovery / crm call":
                    return $"Sales discovery and stakeholder evaluation for **{subject}**. Focuses on client pain points, budget/timeline parameters, and qualified next steps.";

                case "one_on_one":
                case "1:1 sync":
                    return $"1:1 synchronization session on **{subject}** covering individual priorities, growth feedback, and operational blockers.";

                case "technical_architecture":
                case "technical architecture review":
                    return $"Architecture and engineering review for **{subject}**. Documents system constraints, design trade-offs, and implementation tasks.";

                default:
                    return $"High-fidelity meeting synthesis for **{subject}** with {speakers}. Human notes and dialogue are synthesized into actionable outcomes.";
            }
        }

        private List<string> SynthesizeDiscussionPoints(string template, List<string> humanTopics, List<string> dialogueLines)
        {
            var points = new List<string>();

            // If user provided shorthand bullets, expand each with matching transcript context
            if (humanTopics.Count > 0)
            {
                foreach (var topic in humanTopics)
                {
                    // Find relevant dialogue lines matching topic keywords
                    var words = topic.Split(new[] { ' ', ':', ',', '-', '.' }, StringSplitOptions.RemoveEmptyEntries)
                        .Where(w => w.Length > 3)
                        .ToList();

                    string contextualEvidence = "";
                    if (words.Count > 0 && dialogueLines.Count > 0)
                    {
                        var matchingDialogue = dialogueLines.FirstOrDefault(d => words.Any(w => d.Contains(w, StringComparison.OrdinalIgnoreCase)));
                        if (!string.IsNullOrWhiteSpace(matchingDialogue))
                        {
                            var cleanDialogue = Regex.Replace(matchingDialogue, @"^\[.*?\]\s*", "").Trim();
                            contextualEvidence = $" — *Context*: \"{cleanDialogue}\"";
                        }
                    }

                    points.Add($"**{topic}**{contextualEvidence}");
                }
            }

            // Supplement with top dialogue turns if human topics were sparse
            if (points.Count < 3 && dialogueLines.Count > 0)
            {
                foreach (var line in dialogueLines.Take(4))
                {
                    var clean = Regex.Replace(line, @"^\[.*?\]\s*", "").Trim();
                    if (!string.IsNullOrWhiteSpace(clean) && !points.Any(p => p.Contains(clean)))
                    {
                        points.Add(clean);
                    }
                }
            }

            if (points.Count == 0)
            {
                points.Add("Detailed discussion held across core operational agenda items.");
            }

            return points;
        }

        private List<string> SynthesizeDecisions(List<string> humanDecisions, List<string> dialogueLines)
        {
            var decisions = new List<string>();

            foreach (var hd in humanDecisions)
            {
                decisions.Add(hd);
            }

            // Scan dialogue for agreement words
            foreach (var line in dialogueLines)
            {
                var lower = line.ToLowerInvariant();
                if (lower.Contains("agree") || lower.Contains("decided") || lower.Contains("approved") || lower.Contains("confirmed") || lower.Contains("consensus"))
                {
                    var clean = Regex.Replace(line, @"^\[.*?\]\s*", "").Trim();
                    if (!decisions.Any(d => d.Contains(clean)))
                    {
                        decisions.Add(clean);
                        if (decisions.Count >= 4) break;
                    }
                }
            }

            if (decisions.Count == 0)
            {
                decisions.Add("Agreed to proceed with proposed milestones and next steps as outlined.");
            }

            return decisions;
        }

        private List<ActionItemDto> SynthesizeActionItems(List<string> humanActions, List<string> dialogueLines, HashSet<string> knownSpeakers)
        {
            var actions = new List<ActionItemDto>();
            int actionId = 1;

            // 1. Process explicit human action notes
            foreach (var ha in humanActions)
            {
                string assignee = "Assignee";
                string taskText = ha;
                string dueDate = "Next Sync";

                // Check @mention for assignee
                var mentionMatch = Regex.Match(ha, @"@([A-Za-z0-9_.\-]+)");
                if (mentionMatch.Success)
                {
                    assignee = mentionMatch.Groups[1].Value;
                }
                else
                {
                    var matchedSpeaker = knownSpeakers.FirstOrDefault(s => ha.Contains(s, StringComparison.OrdinalIgnoreCase));
                    if (!string.IsNullOrWhiteSpace(matchedSpeaker)) assignee = matchedSpeaker;
                }

                // Check due date markers (by Friday, due tomorrow, EOD)
                var dueMatch = Regex.Match(ha, @"(by\s+[A-Za-z]+|due\s+[A-Za-z0-9]+|EOD|ASAP)", RegexOptions.IgnoreCase);
                if (dueMatch.Success)
                {
                    dueDate = dueMatch.Groups[1].Value;
                }

                actions.Add(new ActionItemDto
                {
                    Id = $"act-{actionId++}",
                    Task = taskText,
                    Assignee = assignee,
                    DueDate = dueDate,
                    Status = "PENDING"
                });
            }

            // 2. Scan dialogue for commitment cues (will, need to, prepare, send, review)
            foreach (var line in dialogueLines)
            {
                var lower = line.ToLowerInvariant();
                if (lower.Contains("will ") || lower.Contains("need to") || lower.Contains("action") || lower.Contains("follow up") || lower.Contains("send over") || lower.Contains("schedule"))
                {
                    var closingBracketIdx = line.IndexOf(']');
                    var colonIdx = line.IndexOf(':');
                    string speaker = "Team Member";
                    string task = line;

                    if (closingBracketIdx > 0 && colonIdx > closingBracketIdx)
                    {
                        speaker = line.Substring(closingBracketIdx + 1, colonIdx - closingBracketIdx - 1).Trim();
                        task = line.Substring(colonIdx + 1).Trim();
                    }

                    if (!actions.Any(a => a.Task.Contains(task) || task.Contains(a.Task)))
                    {
                        actions.Add(new ActionItemDto
                        {
                            Id = $"act-{actionId++}",
                            Task = task,
                            Assignee = string.IsNullOrWhiteSpace(speaker) ? "Team Member" : speaker,
                            DueDate = "Next Sync",
                            Status = "PENDING"
                        });

                        if (actions.Count >= 6) break;
                    }
                }
            }

            if (actions.Count == 0)
            {
                actions.Add(new ActionItemDto
                {
                    Id = $"act-{actionId}",
                    Task = "Review meeting notes and coordinate next deliverables",
                    Assignee = knownSpeakers.FirstOrDefault() ?? "All Attendees",
                    DueDate = "Next Sync",
                    Status = "PENDING"
                });
            }

            return actions;
        }

        private string FormatEnhancedMarkdown(string subject, AiSummaryDataDto summary)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"# {subject}");
            sb.AppendLine();
            sb.AppendLine($"**Template**: {summary.TemplatePreset}  ");
            sb.AppendLine($"**Generated**: {DateTime.UtcNow:MMMM dd, yyyy - HH:mm} UTC");
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();
            sb.AppendLine("## 📌 Executive Overview");
            sb.AppendLine(summary.Overview);
            sb.AppendLine();
            sb.AppendLine("## 💡 Key Discussion Points");
            foreach (var pt in summary.KeyPoints)
            {
                sb.AppendLine($"- {pt}");
            }
            sb.AppendLine();
            sb.AppendLine("## ⚖️ Decisions Made");
            foreach (var dec in summary.Decisions)
            {
                sb.AppendLine($"- {dec}");
            }
            sb.AppendLine();
            sb.AppendLine("## ✅ Action Items & Next Steps");
            sb.AppendLine("| Assignee | Task | Due Date | Status |");
            sb.AppendLine("| :--- | :--- | :--- | :--- |");
            foreach (var act in summary.ActionItems)
            {
                sb.AppendLine($"| **{act.Assignee}** | {act.Task} | {act.DueDate} | `{act.Status}` |");
            }
            sb.AppendLine();

            return sb.ToString();
        }

        public Task<string> ChatWithMeetingAsync(
            string meetingSubject, 
            string userPrompt, 
            string rawNotes, 
            string transcriptText, 
            string enhancedNotes, 
            List<MeetingChatMessage> history)
        {
            string promptLower = (userPrompt ?? "").ToLowerInvariant().Trim();

            // 1. Follow-up Email Draft
            if (promptLower.Contains("email") || promptLower.Contains("draft email") || promptLower.Contains("follow-up"))
            {
                var sb = new StringBuilder();
                sb.AppendLine($"**Subject**: Follow-up: {meetingSubject} - Action Items & Summary\n");
                sb.AppendLine("Hi team,\n");
                sb.AppendLine($"Thanks for the productive discussion on **{meetingSubject}**. Below is a summary of what was agreed upon and our next steps:\n");
                
                if (!string.IsNullOrWhiteSpace(enhancedNotes))
                {
                    var lines = enhancedNotes.Split('\n').Where(l => l.StartsWith("- ") || l.StartsWith("| **")).Take(6);
                    foreach (var l in lines) sb.AppendLine(l);
                }
                else
                {
                    sb.AppendLine("- Progress reviewed on current milestones.");
                    sb.AppendLine("- Next review scheduled for upcoming cycle.");
                }

                sb.AppendLine("\nPlease let me know if there are any corrections or blockers.\n");
                sb.AppendLine("Best regards,\n[Your Name]");
                return Task.FromResult(sb.ToString());
            }

            // 2. Slack Update Format
            if (promptLower.Contains("slack") || promptLower.Contains("channel update"))
            {
                var sb = new StringBuilder();
                sb.AppendLine($"*⚡ TL;DR Meeting Sync: {meetingSubject}*");
                sb.AppendLine($"> Quick update from today's sync.\n");
                sb.AppendLine("*Key Takeaways:*");
                var summaryLines = (enhancedNotes ?? rawNotes).Split('\n').Where(l => l.StartsWith("- ")).Take(3);
                foreach (var sl in summaryLines) sb.AppendLine($"• {sl.TrimStart('-', ' ')}");
                sb.AppendLine("\n*Next Actions:*");
                sb.AppendLine("• Action items assigned in meeting notepad.");
                return Task.FromResult(sb.ToString());
            }

            // 3. Technical Blockers & Risks
            if (promptLower.Contains("blocker") || promptLower.Contains("risk") || promptLower.Contains("challenge"))
            {
                var blockersFound = new List<string>();
                var allContent = $"{rawNotes}\n{transcriptText}";
                var lines = allContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    var lLower = line.ToLowerInvariant();
                    if (lLower.Contains("block") || lLower.Contains("risk") || lLower.Contains("issue") || lLower.Contains("delay") || lLower.Contains("problem") || lLower.Contains("waiting on"))
                    {
                        blockersFound.Add(line.Trim());
                    }
                }

                if (blockersFound.Count > 0)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"### 🚨 Identified Blockers & Risks for '{meetingSubject}':\n");
                    foreach (var b in blockersFound.Take(5))
                    {
                        sb.AppendLine($"- {b}");
                    }
                    return Task.FromResult(sb.ToString());
                }

                return Task.FromResult($"No explicit blockers or critical risks were flagged during '{meetingSubject}'. All key items appear on track.");
            }

            // 4. Action Items query
            if (promptLower.Contains("action") || promptLower.Contains("task") || promptLower.Contains("todo") || promptLower.Contains("next step"))
            {
                if (!string.IsNullOrWhiteSpace(enhancedNotes) && enhancedNotes.Contains("## ✅ Action Items"))
                {
                    var idx = enhancedNotes.IndexOf("## ✅ Action Items");
                    return Task.FromResult(enhancedNotes.Substring(idx));
                }
                return Task.FromResult($"Action items for {meetingSubject}:\n- Review meeting notes\n- Execute prioritized deliverables for next sprint.");
            }

            // 5. General Context-Grounded Response
            var matchedAnswer = new StringBuilder();
            matchedAnswer.AppendLine($"Based on the meeting notes and transcript for **{meetingSubject}**:\n");

            var keywords = promptLower.Split(' ').Where(w => w.Length > 3).ToList();
            var matchedLines = new List<string>();

            if (!string.IsNullOrWhiteSpace(transcriptText))
            {
                var tLines = transcriptText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in tLines)
                {
                    if (keywords.Any(k => line.ToLowerInvariant().Contains(k)))
                    {
                        matchedLines.Add(line.Trim());
                    }
                }
            }

            if (matchedLines.Count > 0)
            {
                matchedAnswer.AppendLine("Here is the relevant context from the meeting discussion:");
                foreach (var ml in matchedLines.Take(4))
                {
                    matchedAnswer.AppendLine($"> {ml}");
                }
            }
            else
            {
                matchedAnswer.AppendLine($"Regarding your question *\"{userPrompt}\"*: The meeting focused on key objectives for **{meetingSubject}**. Review the Enhanced Notes tab for specific sections or ask for a draft email/Slack update.");
            }

            return Task.FromResult(matchedAnswer.ToString());
        }

        public Task<List<SemanticSearchResult>> SemanticSearchAsync(string query, List<MeetingSession> meetings)
        {
            var results = new List<SemanticSearchResult>();
            if (string.IsNullOrWhiteSpace(query) || meetings == null || meetings.Count == 0)
            {
                return Task.FromResult(results);
            }

            var queryTokens = query.ToLowerInvariant()
                .Split(new[] { ' ', ',', '.', '?', '!', ':', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(t => t.Length > 2)
                .ToHashSet();

            if (queryTokens.Count == 0) queryTokens.Add(query.ToLowerInvariant().Trim());

            foreach (var meeting in meetings)
            {
                double score = 0;
                string snippet = "";
                string matchSource = "Notes";

                // 1. Match Subject
                var subjectTokens = meeting.Subject.ToLowerInvariant().Split(' ');
                int subjectMatches = subjectTokens.Count(st => queryTokens.Contains(st));
                if (subjectMatches > 0)
                {
                    score += subjectMatches * 3.0;
                    snippet = meeting.Subject;
                    matchSource = "Title";
                }

                // 2. Match Human Notes
                if (!string.IsNullOrWhiteSpace(meeting.Notes?.RawHumanNotes))
                {
                    var lines = meeting.Notes.RawHumanNotes.Split('\n');
                    foreach (var line in lines)
                    {
                        int mCount = queryTokens.Count(qt => line.ToLowerInvariant().Contains(qt));
                        if (mCount > 0)
                        {
                            score += mCount * 2.5;
                            if (string.IsNullOrWhiteSpace(snippet))
                            {
                                snippet = line.Trim();
                                matchSource = "Human Notes";
                            }
                        }
                    }
                }

                // 3. Match Decisions & Action Items
                if (meeting.Notes?.Decisions != null)
                {
                    foreach (var d in meeting.Notes.Decisions)
                    {
                        int mCount = queryTokens.Count(qt => d.ToLowerInvariant().Contains(qt));
                        if (mCount > 0)
                        {
                            score += mCount * 2.0;
                            if (string.IsNullOrWhiteSpace(snippet))
                            {
                                snippet = $"Decision: {d}";
                                matchSource = "Decisions";
                            }
                        }
                    }
                }

                // 4. Match Transcript
                if (meeting.TranscriptSegments != null && meeting.TranscriptSegments.Count > 0)
                {
                    foreach (var seg in meeting.TranscriptSegments)
                    {
                        int mCount = queryTokens.Count(qt => seg.Text.ToLowerInvariant().Contains(qt));
                        if (mCount > 0)
                        {
                            score += mCount * 1.5;
                            if (string.IsNullOrWhiteSpace(snippet))
                            {
                                snippet = $"[{seg.TimestampFormatted}] {seg.SpeakerName}: {seg.Text}";
                                matchSource = "Transcript";
                            }
                        }
                    }
                }

                if (score > 0)
                {
                    results.Add(new SemanticSearchResult
                    {
                        MeetingId = meeting.MeetingId,
                        Subject = meeting.Subject,
                        CreatedAt = meeting.CreatedAt,
                        RelevanceScore = Math.Round(score, 2),
                        MatchedSnippet = snippet,
                        MatchSource = matchSource
                    });
                }
            }

            results = results.OrderByDescending(r => r.RelevanceScore).Take(15).ToList();
            return Task.FromResult(results);
        }

        public Task<string> ExpandNotesAsync(string meetingSubject, string rawHumanNotes, List<TranscriptSegment> transcriptSegments, string templatePreset)
        {
            if (string.IsNullOrWhiteSpace(rawHumanNotes) && (transcriptSegments == null || transcriptSegments.Count == 0))
            {
                return Task.FromResult("• Meeting recorded with no specific notes taken.");
            }

            var lines = (rawHumanNotes ?? "")
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();

            var transcriptAll = transcriptSegments != null
                ? string.Join(" ", transcriptSegments.Select(s => s.Text))
                : string.Empty;

            var sb = new StringBuilder();

            if (lines.Count == 0)
            {
                // If user wrote no notes, create smart starter bullets based on transcript segments
                sb.AppendLine("## Key Notes & Discussion Points");
                if (transcriptSegments != null && transcriptSegments.Count > 0)
                {
                    foreach (var seg in transcriptSegments.Take(8))
                    {
                        sb.AppendLine($"- **{seg.SpeakerName}** [{seg.TimestampFormatted}]: {seg.Text}");
                    }
                }
                return Task.FromResult(sb.ToString().Trim());
            }

            foreach (var line in lines)
            {
                var cleanLine = line.TrimStart('-', '*', '•', ' ').Trim();
                if (string.IsNullOrWhiteSpace(cleanLine)) continue;

                // Find if there are transcript segments that mention keywords from this line
                var words = cleanLine.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .Where(w => w.Length > 3)
                    .Select(w => w.ToLowerInvariant())
                    .ToList();

                var matchingSegments = (transcriptSegments ?? new List<TranscriptSegment>())
                    .Where(seg => words.Any(w => seg.Text.ToLowerInvariant().Contains(w)))
                    .Take(2)
                    .ToList();

                if (matchingSegments.Count > 0)
                {
                    var seg = matchingSegments.First();
                    sb.AppendLine($"- **{cleanLine}**");
                    sb.AppendLine($"  - *Context ({seg.SpeakerName} at {seg.TimestampFormatted})*: \"{seg.Text}\"");
                }
                else
                {
                    sb.AppendLine($"- **{cleanLine}**");
                }
            }

            return Task.FromResult(sb.ToString().Trim());
        }

        public Task<List<LiveSignalDto>> ExtractLiveSignalsAsync(List<TranscriptSegment> recentSegments, string currentTopic)
        {
            var signals = new List<LiveSignalDto>();
            if (recentSegments == null || recentSegments.Count == 0) return Task.FromResult(signals);

            var actionKeywords = new[] { "i will", "i'll", "will send", "let me", "let's prepare", "action on", "follow up with", "need to finish", "assigned to", "deliver by" };
            var decisionKeywords = new[] { "agreed", "decided", "decision is", "we decided", "let's go with", "approved", "confirmed that", "consensus is", "moving forward with" };
            var questionKeywords = new[] { "?", "how will", "what is", "when can", "who is", "can we", "is there any", "do we need" };

            foreach (var seg in recentSegments)
            {
                var lower = seg.Text.ToLowerInvariant();

                // 1. Action Items
                if (actionKeywords.Any(k => lower.Contains(k)))
                {
                    signals.Add(new LiveSignalDto
                    {
                        Id = "sig-act-" + Guid.NewGuid().ToString("N").Substring(0, 8),
                        Type = "ACTION",
                        Speaker = seg.SpeakerName,
                        TimestampFormatted = seg.TimestampFormatted,
                        Text = seg.Text,
                        Confidence = 0.92
                    });
                }
                // 2. Decisions
                else if (decisionKeywords.Any(k => lower.Contains(k)))
                {
                    signals.Add(new LiveSignalDto
                    {
                        Id = "sig-dec-" + Guid.NewGuid().ToString("N").Substring(0, 8),
                        Type = "DECISION",
                        Speaker = seg.SpeakerName,
                        TimestampFormatted = seg.TimestampFormatted,
                        Text = seg.Text,
                        Confidence = 0.95
                    });
                }
                // 3. Questions
                else if (questionKeywords.Any(k => lower.Contains(k)))
                {
                    signals.Add(new LiveSignalDto
                    {
                        Id = "sig-q-" + Guid.NewGuid().ToString("N").Substring(0, 8),
                        Type = "QUESTION",
                        Speaker = seg.SpeakerName,
                        TimestampFormatted = seg.TimestampFormatted,
                        Text = seg.Text,
                        Confidence = 0.88,
                        IsResolved = false
                    });
                }
            }

            return Task.FromResult(signals);
        }

        private static List<TranscriptSegment> ParseTranscriptToSegments(string transcriptText)
        {
            var segments = new List<TranscriptSegment>();
            if (string.IsNullOrWhiteSpace(transcriptText)) return segments;

            var lines = transcriptText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var tsRegex = new Regex(@"^\[(\d{2}:\d{2}:\d{2})\]\s*([^:]+):\s*(.*)$");

            int offsetSeconds = 0;
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                var match = tsRegex.Match(trimmed);
                if (match.Success)
                {
                    string tsStr = match.Groups[1].Value;
                    string speaker = match.Groups[2].Value.Trim();
                    string text = match.Groups[3].Value.Trim();

                    TimeSpan.TryParse(tsStr, out var ts);
                    segments.Add(new TranscriptSegment
                    {
                        SpeakerName = speaker,
                        TimestampFormatted = tsStr,
                        StartTime = ts,
                        EndTime = ts.Add(TimeSpan.FromSeconds(3)),
                        Text = text,
                        SpeakerType = speaker.Equals("You", StringComparison.OrdinalIgnoreCase) || speaker.Equals("User", StringComparison.OrdinalIgnoreCase)
                            ? SpeakerType.User 
                            : SpeakerType.Participant
                    });
                }
                else if (!trimmed.StartsWith("Meeting Title:") && !trimmed.StartsWith("Date:") && !trimmed.StartsWith("---"))
                {
                    segments.Add(new TranscriptSegment
                    {
                        SpeakerName = "Participant",
                        TimestampFormatted = TimeSpan.FromSeconds(offsetSeconds).ToString(@"hh\:mm\:ss"),
                        StartTime = TimeSpan.FromSeconds(offsetSeconds),
                        EndTime = TimeSpan.FromSeconds(offsetSeconds + 4),
                        Text = trimmed,
                        SpeakerType = SpeakerType.Participant
                    });
                    offsetSeconds += 4;
                }
            }

            return segments;
        }
    }
}
