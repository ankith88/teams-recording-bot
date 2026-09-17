using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TeamsBot.Core.Models;
using TeamsBot.Server.Services;
using Xunit;

namespace TeamsBot.Tests.Unit
{
    public class GranolaAiNotepadTests
    {
        private readonly AiSummaryService _aiService = new();

        [Fact]
        public async Task EnhanceNotesAsync_CombinesHumanNotesAndTranscriptCorrectly()
        {
            // Arrange
            string subject = "Q3 Infrastructure Planning";
            string rawHumanNotes = 
                "- @Sarah: optimize API latency under 200ms\n" +
                "- @Alex: database migration this Friday\n" +
                "- ! Decided to release beta on Tuesday\n" +
                "- [ ] Prepare deployment runbook by Friday";

            var segments = new List<TranscriptSegment>
            {
                new() { SpeakerName = "Sarah", Text = "We need our API latency to remain below 200ms under load.", TimestampFormatted = "00:01:10" },
                new() { SpeakerName = "Alex", Text = "I will perform the database indexing migration this Friday night.", TimestampFormatted = "00:02:40" },
                new() { SpeakerName = "Sarah", Text = "Everyone agrees we launch the beta version next Tuesday.", TimestampFormatted = "00:04:00" }
            };

            var attendees = new List<string> { "Sarah", "Alex" };

            // Act
            var result = await _aiService.EnhanceNotesAsync(subject, rawHumanNotes, segments, "Technical Architecture Review", attendees);

            // Assert
            Assert.NotNull(result);
            Assert.Contains("Q3 Infrastructure Planning", result.EnhancedMarkdown);
            Assert.True(result.ActionItems.Count >= 2);
            Assert.Contains(result.ActionItems, a => a.Assignee.Equals("Sarah", StringComparison.OrdinalIgnoreCase) || a.Assignee.Equals("Alex", StringComparison.OrdinalIgnoreCase));
            Assert.NotEmpty(result.Decisions);
            Assert.NotEmpty(result.KeyPoints);
        }

        [Fact]
        public async Task ChatWithMeetingAsync_GeneratesFollowUpEmailAndSlackFormat()
        {
            // Arrange
            string subject = "Design System Review";
            string rawNotes = "- Agreed on dark theme palette\n- Sarah to export SVG icons";
            string transcript = "[00:00:10] Sarah: We finalized the dark theme palette today.";
            string enhancedNotes = "# Design System Review\n## 💡 Key Discussion Points\n- Dark theme approved\n## ✅ Action Items\n| **Sarah** | Export SVG icons | EOD |";

            // Act: Follow-up email
            var emailResponse = await _aiService.ChatWithMeetingAsync(
                subject, 
                "Draft a follow-up email", 
                rawNotes, 
                transcript, 
                enhancedNotes, 
                new List<MeetingChatMessage>()
            );

            // Act: Slack update
            var slackResponse = await _aiService.ChatWithMeetingAsync(
                subject, 
                "Summarize for Slack channel", 
                rawNotes, 
                transcript, 
                enhancedNotes, 
                new List<MeetingChatMessage>()
            );

            // Assert
            Assert.Contains("Follow-up:", emailResponse);
            Assert.Contains("Design System Review", emailResponse);
            Assert.Contains("TL;DR", slackResponse);
        }

        [Fact]
        public async Task SemanticSearchAsync_RanksRelevantHistoricalMeetings()
        {
            // Arrange
            var meetings = new List<MeetingSession>
            {
                new()
                {
                    MeetingId = "m1",
                    Subject = "Mobile App Payment Gateway Integration",
                    Notes = new MeetingNotes
                    {
                        RawHumanNotes = "- Stripe payment gateway integration discussed with finance team",
                        Decisions = new List<string> { "Selected Stripe as default provider" }
                    }
                },
                new()
                {
                    MeetingId = "m2",
                    Subject = "Office Relocation Planning",
                    Notes = new MeetingNotes
                    {
                        RawHumanNotes = "- Review lease contract and desk allocations"
                    }
                }
            };

            // Act
            var searchResults = await _aiService.SemanticSearchAsync("Stripe payment gateway", meetings);

            // Assert
            Assert.NotEmpty(searchResults);
            Assert.Equal("m1", searchResults[0].MeetingId);
            Assert.True(searchResults[0].RelevanceScore > 0);
        }
    }
}
