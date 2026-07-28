using System;
using System.IO;
using System.Threading.Tasks;
using TeamsBot.Core.Models;
using TeamsBot.Storage.Services;
using Xunit;

namespace TeamsBot.Tests.Unit
{
    public class StorageServiceTests
    {
        private readonly StorageService _storageService = new StorageService();

        [Fact]
        public void TestSanitizeFileName_StripsInvalidCharacters()
        {
            string rawSubject = "Weekly Sync: Sales / Marketing ? * < > | \"Test\"";
            string sanitized = _storageService.SanitizeFileName(rawSubject);

            Assert.DoesNotContain(":", sanitized);
            Assert.DoesNotContain("/", sanitized);
            Assert.DoesNotContain("?", sanitized);
        }

        [Fact]
        public void TestSanitizeFileName_EmptyFallback()
        {
            string sanitized = _storageService.SanitizeFileName("");
            Assert.StartsWith("Teams_Meeting_", sanitized);
        }

        [Fact]
        public async Task TestGenerateTranscriptExports_CreatesAllFormats()
        {
            var session = new MeetingSession
            {
                Subject = "Quarterly Strategy Review",
                Participants = { new MeetingParticipant { DisplayName = "Fiona Harrison" } },
                TranscriptSegments =
                {
                    new TranscriptSegment
                    {
                        SpeakerName = "Fiona Harrison",
                        StartTime = TimeSpan.FromSeconds(10),
                        EndTime = TimeSpan.FromSeconds(15),
                        Text = "Welcome to the quarterly strategy session."
                    }
                }
            };

            var exports = await _storageService.GenerateTranscriptExportsAsync(session);

            Assert.Equal(4, exports.Count);
        }
    }
}
