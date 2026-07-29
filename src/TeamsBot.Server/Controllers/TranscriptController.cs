using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using TeamsBot.Server.Services;

namespace TeamsBot.Server.Controllers
{
    [ApiController]
    [Route("api/transcript")]
    public class TranscriptController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly HttpClient _httpClient;
        private readonly IAiSummaryService _aiSummaryService;

        public TranscriptController(IConfiguration configuration, HttpClient httpClient, IAiSummaryService aiSummaryService)
        {
            _configuration = configuration;
            _httpClient = httpClient;
            _aiSummaryService = aiSummaryService;
        }

        public class FetchTranscriptRequest
        {
            public string JoinUrl { get; set; } = string.Empty;
            public string UserEmail { get; set; } = string.Empty;
            public string Subject { get; set; } = string.Empty;
        }

        [HttpPost("fetch")]
        public async Task<IActionResult> FetchTranscript([FromBody] FetchTranscriptRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.JoinUrl))
            {
                return BadRequest(new { success = false, message = "Teams Meeting Join Link or Meeting ID is required." });
            }

            string userEmail = string.IsNullOrWhiteSpace(request.UserEmail) ? "ankith.ravindran@mailplus.com.au" : request.UserEmail.Trim().ToLowerInvariant();
            string meetingSubject = string.IsNullOrWhiteSpace(request.Subject) ? "Teams Meeting" : request.Subject.Trim();

            try
            {
                // 1. Get Azure AD Access Token for Graph API
                string? accessToken = await GetGraphAccessTokenAsync();
                if (string.IsNullOrWhiteSpace(accessToken))
                {
                    return Ok(new
                    {
                        success = false,
                        message = "Azure Entra ID configuration is missing or invalid. Please check AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in backend configuration."
                    });
                }

                // 2. Resolve Online Meeting ID from Graph API
                string? meetingId = await ResolveMeetingIdAsync(accessToken, userEmail, request.JoinUrl);
                if (string.IsNullOrWhiteSpace(meetingId))
                {
                    return Ok(new
                    {
                        success = false,
                        message = $"Could not locate an active online meeting on Microsoft Teams for URL/ID '{request.JoinUrl}'. Ensure the meeting URL is valid and hosted on Microsoft Teams."
                    });
                }

                // 3. Fetch Transcript Metadata List from Graph API
                string? transcriptId = await GetLatestTranscriptIdAsync(accessToken, userEmail, meetingId);
                if (string.IsNullOrWhiteSpace(transcriptId))
                {
                    return Ok(new
                    {
                        success = false,
                        message = $"No recorded transcript was found on Microsoft Teams for meeting '{meetingSubject}'. Please make sure 'Start Transcription' was turned on during your Teams call."
                    });
                }

                // 4. Download Raw VTT Transcript Content
                string? rawVttContent = await DownloadTranscriptVttAsync(accessToken, userEmail, meetingId, transcriptId);
                if (string.IsNullOrWhiteSpace(rawVttContent))
                {
                    return Ok(new
                    {
                        success = false,
                        message = "Transcript file was found on Microsoft Teams, but content could not be retrieved from Microsoft Graph API."
                    });
                }

                // 5. Parse WebVTT content into clean speaker dialogue text
                string formattedDialogue = ParseVttToDialogueText(rawVttContent);
                string timestampStr = DateTime.UtcNow.ToString("dd/MM/yyyy, h:mm:ss tt", CultureInfo.InvariantCulture);

                string plainTextContent = $"Meeting Title: {meetingSubject}\n" +
                    $"Date: {timestampStr}\n" +
                    $"Authenticated MailPlus User: {userEmail}\n" +
                    $"--------------------------------------------------\n\n" +
                    formattedDialogue;

                // 6. Generate Real AI Summary
                var aiSummary = await _aiSummaryService.GenerateSummaryAsync(meetingSubject, plainTextContent);

                return Ok(new
                {
                    success = true,
                    meetingSubject,
                    plainTextContent,
                    aiSummary,
                    dateSaved = timestampStr,
                    source = "microsoft_graph_api"
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TranscriptController] Error fetching transcript: {ex}");
                return StatusCode(500, new
                {
                    success = false,
                    message = $"Failed to fetch transcript from Microsoft Graph API: {ex.Message}"
                });
            }
        }

        private async Task<string?> GetGraphAccessTokenAsync()
        {
            string tenantId = _configuration["AZURE_TENANT_ID"] ?? _configuration["AzureAd:TenantId"] ?? "";
            string clientId = _configuration["AZURE_CLIENT_ID"] ?? _configuration["AzureAd:ClientId"] ?? "";
            string clientSecret = _configuration["AZURE_CLIENT_SECRET"] ?? _configuration["AzureAd:ClientSecret"] ?? "";

            if (string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(clientSecret))
            {
                return null;
            }

            var tokenUrl = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";
            var tokenContent = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", "client_credentials"),
                new KeyValuePair<string, string>("client_id", clientId),
                new KeyValuePair<string, string>("client_secret", clientSecret),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default")
            });

            var tokenResponse = await _httpClient.PostAsync(tokenUrl, tokenContent);
            if (!tokenResponse.IsSuccessStatusCode)
            {
                string err = await tokenResponse.Content.ReadAsStringAsync();
                Console.WriteLine($"[TranscriptController] Token Error ({tokenResponse.StatusCode}): {err}");
                return null;
            }

            var tokenJson = await tokenResponse.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(tokenJson);
            return doc.RootElement.GetProperty("access_token").GetString();
        }

        private async Task<string?> ResolveMeetingIdAsync(string accessToken, string userEmail, string joinUrl)
        {
            try
            {
                // If input looks like an explicit meeting ID already
                if (!joinUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase) && joinUrl.Length > 10 && !joinUrl.Contains("/"))
                {
                    return joinUrl;
                }

                string encodedUrl = Uri.EscapeDataString(joinUrl);
                string graphUrl = $"https://graph.microsoft.com/v1.0/users/{userEmail}/onlineMeetings?$filter=joinWebUrl eq '{encodedUrl}'";

                var request = new HttpRequestMessage(HttpMethod.Get, graphUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await _httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("value", out var valArr) && valArr.GetArrayLength() > 0)
                    {
                        var first = valArr[0];
                        if (first.TryGetProperty("id", out var idProp))
                        {
                            return idProp.GetString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TranscriptController] Meeting lookup error: {ex.Message}");
            }

            return null;
        }

        private async Task<string?> GetLatestTranscriptIdAsync(string accessToken, string userEmail, string meetingId)
        {
            try
            {
                string graphUrl = $"https://graph.microsoft.com/v1.0/users/{userEmail}/onlineMeetings/{meetingId}/transcripts";

                var request = new HttpRequestMessage(HttpMethod.Get, graphUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await _httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("value", out var valArr) && valArr.GetArrayLength() > 0)
                    {
                        var first = valArr[0];
                        if (first.TryGetProperty("id", out var idProp))
                        {
                            return idProp.GetString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TranscriptController] Transcript list error: {ex.Message}");
            }

            return null;
        }

        private async Task<string?> DownloadTranscriptVttAsync(string accessToken, string userEmail, string meetingId, string transcriptId)
        {
            try
            {
                string graphUrl = $"https://graph.microsoft.com/v1.0/users/{userEmail}/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content?$format=text/vtt";

                var request = new HttpRequestMessage(HttpMethod.Get, graphUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await _httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TranscriptController] Download VTT error: {ex.Message}");
            }

            return null;
        }

        private static string ParseVttToDialogueText(string vttContent)
        {
            if (string.IsNullOrWhiteSpace(vttContent)) return "";

            var lines = vttContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();

            string currentTimestamp = "00:00:00";
            var timestampRegex = new Regex(@"(\d{2}:\d{2}:\d{2})\.\d{3}\s*-->");
            var voiceTagRegex = new Regex(@"<v\s+([^>]+)>(.*?)(</v>|$)", RegexOptions.Singleline);

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed == "WEBVTT" || trimmed.StartsWith("NOTE") || int.TryParse(trimmed, out _))
                {
                    continue;
                }

                var matchTs = timestampRegex.Match(trimmed);
                if (matchTs.Success)
                {
                    currentTimestamp = matchTs.Groups[1].Value;
                    continue;
                }

                var matchVoice = voiceTagRegex.Match(trimmed);
                if (matchVoice.Success)
                {
                    string speakerName = matchVoice.Groups[1].Value.Trim();
                    string text = matchVoice.Groups[2].Value.Trim();
                    sb.AppendLine($"[{currentTimestamp}] {speakerName}: {text}\n");
                }
                else if (!string.IsNullOrWhiteSpace(trimmed) && !trimmed.Contains("-->"))
                {
                    sb.AppendLine($"[{currentTimestamp}] Participant: {trimmed}\n");
                }
            }

            return sb.ToString().TrimEnd();
        }
    }
}
