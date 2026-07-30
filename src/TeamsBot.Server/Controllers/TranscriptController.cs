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
                var resolvedInfo = await ResolveMeetingInfoAsync(accessToken, userEmail, request.JoinUrl);
                if (resolvedInfo == null || string.IsNullOrWhiteSpace(resolvedInfo.MeetingId))
                {
                    return Ok(new
                    {
                        success = false,
                        message = $"Could not locate an active online meeting on Microsoft Teams for URL/ID '{request.JoinUrl}'. Ensure the meeting URL is valid and hosted on Microsoft Teams."
                    });
                }

                string meetingId = resolvedInfo.MeetingId;
                string effectiveEmail = resolvedInfo.EffectiveUserEmail;
                if (!string.IsNullOrWhiteSpace(resolvedInfo.MeetingSubject))
                {
                    meetingSubject = resolvedInfo.MeetingSubject;
                }

                // 3. Fetch Transcript Metadata List from Graph API (try effectiveEmail, then fallback to userEmail)
                string? transcriptId = await GetLatestTranscriptIdAsync(accessToken, effectiveEmail, meetingId);
                if (string.IsNullOrWhiteSpace(transcriptId) && !effectiveEmail.Equals(userEmail, StringComparison.OrdinalIgnoreCase))
                {
                    transcriptId = await GetLatestTranscriptIdAsync(accessToken, userEmail, meetingId);
                    if (!string.IsNullOrWhiteSpace(transcriptId))
                    {
                        effectiveEmail = userEmail;
                    }
                }

                if (string.IsNullOrWhiteSpace(transcriptId))
                {
                    return Ok(new
                    {
                        success = false,
                        message = $"No recorded transcript was found on Microsoft Teams for meeting '{meetingSubject}'. Please make sure 'Start Transcription' was turned on during your Teams call."
                    });
                }

                // 4. Download Raw VTT Transcript Content
                string? rawVttContent = await DownloadTranscriptVttAsync(accessToken, effectiveEmail, meetingId, transcriptId);
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

        [HttpGet("debug")]
        public async Task<IActionResult> DebugLookup([FromQuery] string? userEmail, [FromQuery] string? joinUrl)
        {
            string email = string.IsNullOrWhiteSpace(userEmail) ? "claude.busse@mailplus.com.au" : userEmail.Trim().ToLowerInvariant();
            string targetUrl = string.IsNullOrWhiteSpace(joinUrl)
                ? "https://teams.microsoft.com/l/meetup-join/19%3ameeting_NzI4ZDE5NjktNGEzOC00YWE1LTgwYzYtZWYwMDRjYjZhZmQ4%40thread.v2/0?context=%7b%22Tid%22%3a%22e7b892da-d63d-410e-8aba-3e936bb7838d%22%2c%22Oid%22%3a%22e6516f0b-df7a-4180-aee9-7aa5944d7c02%22%7d"
                : joinUrl.Trim();

            string? accessToken = await GetGraphAccessTokenAsync();
            if (accessToken == null) return BadRequest(new { error = "Failed to obtain Graph API access token" });

            string? userGuid = await GetUserGuidAsync(accessToken, email);

            try
            {
                var results = new Dictionary<string, object>();
                results["userGuid"] = userGuid ?? "null";

                if (!string.IsNullOrWhiteSpace(userGuid))
                {
                    // 1. Try URL encoded
                    string safeUrl1 = targetUrl.Replace("'", "''");
                    string omUrl1 = $"https://graph.microsoft.com/v1.0/users/{userGuid}/onlineMeetings?$filter=joinWebUrl eq '{safeUrl1}'";
                    var omReq1 = new HttpRequestMessage(HttpMethod.Get, omUrl1);
                    omReq1.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    var omResp1 = await _httpClient.SendAsync(omReq1);
                    results["encoded_url_response"] = await omResp1.Content.ReadAsStringAsync();

                    // 2. Try URL unescaped
                    string unescaped = Uri.UnescapeDataString(targetUrl);
                    string safeUrl2 = unescaped.Replace("'", "''");
                    string omUrl2 = $"https://graph.microsoft.com/v1.0/users/{userGuid}/onlineMeetings?$filter=joinWebUrl eq '{safeUrl2}'";
                    var omReq2 = new HttpRequestMessage(HttpMethod.Get, omUrl2);
                    omReq2.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    var omResp2 = await _httpClient.SendAsync(omReq2);
                    results["unescaped_url_response"] = await omResp2.Content.ReadAsStringAsync();
                }

                return Ok(results);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
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

        public class ResolvedMeetingInfo
        {
            public string MeetingId { get; set; } = string.Empty;
            public string EffectiveUserEmail { get; set; } = string.Empty;
            public string MeetingSubject { get; set; } = string.Empty;
        }

        private async Task<ResolvedMeetingInfo?> ResolveMeetingInfoAsync(string accessToken, string userEmail, string joinUrl)
        {
            try
            {
                joinUrl = joinUrl.Trim();

                // 1. If input looks like an explicit meeting ID already (no http/slashes)
                if (!joinUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase) && joinUrl.Length > 10 && !joinUrl.Contains("/"))
                {
                    return new ResolvedMeetingInfo { MeetingId = joinUrl, EffectiveUserEmail = userEmail };
                }

                var candidateUrls = new List<string> { joinUrl };

                // 2. Extract numeric meeting ID if URL is in /meet/123456 format (e.g. https://teams.microsoft.com/meet/413353030648656?p=...)
                string extractedNumericId = string.Empty;
                var meetMatch = Regex.Match(joinUrl, @"/meet/(\d+)");
                if (meetMatch.Success)
                {
                    extractedNumericId = meetMatch.Groups[1].Value;
                }
                else if (Regex.IsMatch(joinUrl, @"^\d{9,15}$"))
                {
                    extractedNumericId = joinUrl;
                }

                // 3. Follow HTTP redirect to expand /meet/ short URLs to full meetup-join URL
                try
                {
                    using var handler = new HttpClientHandler { AllowAutoRedirect = true };
                    using var redirectClient = new HttpClient(handler);
                    redirectClient.Timeout = TimeSpan.FromSeconds(4);
                    var req = new HttpRequestMessage(HttpMethod.Get, joinUrl);
                    req.Headers.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
                    var resp = await redirectClient.SendAsync(req);
                    string finalUrl = resp.RequestMessage?.RequestUri?.ToString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(finalUrl) && !finalUrl.Equals(joinUrl, StringComparison.OrdinalIgnoreCase))
                    {
                        candidateUrls.Add(finalUrl);
                        if (finalUrl.Contains("?"))
                        {
                            candidateUrls.Add(finalUrl.Split('?')[0]);
                        }
                    }
                }
                catch (Exception exRedirect)
                {
                    Console.WriteLine($"[TranscriptController] Redirect resolution warning: {exRedirect.Message}");
                }

                // 4. Add unescaped and query-stripped variations
                string unescaped = Uri.UnescapeDataString(joinUrl);
                if (!candidateUrls.Contains(unescaped)) candidateUrls.Add(unescaped);

                if (joinUrl.Contains("?"))
                {
                    string noQuery = joinUrl.Split('?')[0];
                    if (!candidateUrls.Contains(noQuery)) candidateUrls.Add(noQuery);

                    string unescapedNoQuery = Uri.UnescapeDataString(noQuery);
                    if (!candidateUrls.Contains(unescapedNoQuery)) candidateUrls.Add(unescapedNoQuery);
                }

                // 4.5. Query Graph API by joinMeetingIdSettings/joinMeetingId for numeric IDs (e.g. 413353030648656)
                if (!string.IsNullOrWhiteSpace(extractedNumericId))
                {
                    var endpointsToTry = new List<string>
                    {
                        $"https://graph.microsoft.com/v1.0/users/{userEmail}/onlineMeetings?$filter=joinMeetingIdSettings/joinMeetingId eq '{extractedNumericId}'",
                        $"https://graph.microsoft.com/v1.0/communications/onlineMeetings?$filter=joinMeetingIdSettings/joinMeetingId eq '{extractedNumericId}'"
                    };

                    foreach (var endpoint in endpointsToTry)
                    {
                        try
                        {
                            var req = new HttpRequestMessage(HttpMethod.Get, endpoint);
                            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                            var resp = await _httpClient.SendAsync(req);
                            if (resp.IsSuccessStatusCode)
                            {
                                var json = await resp.Content.ReadAsStringAsync();
                                using var doc = JsonDocument.Parse(json);
                                if (doc.RootElement.TryGetProperty("value", out var valArr) && valArr.GetArrayLength() > 0)
                                {
                                    var first = valArr[0];
                                    if (first.TryGetProperty("id", out var idProp))
                                    {
                                        string? foundId = idProp.GetString();
                                        string? subj = first.TryGetProperty("subject", out var sP) ? sP.GetString() : null;
                                        if (!string.IsNullOrWhiteSpace(foundId))
                                        {
                                            return new ResolvedMeetingInfo
                                            {
                                                MeetingId = foundId,
                                                EffectiveUserEmail = userEmail,
                                                MeetingSubject = subj ?? "Teams Meeting"
                                            };
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception exId)
                        {
                            Console.WriteLine($"[TranscriptController] Meeting ID query warning: {exId.Message}");
                        }
                    }
                }

                // 5. Direct lookup in Graph API onlineMeetings endpoint for userEmail
                foreach (var candidate in candidateUrls)
                {
                    if (candidate.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    {
                        string safeODataString = candidate.Replace("'", "''");
                        string graphUrl = $"https://graph.microsoft.com/v1.0/users/{userEmail}/onlineMeetings?$filter=joinWebUrl eq '{safeODataString}'";

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
                                    string? foundId = idProp.GetString();
                                    if (!string.IsNullOrWhiteSpace(foundId))
                                    {
                                        return new ResolvedMeetingInfo { MeetingId = foundId, EffectiveUserEmail = userEmail };
                                    }
                                }
                            }
                        }
                    }
                }

                // 6. Calendar View Lookup across target user and organizer OID
                string? organizerOid = null;
                var oidMatch = Regex.Match(joinUrl, @"Oid%22%3a%22([a-f0-9\-]+)%22", RegexOptions.IgnoreCase);
                if (!oidMatch.Success)
                {
                    oidMatch = Regex.Match(joinUrl, @"""Oid"":""([a-f0-9\-]+)""", RegexOptions.IgnoreCase);
                }
                if (oidMatch.Success)
                {
                    organizerOid = oidMatch.Groups[1].Value;
                }

                var usersToQuery = new List<string> { userEmail };
                if (!string.IsNullOrWhiteSpace(organizerOid) && !usersToQuery.Contains(organizerOid))
                {
                    usersToQuery.Add(organizerOid);
                }

                var now = DateTime.UtcNow;
                var startDateTime = now.AddDays(-60).ToString("yyyy-MM-ddTHH:mm:ssZ");
                var endDateTime = now.AddDays(14).ToString("yyyy-MM-ddTHH:mm:ssZ");

                foreach (var targetUser in usersToQuery)
                {
                    string calUrl = $"https://graph.microsoft.com/v1.0/users/{targetUser}/calendarView?startDateTime={startDateTime}&endDateTime={endDateTime}&$orderby=start/dateTime desc&$top=100";

                    var calReq = new HttpRequestMessage(HttpMethod.Get, calUrl);
                    calReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    calReq.Headers.Add("Prefer", "outlook.timezone=\"UTC\"");

                    var calResp = await _httpClient.SendAsync(calReq);
                    if (calResp.IsSuccessStatusCode)
                    {
                        var json = await calResp.Content.ReadAsStringAsync();
                        using var doc = JsonDocument.Parse(json);
                        if (doc.RootElement.TryGetProperty("value", out var eventsArr))
                        {
                            foreach (var evt in eventsArr.EnumerateArray())
                            {
                                string? omId = null;
                                string? omJoin = null;
                                string? organizerEmail = null;

                                if (evt.TryGetProperty("organizer", out var orgObj) && orgObj.TryGetProperty("emailAddress", out var eaObj))
                                {
                                    organizerEmail = eaObj.TryGetProperty("address", out var addrProp) ? addrProp.GetString() : null;
                                }

                                if (evt.TryGetProperty("onlineMeeting", out var omObj))
                                {
                                    omId = omObj.TryGetProperty("id", out var idP) ? idP.GetString() : null;
                                    omJoin = omObj.TryGetProperty("joinUrl", out var jP) ? jP.GetString() : null;
                                }

                                string bodyPreview = evt.TryGetProperty("bodyPreview", out var bp) ? bp.GetString() ?? "" : "";
                                string bodyContent = "";
                                if (evt.TryGetProperty("body", out var bodyObj) && bodyObj.TryGetProperty("content", out var bc))
                                {
                                    bodyContent = bc.GetString() ?? "";
                                }
                                string subject = evt.TryGetProperty("subject", out var subj) ? subj.GetString() ?? "" : "";

                                string digitsOnlyText = Regex.Replace($"{subject} {bodyPreview} {bodyContent} {omJoin}", @"\D", "");

                                bool isMatch = false;
                                if (!string.IsNullOrWhiteSpace(extractedNumericId) && digitsOnlyText.Contains(extractedNumericId))
                                {
                                    isMatch = true;
                                }
                                else if (!string.IsNullOrWhiteSpace(omJoin) && candidateUrls.Exists(c => omJoin.Equals(c, StringComparison.OrdinalIgnoreCase) || omJoin.Contains(c, StringComparison.OrdinalIgnoreCase) || c.Contains(omJoin, StringComparison.OrdinalIgnoreCase)))
                                {
                                    isMatch = true;
                                }

                                if (isMatch)
                                {
                                    string effectiveEmail = !string.IsNullOrWhiteSpace(organizerEmail) ? organizerEmail : targetUser;
                                    string? finalId = omId;

                                    if (string.IsNullOrWhiteSpace(finalId) && !string.IsNullOrWhiteSpace(omJoin))
                                    {
                                        finalId = await LookupOnlineMeetingIdByJoinUrlAsync(accessToken, effectiveEmail, omJoin);
                                    }

                                    if (!string.IsNullOrWhiteSpace(finalId))
                                    {
                                        return new ResolvedMeetingInfo
                                        {
                                            MeetingId = finalId,
                                            EffectiveUserEmail = effectiveEmail,
                                            MeetingSubject = subject
                                        };
                                    }
                                }
                            }
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

        private async Task<string?> GetUserGuidAsync(string accessToken, string userEmailOrGuid)
        {
            if (string.IsNullOrWhiteSpace(userEmailOrGuid)) return null;

            userEmailOrGuid = userEmailOrGuid.Trim();
            if (Guid.TryParse(userEmailOrGuid, out _))
            {
                return userEmailOrGuid;
            }

            try
            {
                // 1. Direct UPN lookup
                string graphUrl = $"https://graph.microsoft.com/v1.0/users/{userEmailOrGuid}?$select=id";
                var request = new HttpRequestMessage(HttpMethod.Get, graphUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                var response = await _httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("id", out var idProp))
                    {
                        return idProp.GetString();
                    }
                }

                // 2. OData filter lookup by mail or userPrincipalName
                string safeEmail = userEmailOrGuid.Replace("'", "''");
                string filterUrl = $"https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '{safeEmail}' or mail eq '{safeEmail}'&$select=id";
                var filterReq = new HttpRequestMessage(HttpMethod.Get, filterUrl);
                filterReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                var filterResp = await _httpClient.SendAsync(filterReq);
                if (filterResp.IsSuccessStatusCode)
                {
                    var json = await filterResp.Content.ReadAsStringAsync();
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
                Console.WriteLine($"[TranscriptController] GetUserGuid error: {ex.Message}");
            }

            return null;
        }

        private async Task<string?> LookupOnlineMeetingIdByJoinUrlAsync(string accessToken, string userEmail, string joinUrl)
        {
            try
            {
                string? userGuid = await GetUserGuidAsync(accessToken, userEmail);

                var urlsToTry = new List<string> { joinUrl };
                string unescaped = Uri.UnescapeDataString(joinUrl);
                if (!urlsToTry.Contains(unescaped)) urlsToTry.Add(unescaped);

                if (joinUrl.Contains("%"))
                {
                    string doubleUnescaped = Uri.UnescapeDataString(unescaped);
                    if (!urlsToTry.Contains(doubleUnescaped)) urlsToTry.Add(doubleUnescaped);
                }

                foreach (var url in urlsToTry)
                {
                    string safeODataString = url.Replace("'", "''");
                    var targetUrls = new List<string>();
                    if (!string.IsNullOrWhiteSpace(userGuid))
                    {
                        targetUrls.Add($"https://graph.microsoft.com/v1.0/users/{userGuid}/onlineMeetings?$filter=joinWebUrl eq '{safeODataString}'");
                    }
                    targetUrls.Add($"https://graph.microsoft.com/v1.0/communications/onlineMeetings?$filter=joinWebUrl eq '{safeODataString}'");

                    foreach (var graphUrl in targetUrls)
                    {
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
                                    string? foundId = idProp.GetString();
                                    if (!string.IsNullOrWhiteSpace(foundId)) return foundId;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TranscriptController] LookupOnlineMeetingIdByJoinUrl error: {ex.Message}");
            }

            return null;
        }

        private async Task<string?> GetLatestTranscriptIdAsync(string accessToken, string userEmail, string meetingId)
        {
            try
            {
                string? userGuid = await GetUserGuidAsync(accessToken, userEmail);
                var targetUrls = new List<string>();
                if (!string.IsNullOrWhiteSpace(userGuid))
                {
                    targetUrls.Add($"https://graph.microsoft.com/v1.0/users/{userGuid}/onlineMeetings/{meetingId}/transcripts");
                }
                targetUrls.Add($"https://graph.microsoft.com/v1.0/communications/onlineMeetings/{meetingId}/transcripts");

                foreach (var graphUrl in targetUrls)
                {
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
                string? userGuid = await GetUserGuidAsync(accessToken, userEmail);
                var targetUrls = new List<string>();
                if (!string.IsNullOrWhiteSpace(userGuid))
                {
                    targetUrls.Add($"https://graph.microsoft.com/v1.0/users/{userGuid}/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content?$format=text/vtt");
                }
                targetUrls.Add($"https://graph.microsoft.com/v1.0/communications/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content?$format=text/vtt");

                foreach (var graphUrl in targetUrls)
                {
                    var request = new HttpRequestMessage(HttpMethod.Get, graphUrl);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var response = await _httpClient.SendAsync(request);
                    if (response.IsSuccessStatusCode)
                    {
                        return await response.Content.ReadAsStringAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TranscriptController] DownloadTranscriptVtt error: {ex.Message}");
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
