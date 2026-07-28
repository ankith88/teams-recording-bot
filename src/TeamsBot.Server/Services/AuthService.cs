using System;
using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Threading.Tasks;
using TeamsBot.Core.Interfaces;

namespace TeamsBot.Server.Services
{
    public class AuthService : IAuthService
    {
        private readonly IEmailService _emailService;
        private readonly ConcurrentDictionary<string, OtpEntry> _otpStore = new(StringComparer.OrdinalIgnoreCase);

        private class OtpEntry
        {
            public string Code { get; set; } = string.Empty;
            public DateTime ExpiresAt { get; set; }
        }

        public AuthService(IEmailService emailService)
        {
            _emailService = emailService;
        }

        public async Task<(bool Success, string Message)> SendVerificationCodeAsync(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return (false, "Please enter your Mail Plus corporate email address.");
            }

            string trimmedEmail = email.Trim().ToLower();

            if (!trimmedEmail.EndsWith("@mailplus.com.au") && !trimmedEmail.EndsWith("@mailplus.com"))
            {
                return (false, "Access Denied: Only authorized Mail Plus corporate users (@mailplus.com.au) can access this application.");
            }

            // Generate secure 6-digit verification code using RandomNumberGenerator
            int randomNumber = RandomNumberGenerator.GetInt32(100000, 999999);
            string code = randomNumber.ToString();

            // Store code with 10 minute expiration
            _otpStore[trimmedEmail] = new OtpEntry
            {
                Code = code,
                ExpiresAt = DateTime.UtcNow.AddMinutes(10)
            };

            // Trigger email dispatch via IEmailService
            await _emailService.SendVerificationCodeAsync(trimmedEmail, code);

            return (true, $"Security passcode sent to {trimmedEmail}!");
        }

        public async Task<(bool Success, string Message)> VerifyCodeAsync(string email, string code)
        {
            if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(code))
            {
                return await Task.FromResult((false, "Email and 6-digit verification code are required."));
            }

            string trimmedEmail = email.Trim().ToLower();
            string trimmedCode = code.Trim();

            if (!_otpStore.TryGetValue(trimmedEmail, out var entry))
            {
                return await Task.FromResult((false, "No verification code was sent to this email or code has expired. Please request a new code."));
            }

            if (DateTime.UtcNow > entry.ExpiresAt)
            {
                _otpStore.TryRemove(trimmedEmail, out _);
                return await Task.FromResult((false, "Verification code has expired. Please click 'Resend Code' to receive a new passcode."));
            }

            if (entry.Code != trimmedCode)
            {
                return await Task.FromResult((false, "Invalid verification code. Please check your email inbox and try again."));
            }

            // Successfully verified - clear OTP entry to prevent reuse
            _otpStore.TryRemove(trimmedEmail, out _);
            return await Task.FromResult((true, "Passcode verified successfully. Logged in to Mail Plus Portal."));
        }
    }
}
