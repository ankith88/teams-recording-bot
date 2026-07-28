using System.Threading.Tasks;

namespace TeamsBot.Core.Interfaces
{
    public interface IAuthService
    {
        Task<(bool Success, string Message)> SendVerificationCodeAsync(string email);
        Task<(bool Success, string Message)> VerifyCodeAsync(string email, string code);
    }
}
