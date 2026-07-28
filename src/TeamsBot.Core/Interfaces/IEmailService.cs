using System.Threading.Tasks;

namespace TeamsBot.Core.Interfaces
{
    public interface IEmailService
    {
        Task<bool> SendVerificationCodeAsync(string toEmail, string code);
    }
}
