using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using TeamsBot.Core.Interfaces;

namespace TeamsBot.Server.Controllers
{
    [ApiController]
    [Route("api/auth")]
    public class AuthController : ControllerBase
    {
        private readonly IAuthService _authService;

        public AuthController(IAuthService authService)
        {
            _authService = authService;
        }

        public class SendCodeRequest
        {
            public string Email { get; set; } = string.Empty;
        }

        public class VerifyCodeRequest
        {
            public string Email { get; set; } = string.Empty;
            public string Code { get; set; } = string.Empty;
        }

        [HttpPost("send-code")]
        public async Task<IActionResult> SendCode([FromBody] SendCodeRequest request)
        {
            var result = await _authService.SendVerificationCodeAsync(request.Email);
            if (!result.Success)
            {
                return BadRequest(new { success = false, message = result.Message });
            }
            return Ok(new { success = true, message = result.Message });
        }

        [HttpPost("verify-code")]
        public async Task<IActionResult> VerifyCode([FromBody] VerifyCodeRequest request)
        {
            var result = await _authService.VerifyCodeAsync(request.Email, request.Code);
            if (!result.Success)
            {
                return BadRequest(new { success = false, message = result.Message });
            }
            return Ok(new { success = true, message = result.Message });
        }
    }
}
