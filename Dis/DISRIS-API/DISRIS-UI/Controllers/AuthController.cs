using DISRIS_UI.Inferface;
using DISRIS_UI.Modal;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DISRIS_UI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AuthController : ControllerBase
    {


        public readonly IAuthService authService; 
        public AuthController(IAuthService _authService)
        {
            authService = _authService;
        }
        [HttpGet]
        public async Task<IActionResult> Login([FromBody] LoginRequest request)
        {
            // Implementation for login logic
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            try
            {
                var response = await authService.loginUser(request);
                return Ok(new { message = "Login successful" });
            }
            catch (Exception ex)
            {

                throw;
            }


            
        }

    }
}
