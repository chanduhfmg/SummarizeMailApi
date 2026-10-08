using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/auth")]
public class AuthController : ControllerBase
{
    private readonly GraphAuthService _authService;

    public AuthController(GraphAuthService authService)
    {
        _authService = authService;
    }

    [HttpGet("login")]
    public async Task<IActionResult> Login()
    {
        var loginUrl = await _authService.GetLoginUrlAsync();
        return Redirect(loginUrl.ToString());
    }

    [HttpGet("callback")]
    public async Task<IActionResult> Callback(
        [FromQuery] string code,
        [FromQuery] string error,
        [FromQuery] string error_description)
    {
        if (!string.IsNullOrEmpty(error))
            return BadRequest(new { error, error_description });

        if (string.IsNullOrEmpty(code))
            return BadRequest("No authorization code received from Microsoft.");

        var result = await _authService.CompleteLoginAsync(code);
        var profile = await _authService.GetUserProfileAsync(result.AccessToken);

        return Ok(new
        {
            message = "Logged in successfully",
            name = profile.DisplayName,
            email = profile.Email,
            role = profile.JobTitle,           // display-only, not an authorization control
            homeAccountId = result.Account.HomeAccountId.Identifier,
            expiresOn = result.ExpiresOn
        });
    }
}