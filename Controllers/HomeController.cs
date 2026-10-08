using HorizonChat.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc; 
using Microsoft.Exchange.WebServices.Data;
using SummarizeMail.Service;
using System.DirectoryServices.AccountManagement;
using System.Text;

namespace SummarizeMail.Controllers
{
    [ApiController]
    public class HomeController : ControllerBase
    {
        [HttpPost("login")]
        public IActionResult Login([FromBody] LoginRequestDto request)
        {
            if (request == null)
                return BadRequest(new { message = "Request body is required" });

            if (string.IsNullOrWhiteSpace(request.Email) ||
                string.IsNullOrWhiteSpace(request.Password))
                return BadRequest(new { message = "Email and password are required" });

            try
            {
                using var pc = new PrincipalContext(ContextType.Domain, "hfmg.net");

                var isValid = pc.ValidateCredentials(request.Email, request.Password);

                if (!isValid)
                    return Unauthorized(new { message = "Invalid credentials" });

                var user = UserPrincipal.FindByIdentity(pc, request.Email);

                if (user == null)
                    return NotFound(new { message = "User not found in AD" });

                var response = new LoginResponseDto
                {
                    Account = user.EmailAddress,
                    Name = user.Name,
                    Mail = user.EmailAddress,
                    Role = user.Description
                };
                
                return Ok(response);
            }
            catch (PrincipalServerDownException)
            {
                return StatusCode(503, new { message = "Domain controller unavailable" });
            }
            catch (Exception)
            {
                // TODO: log exception
                return StatusCode(500, new { message = "Internal server error" });
            }
        } 
        [HttpGet("get")]
        public async Task<ActionResult> Get()
        {
            using var pc = new PrincipalContext(ContextType.Domain, "hfmg.net");
            List<GroupMail> groupMails = new List<GroupMail>();
            using (PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(pc)))
            {
                foreach (Principal result in searcher.FindAll())
                {
                    if (result is UserPrincipal user1)
                    {
                        if (!string.IsNullOrEmpty(user1.EmailAddress))
                        {
                            groupMails.Add(new GroupMail
                            {
                                MailId = user1.EmailAddress,
                                GroupName = user1.StructuralObjectClass
                            });
                        }
                    }
                }


            }
            return Ok(groupMails);
        }
        
        public class LoginResponseDto
        {
            public string Account { get; set; }
            public string Name { get; set; }
            public string Mail { get; set; }
            public string Role { get; set; }
        }

        public class LoginRequestDto
        {
            public string Email { get; set; }
            public string Password { get; set; }
        }

        public partial class GroupMail
        { 

            public string MailId { get; set; } = null!;

            public string GroupName { get; set; } = null!;
        }

    }
}