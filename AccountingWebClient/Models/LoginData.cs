using System.ComponentModel.DataAnnotations;

namespace AccountingWebClient.Models
{
    public class LoginData
    {
        [Required]
        public string Username { get; set; }

        [Required, DataType(DataType.Password)]
        public string Password { get; set; }

        public bool RememberMe { get; set; }
    }
}