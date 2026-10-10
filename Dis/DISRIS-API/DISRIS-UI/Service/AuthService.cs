using DISRIS_UI.Inferface;
using DISRIS_UI.Modal;
using System.Data.Odbc;

namespace DISRIS_UI.Service
{
    public class AuthService : IAuthService
    {
        private readonly string _con;
   
   
        public AuthService(IConfiguration configuration)
        {
            _con = configuration.GetConnectionString("DefaultConnection");
        }

        public async Task<bool> loginUser(LoginRequest loginRequest)
        {
            using (OdbcConnection connection = new OdbcConnection(_con))
            {
                await connection.OpenAsync();
                string query = "SELECT COUNT(*) FROM patoolslib.logpatient WHERE userid = ? AND tmhpwd= ?";
                using (OdbcCommand command = new OdbcCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@userid", loginRequest.userName);
                    command.Parameters.AddWithValue("@tmhpwd", loginRequest.password);
                    int count = Convert.ToInt32(await command.ExecuteScalarAsync());
                    return count > 0; // Return true if user exists, false otherwise
                }
            }
        }
    }
}
