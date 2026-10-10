using DISRIS_UI.Modal;

namespace DISRIS_UI.Inferface
{
    public interface IAuthService
    {
        Task<bool> loginUser(LoginRequest loginRequest);
    }
}
