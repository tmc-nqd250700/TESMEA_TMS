using TESMEA_TMS.Models.Entities;

namespace TESMEA_TMS.Configs
{
    public interface ICurrentUser
    {
        UserAccount UserAccount { get; set; }
        Role Role { get; set; }
        IEnumerable<string> Claims { get; set; }
        Guid SessionId { get; set; }
        void SetUser(UserAccount userAccount, Role role, IEnumerable<string> claims, Guid sessionId);
        void Clear();
    }

    public class CurrentUser : ICurrentUser
    {
        private static readonly CurrentUser _instance = new CurrentUser();
        public static CurrentUser Instance => _instance;

        public CurrentUser() { }

        public UserAccount UserAccount { get; set; }
        public Role Role { get; set; }
        public IEnumerable<string> Claims { get; set; }
        public Guid SessionId { get; set; }

        public void SetUser(UserAccount user, Role role, IEnumerable<string> claims, Guid sessionId)
        {
            UserAccount = user;
            Role = role;
            Claims = claims;
            SessionId = sessionId;
        }

        public void Clear()
        {
            UserAccount = null;
            Role = null;
            Claims = null;
            SessionId = Guid.Empty;
        }
    }
}
