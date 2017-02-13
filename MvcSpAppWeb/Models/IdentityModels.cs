using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;


namespace MvcSpAppWeb.Models
{
    // Vous pouvez ajouter des données de profil pour l'utilisateur en ajoutant plus de propriétés à votre classe ApplicationUser ; consultez http://go.microsoft.com/fwlink/?LinkID=317594 pour en savoir davantage.
    public class ApplicationUser : IdentityUser
    {
        private int sharePointId;

        public int SharePointId
        {
            get { return sharePointId; }
            set { sharePointId = value; }
        }
        private string targetHost;

        public string TargetHost
        {
            get { return targetHost; }
            set { targetHost = value; }
        }
        private string targetPrincipalName;

        public string TargetPrincipalName
        {
            get { return targetPrincipalName; }
            set { targetPrincipalName = value; }
        }
        private string displayName;

        public string DisplayName
        {
            get { return displayName; }
            set { displayName = value; }
        }
        private string refreshToken;

        public string RefreshToken
        {
            get { return refreshToken; }
            set { refreshToken = value; }
        }
        private string targetRealm;

        public string TargetRealm
        {
            get { return targetRealm; }
            set { targetRealm = value; }
        }

        private string ressource;

        public string Ressource
        {
            get { return ressource; }
            set { ressource = value; }
        }

        private string securityTokenServiceUrl;

        public string SecurityTokenServiceUrl
        {
            get { return securityTokenServiceUrl; }
            set { securityTokenServiceUrl = value; }
        }

        private string accessToken;

        public string AccessToken
        {
            get { return accessToken; }
            set { accessToken = value; }
        }

        private string appOnlyAccessToken;

        public string AppOnlyAccessToken
        {
            get { return appOnlyAccessToken; }
            set { appOnlyAccessToken = value; }
        }

        public async Task<ClaimsIdentity> GenerateUserIdentityAsync(UserManager<ApplicationUser> manager)
        {
            // Notez qu'authenticationType doit correspondre à l'élément défini dans CookieAuthenticationOptions.AuthenticationType
            var userIdentity = await manager.CreateIdentityAsync(this, DefaultAuthenticationTypes.ApplicationCookie);
            // Ajouter les revendications personnalisées de l’utilisateur ici
            return userIdentity;
        }
    }

    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext()
            : base("DefaultConnection", throwIfV1Schema: false)
        {
        }

        public static ApplicationDbContext Create()
        {
            return new ApplicationDbContext();
        }
    }
}