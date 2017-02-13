using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(MvcSpAppWeb.Startup))]
namespace MvcSpAppWeb
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
