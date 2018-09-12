using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(AdminPaneNew.Startup))]
namespace AdminPaneNew
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
