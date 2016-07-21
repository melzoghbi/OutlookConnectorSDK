using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(webapp45.Startup))]
namespace webapp45
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
