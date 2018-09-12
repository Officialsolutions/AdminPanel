using System.Web.Mvc;

namespace AdminPaneNew.Areas.OfficialAdmin
{
    public class OfficialAdminAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "OfficialAdmin";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "OfficialAdmin_default",
                "OfficialAdmin/{controller}/{action}/{id}",
                    defaults: new { controller = "Accounts", action = "Login", id = UrlParameter.Optional }
            );
        }
    }
}