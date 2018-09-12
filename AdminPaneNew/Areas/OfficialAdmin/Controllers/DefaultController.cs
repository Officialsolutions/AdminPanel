using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace AdminPaneNew.Areas.OfficialAdmin.Controllers
{
    public class DefaultController : Controller
    {
        // GET: OfficialAdmin/Default
        public ActionResult Index()
        {
            return View();
        }
      
    }
}