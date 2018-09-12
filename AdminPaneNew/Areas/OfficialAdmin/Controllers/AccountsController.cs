using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using AdminPaneNew.Areas.OfficialAdmin.Models;
using onlineportal.Areas.AdminPanel.Models;
using System.Security.Cryptography;
using System.Web.Security;

namespace AdminPaneNew.Areas.OfficialAdmin.Controllers
{
    public class AccountsController : Controller
    {
        private dbcontext db = new dbcontext();
      
        // GET: OfficialAdmin/Accounts
        
        public ActionResult Index()
        {
            return View(db.Accounts.ToList());
        }

        // GET: OfficialAdmin/Accounts/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Account account = db.Accounts.Find(id);
            if (account == null)
            {
                return HttpNotFound();
            }
            return View(account);
        }

        // GET: OfficialAdmin/Accounts/Create
       
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/Accounts/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Accountid,Name,Mobile,Usename,Password,date,Role")] Account account, Helper Help)
        {
            if (ModelState.IsValid)
            {
                
                Account acc = new Account();
                account.date = System.DateTime.Now;
                db.Accounts.Add(account);
                db.SaveChanges();
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(account);
        }

        // GET: OfficialAdmin/Accounts/Edit/5
        [Authorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Account account = db.Accounts.Find(id);
            if (account == null)
            {
                return HttpNotFound();
            }
            return View(account);
        }

        // POST: OfficialAdmin/Accounts/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Accountid,Name,Mobile,Usename,Password,date,Role")] Account account, Helper Help)
        {
            if (ModelState.IsValid)
            {
                account.date = System.DateTime.Now;
                db.Entry(account).State = EntityState.Modified;
                db.SaveChanges();
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(account);
        }

        // GET: OfficialAdmin/Accounts/Delete/5
        [Authorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Account account = db.Accounts.Find(id);
            if (account == null)
            {
                return HttpNotFound();
            }
            return View(account);
        }

        // POST: OfficialAdmin/Accounts/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Account account = db.Accounts.Find(id);
            db.Accounts.Remove(account);
            db.SaveChanges();
            TempData["Success"] = "Deleted Successfully";
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
        [HttpGet]
        public ActionResult Login()
        {
            return View();
        }
        //[HttpPost]
        //public ActionResult Login([Bind(Include = "id,Name,Mobile,Username,Password")] Account account)
        //{
        //    Account acc = db.Accounts.Where(x => x.Usename == account.Usename && x.Password == account.Password).FirstOrDefault();
        //    if (acc != null)
        //    {
        //        return RedirectToAction("Index","Default");
        //    }
        //    else
        //    {
        //        return View();
        //    }

        //}
        [HttpPost]
        public ActionResult Login(Account model, string returnUrl)
        {
            dbcontext db = new dbcontext();
            var dataItem = db.Accounts.Where(x => x.Usename == model.Usename && x.Password == model.Password).First();
            if (dataItem != null)
            {
                FormsAuthentication.SetAuthCookie(dataItem.Usename, false);
                if (Url.IsLocalUrl(returnUrl) && returnUrl.Length > 1 && returnUrl.StartsWith("/")
                         && !returnUrl.StartsWith("//") && !returnUrl.StartsWith("/\\"))
                {
                    return Redirect(returnUrl);
                }
                else
                {
                    TempData["Success"] = "Login Successfully";
                 
                    return RedirectToAction("Index","Default");
                    
                }
            }
            else
            {
                ModelState.AddModelError("", "Invalid user/pass");
                return View();
            }
        }
    
        //[Authorize]
        public ActionResult SignOut()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Login", "Accounts");
        }

    }
}
