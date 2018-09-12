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

namespace AdminPaneNew.Areas.OfficialAdmin.Controllers
{
    public class ServicesController : Controller
    {
        private dbcontext db = new dbcontext();
        public static string img, thumb;
        // GET: OfficialAdmin/Services
        public ActionResult Index()
        {
            return View(db.Services.ToList());
        }

        // GET: OfficialAdmin/Services/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Service service = db.Services.Find(id);
            if (service == null)
            {
                return HttpNotFound();
            }
            return View(service);
        }

        // GET: OfficialAdmin/Services/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/Services/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Serviceid,Name,Description,ShortDescription,Image,date,Thumbnail,Keyword,MetaDescription")] Service service, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                service.date = System.DateTime.Now;
                service.Image = Help.uploadfile(file);
                service.Thumbnail = Help.uploadfile(file);
                db.Services.Add(service); 
                db.SaveChanges();
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(service);
        }

        // GET: OfficialAdmin/Services/Edit/5
        [Authorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Service service = db.Services.Find(id);
            img = service.Image;
            thumb = service.Thumbnail;
            if (service == null)
            {
                return HttpNotFound();
            }
            return View(service);
        }

        // POST: OfficialAdmin/Services/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Serviceid,Name,Description,ShortDescription,Image,date,Thumbnail,Keyword,MetaDescription")] Service service, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                service.date = System.DateTime.Now;
                service.Image = file != null ? Help.uploadfile(file) : img;
                #region delete file
                string fullPath = Request.MapPath("~/UploadedFiles/" + img);
                if (img == service.Image)
                {
                }
                else
                {
                    if (System.IO.File.Exists(fullPath))
                    {
                        System.IO.File.Delete(fullPath);
                    }
                }
                #endregion
                service.Thumbnail = file != null ? Help.uploadfile(file) : thumb;
                #region delete file
                string fullPath2 = Request.MapPath("~/UploadedFiles/" + thumb);
                if (thumb == service.Thumbnail)
                {
                }
                else
                {
                    if (System.IO.File.Exists(fullPath2))
                    {
                        System.IO.File.Delete(fullPath2);
                    }
                }
                #endregion
                db.Entry(service).State = EntityState.Modified;
                db.SaveChanges();
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(service);
        }

        // GET: OfficialAdmin/Services/Delete/5
        [Authorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Service service = db.Services.Find(id);
            if (service == null)
            {
                return HttpNotFound();
            }
            return View(service);
        }

        // POST: OfficialAdmin/Services/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Service service = db.Services.Find(id);
            db.Services.Remove(service);
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
    }
}
