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
    public class FeaturesController : Controller
    {
        private dbcontext db = new dbcontext();
        public static string img, thumb;
        // GET: OfficialAdmin/Features
        public ActionResult Index()
        {
            return View(db.Features.ToList());
        }

        // GET: OfficialAdmin/Features/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Features features = db.Features.Find(id);
            if (features == null)
            {
                return HttpNotFound();
            }
            return View(features);
        }

        // GET: OfficialAdmin/Features/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/Features/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Featureid,Name,Description,ShortDescription,Image,date,Thumbnail,Keyword,MetaDescription")] Features features, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                features.date = System.DateTime.Now;
                features.Image = Help.uploadfile(file);
                features.Thumbnail = Help.uploadfile(file);
                db.Features.Add(features);
                db.SaveChanges();
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(features);
        }

        // GET: OfficialAdmin/Features/Edit/5
        [Authorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Features features = db.Features.Find(id);
            img = features.Image;
            thumb = features.Thumbnail;
            if (features == null)
            {
                return HttpNotFound();
            }
            return View(features);
        }

        // POST: OfficialAdmin/Features/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Featureid,Name,Description,ShortDescription,Image,date,Thumbnail,Keyword,MetaDescription")] Features features, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                features.date = System.DateTime.Now;
                features.Image = file != null ? Help.uploadfile(file) : img;
                #region delete file
                string fullPath = Request.MapPath("~/UploadedFiles/" + img);
                if (img == features.Image)
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
                features.Thumbnail = file != null ? Help.uploadfile(file) : thumb;
                #region delete file
                string fullPath2 = Request.MapPath("~/UploadedFiles/" + thumb);
                if (thumb == features.Thumbnail)
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
                db.Entry(features).State = EntityState.Modified;
                db.SaveChanges();
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(features);
        }

        // GET: OfficialAdmin/Features/Delete/5
        [Authorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Features features = db.Features.Find(id);
            if (features == null)
            {
                return HttpNotFound();
            }
            return View(features);
        }

        // POST: OfficialAdmin/Features/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Features features = db.Features.Find(id);
            db.Features.Remove(features);
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
