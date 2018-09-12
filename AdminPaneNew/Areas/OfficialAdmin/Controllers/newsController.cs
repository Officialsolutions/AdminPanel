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
    public class newsController : Controller
    {
        private dbcontext db = new dbcontext();
        public static string img, thumb;
        // GET: OfficialAdmin/news
        public ActionResult Index()
        {
            return View(db.news.ToList());
        }

        // GET: OfficialAdmin/news/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            news news = db.news.Find(id);
            if (news == null)
            {
                return HttpNotFound();
            }
            return View(news);
        }

        // GET: OfficialAdmin/news/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/news/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Newsid,Name,Description,ShortDescription,Image,Thumbnail,date")] news news, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                news.date = System.DateTime.Now;
                news.Image = Help.uploadfile(file);
                news.Thumbnail = Help.uploadfile(file);
                db.news.Add(news);
                db.SaveChanges();
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(news);
        }

        // GET: OfficialAdmin/news/Edit/5
        [Authorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            news news = db.news.Find(id);
            img = news.Image;
            thumb = news.Thumbnail;
            if (news == null)
            {
                return HttpNotFound();
            }
            return View(news);
        }

        // POST: OfficialAdmin/news/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Newsid,Name,Description,ShortDescription,Image,Thumbnail,date")] news news, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                news.date = System.DateTime.Now;
                news.Image = file != null ? Help.uploadfile(file) : img;
                #region delete file
                string fullPath = Request.MapPath("~/UploadedFiles/" + img);
                if (img == news.Image)
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
                news.Thumbnail = file != null ? Help.uploadfile(file) : thumb;
                db.Entry(news).State = EntityState.Modified;
                db.SaveChanges();
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(news);
        }

        // GET: OfficialAdmin/news/Delete/5
        [Authorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            news news = db.news.Find(id);
            if (news == null)
            {
                return HttpNotFound();
            }
            return View(news);
        }

        // POST: OfficialAdmin/news/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            news news = db.news.Find(id);
            db.news.Remove(news);
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
