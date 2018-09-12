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
    public class slidersController : Controller
    {
        private dbcontext db = new dbcontext();
        public static string img;
        [Authorize]
        // GET: OfficialAdmin/sliders
        public ActionResult Index()
        {
            return View(db.sliders.ToList());
        }

        // GET: OfficialAdmin/sliders/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            slider slider = db.sliders.Find(id);
            if (slider == null)
            {
                return HttpNotFound();
            }
            return View(slider);
        }

        // GET: OfficialAdmin/sliders/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/sliders/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
 
        public ActionResult Create([Bind(Include = "Sliderid,Name,Description,Image,url,date")] slider slider, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {

                slider.date = System.DateTime.Now;
                slider.Image =Help.uploadfile(file);
                db.sliders.Add(slider);
                db.SaveChanges();
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(slider);
        }

        // GET: OfficialAdmin/sliders/Edit/5
        [Authorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            slider slider = db.sliders.Find(id);
            img = slider.Image;
            if (slider == null)
            {
                return HttpNotFound();
            }
            return View(slider);
        }

        // POST: OfficialAdmin/sliders/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Sliderid,Name,Description,Image,url,date")] slider slider, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                slider.date = System.DateTime.Now;
                slider.Image = file != null ? Help.uploadfile(file) : img;
                #region delete file
                string fullPath = Request.MapPath("~/UploadedFiles/" + img);
                if (img == slider.Image)
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
                db.Entry(slider).State = EntityState.Modified;
                db.SaveChanges();
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(slider);
        }

        // GET: OfficialAdmin/sliders/Delete/5
        [Authorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            slider slider = db.sliders.Find(id);
            if (slider == null)
            {
                return HttpNotFound();
            }
            return View(slider);
        }

        // POST: OfficialAdmin/sliders/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            slider slider = db.sliders.Find(id);
            db.sliders.Remove(slider);
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
