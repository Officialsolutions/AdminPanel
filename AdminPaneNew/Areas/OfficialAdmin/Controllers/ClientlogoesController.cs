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
    public class ClientlogoesController : Controller
    {
        private dbcontext db = new dbcontext();
        public static string img;
        // GET: OfficialAdmin/Clientlogoes
        public ActionResult Index()
        {
            return View(db.Clientlogoes.ToList());
        }

        // GET: OfficialAdmin/Clientlogoes/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Clientlogo clientlogo = db.Clientlogoes.Find(id);
            if (clientlogo == null)
            {
                return HttpNotFound();
            }
            return View(clientlogo);
        }

        // GET: OfficialAdmin/Clientlogoes/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/Clientlogoes/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Clientid,Name,Image,date")] Clientlogo clientlogo, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                clientlogo.date = System.DateTime.Now;
                clientlogo.Image = Help.uploadfile(file);
                db.Clientlogoes.Add(clientlogo);
                db.SaveChanges();
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(clientlogo);
        }

        // GET: OfficialAdmin/Clientlogoes/Edit/5
        [Authorize(Roles = "Admin")]
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Clientlogo clientlogo = db.Clientlogoes.Find(id);
            img = clientlogo.Image;
            if (clientlogo == null)
            {
                return HttpNotFound();
            }
            return View(clientlogo);
        }

        // POST: OfficialAdmin/Clientlogoes/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Clientid,Name,Image,date")] Clientlogo clientlogo, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                clientlogo.date = System.DateTime.Now;
                clientlogo.Image = file != null ? Help.uploadfile(file) : img;
                #region delete file
                string fullPath = Request.MapPath("~/UploadedFiles/" + img);
                if (img == clientlogo.Image)
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
                db.Entry(clientlogo).State = EntityState.Modified;
                db.SaveChanges();
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(clientlogo);
        }

        // GET: OfficialAdmin/Clientlogoes/Delete/5
        [Authorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Clientlogo clientlogo = db.Clientlogoes.Find(id);
            if (clientlogo == null)
            {
                return HttpNotFound();
            }
            return View(clientlogo);
        }

        // POST: OfficialAdmin/Clientlogoes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Clientlogo clientlogo = db.Clientlogoes.Find(id);
            db.Clientlogoes.Remove(clientlogo);
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
