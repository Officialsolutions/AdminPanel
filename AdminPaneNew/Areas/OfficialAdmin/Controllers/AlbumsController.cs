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
    public class AlbumsController : Controller
    {
     
        private dbcontext db = new dbcontext();
        public static string img;
        public static string imags;
       
        int aid;
        // GET: OfficialAdmin/Albums
        public ActionResult Index()
        {
            return View(db.Albums.ToList());
        }

        // GET: OfficialAdmin/Albums/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Album album = db.Albums.Find(id);
            if (album == null)
            {
                return HttpNotFound();
            }
            return View(album);
        }

        // GET: OfficialAdmin/Albums/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/Albums/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Albumid,AlbumName,Image,date")] Album album, HttpPostedFileBase file, Gallery gallery, IEnumerable<HttpPostedFileBase> file2, Helper Help)
        {
            if (ModelState.IsValid)
            {
                album.date = System.DateTime.Now;
                album.Image = Help.uploadfile(file);
                db.Albums.Add(album);
                db.SaveChanges();

                gallery.date = System.DateTime.Now;
            
                gallery.Albumid = db.Albums.Max(x => x.Albumid);
                foreach (var a in file2)
                {
                    gallery.Images = Help.uploadfile(a);

                    db.Galleries.Add(gallery);
                    
                    db.SaveChanges();
                }

                //gallery.Images = Help.uploadfile(file2);
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(album);
        }

        // GET: OfficialAdmin/Albums/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DataTable dt = new DataTable();
            dt.Columns.Add("id", System.Type.GetType("System.Int32"));
            dt.Columns.Add("Image");
            Album album = db.Albums.Find(id);
            aid = album.Albumid;
            img = album.Image;
            imags = db.Galleries.FirstOrDefault(x => x.Albumid == album.Albumid).Images;
            //MyDt = db.Galleries.Where(x => x.Albumid == id).ToListAsync();
            //if (gallery != null)
            //{
            //    DataRow dr = new DataRow();
            //    dr["Image"]
            //    imags = gallery.Images;
            //}
            // db.Galleries.Find(id);

            if (album == null)
            {
                return HttpNotFound();
            }
            return View(album);
        }

        // POST: OfficialAdmin/Albums/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Albumid,AlbumName,Image,date")] Album album,Gallery gallery, HttpPostedFileBase file, IEnumerable<HttpPostedFileBase> file2, Helper Help)
        {
            if (ModelState.IsValid)
            {
                album.date = System.DateTime.Now;
                album.Image = file != null ? Help.uploadfile(file) : img;
                #region delete file
                string fullPath = Request.MapPath("~/UploadedFiles/" + img);
                if (img == album.Image)
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
                db.Entry(album).State = EntityState.Modified;
                db.SaveChanges();

                ////////////////////////////////

                gallery.date = System.DateTime.Now;
                gallery.Albumid = aid;
                Session["val"] = db.Galleries.Where(x => x.Albumid == album.Albumid).ToList();
                string ss = Session["val"].ToString();
                
                if (file2 != null)
                {
                    foreach (var a in file2)
                    {
                        //gallery.Images = Help.uploadfile(a);
                        gallery.Images = file2 != null ? Help.uploadfile(a) : imags;
                        db.Galleries.Add(gallery);

                        db.SaveChanges();

                    }
                }
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(album);
        }

        // GET: OfficialAdmin/Albums/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Album album = db.Albums.Find(id);
            if (album == null)
            {
                return HttpNotFound();
            }
            return View(album);
        }

        // POST: OfficialAdmin/Albums/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Album album = db.Albums.Find(id);
            db.Albums.Remove(album);
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
