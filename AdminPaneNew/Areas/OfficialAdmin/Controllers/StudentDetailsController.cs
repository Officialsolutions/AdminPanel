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
    public class StudentDetailsController : Controller
    {
        private dbcontext db = new dbcontext();
        public static string img;
        DateTime date = new DateTime();
        // GET: OfficialAdmin/StudentDetails
        public ActionResult Index()
        {
            return View(db.StudentDetails.ToList());
        }

        // GET: OfficialAdmin/StudentDetails/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            StudentDetail studentDetail = db.StudentDetails.Find(id);
            if (studentDetail == null)
            {
                return HttpNotFound();
            }
            return View(studentDetail);
        }

        // GET: OfficialAdmin/StudentDetails/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/StudentDetails/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Studentid,Name,FatherName,Mobile,Email,Gender,DOB,Image,Address,Qualification,CourseType,RollNo,JoiningDate,UserName,Password,date,Status")] StudentDetail studentDetail,HttpPostedFileBase file,Helper Help)
        {
            if (ModelState.IsValid)
            {

                studentDetail.Image = Help.uploadfile(file);
                studentDetail.date = System.DateTime.Now;
                db.StudentDetails.Add(studentDetail);
                db.SaveChanges();
                TempData["Success"] = "Saved Successfully";
                return RedirectToAction("Index");
            }

            return View(studentDetail);
        }

        // GET: OfficialAdmin/StudentDetails/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            StudentDetail studentDetail = db.StudentDetails.Find(id);
            date = studentDetail.date;
            img = studentDetail.Image;
            if (studentDetail == null)
            {
                return HttpNotFound();
            }
            return View(studentDetail);
        }

        // POST: OfficialAdmin/StudentDetails/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Studentid,Name,FatherName,Mobile,Email,Gender,DOB,Image,Address,Qualification,CourseType,RollNo,JoiningDate,UserName,Password,date,Status")] StudentDetail studentDetail, HttpPostedFileBase file, Helper Help)
        {
            if (ModelState.IsValid)
            {
                studentDetail.date = System.DateTime.Now;
                studentDetail.Image = file != null ? Help.uploadfile(file) : img;
                #region delete file
                string fullPath = Request.MapPath("~/UploadedFiles/" + img);
                if (img == studentDetail.Image)
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
                db.Entry(studentDetail).State = EntityState.Modified;
                db.SaveChanges();
                TempData["Success"] = "Updated Successfully";
                return RedirectToAction("Index");
            }
            return View(studentDetail);
        }

        // GET: OfficialAdmin/StudentDetails/Delete/5
        [Authorize(Roles = "Admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            StudentDetail studentDetail = db.StudentDetails.Find(id);
            if (studentDetail == null)
            {
                return HttpNotFound();
            }
            return View(studentDetail);
        }

        // POST: OfficialAdmin/StudentDetails/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            StudentDetail studentDetail = db.StudentDetails.Find(id);
            db.StudentDetails.Remove(studentDetail);
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
        public ActionResult AssignTest(int id)
        {
            return View();
        }
    }
}
