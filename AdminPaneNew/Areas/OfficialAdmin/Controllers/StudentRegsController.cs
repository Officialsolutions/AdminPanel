using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using AdminPaneNew.Areas.OfficialAdmin.Models;
using LinqToExcel;
using System.Data.OleDb;
using System.Data.Entity.Validation;
using onlineportal.Areas.AdminPanel.Models;

namespace AdminPaneNew.Areas.OfficialAdmin.Controllers
{
    public class StudentRegsController : Controller
    {
        private dbcontext db = new dbcontext();

        // GET: OfficialAdmin/StudentRegs
        public ActionResult Index()
        {
            return View(db.StudentRegs.ToList());
        }
        public ActionResult updateExcel()
        {
            return View();
        }
        [HttpPost]
        public JsonResult updateExcel(StudentReg studentreg, HttpPostedFileBase FileUpload)
        {

            List<string> data = new List<string>();
            if (FileUpload != null)
            {
                // tdata.ExecuteCommand("truncate table OtherCompanyAssets");  
                if (FileUpload.ContentType == "application/vnd.ms-excel" || FileUpload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    string filename = FileUpload.FileName;
                    string targetpath = Server.MapPath("/DetailFormatInExcel/");
                    FileUpload.SaveAs(targetpath + filename);
                    string pathToExcelFile = targetpath + filename;
                    var connectionString = "";
                    if (filename.EndsWith(".xls"))
                    {
                        connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", pathToExcelFile);
                    }
                    else if (filename.EndsWith(".xlsx"))
                    {
                        connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", pathToExcelFile);
                    }

                    var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
                    var ds = new DataSet();

                    adapter.Fill(ds, "ExcelTable");

                    DataTable dtable = ds.Tables["ExcelTable"];

                    string sheetName = "Sheet1";

                    var excelFile = new ExcelQueryFactory(pathToExcelFile);
                    var artistAlbums = from a in excelFile.Worksheet<StudentReg>(sheetName) select a;

                    foreach (var a in artistAlbums)
                    {
                        try
                        {
                            if (a.StudentName != "" && a.Address != "" && a.Contact != "")
                            {
                                string rollno = null;
                                StudentReg TU = new StudentReg();
                                // string reg = Regno(a.RollNo);
                                //TU.RollNo = Regno(rollno);
                                TU.RollNo = a.RollNo;
                                TU.StudentName = a.StudentName;
                                TU.FatherName = a.FatherName;
                                TU.Address = a.Address;
                                TU.Contact = a.Contact;
                                TU.Laststudy = a.Laststudy;
                                TU.Medical = a.Medical;
                                TU.Refusal = a.Refusal;
                                TU.Email = a.Email;
                                TU.Password = a.Password;
                                db.StudentRegs.Add(TU);

                                db.SaveChanges();
                                //    a.RollNo, a.StudentName, a.FatherName, a.Address, a.Contact, a.Laststudy, a.Medical, a.Refusal, a.Email, a.Password


                            }
                            else
                            {
                                data.Add("<ul>");
                                if (a.StudentName == "" || a.StudentName == null) data.Add("<li> name is required</li>");
                                if (a.FatherName == "" || a.FatherName == null) data.Add("<li> Father Name is required</li>");
                                if (a.Address == "" || a.Address == null) data.Add("<li> Address is required</li>");
                                if (a.Contact == "" || a.Contact == null) data.Add("<li>ContactNo is required</li>");

                                data.Add("</ul>");
                                data.ToArray();
                                return Json(data, JsonRequestBehavior.AllowGet);
                            }
                        }

                        catch (DbEntityValidationException ex)
                        {
                            foreach (var entityValidationErrors in ex.EntityValidationErrors)
                            {

                                foreach (var validationError in entityValidationErrors.ValidationErrors)
                                {

                                    Response.Write("Property: " + validationError.PropertyName + " Error: " + validationError.ErrorMessage);

                                }

                            }
                        }
                    }
                    //deleting excel file from folder  
                    if ((System.IO.File.Exists(pathToExcelFile)))
                    {
                        System.IO.File.Delete(pathToExcelFile);
                    }
                    return Json("success", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //alert message for invalid file format  
                    data.Add("<ul>");
                    data.Add("<li>Only Excel file format is allowed</li>");
                    data.Add("</ul>");
                    data.ToArray();
                    return Json(data, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                data.Add("<ul>");
                if (FileUpload == null) data.Add("<li>Please choose Excel file</li>");
                data.Add("</ul>");
                data.ToArray();
                return Json(data, JsonRequestBehavior.AllowGet);
            }
        }
        //private string Regno(string rollno)
        //{
        //    //dbcontext db = new dbcontext();
        //    StudentReg stu = new StudentReg();
        //    stu = db.StudentRegs.Where(x => x.Studentid == stu.Studentid).Max();
        //    string rol = stu.RollNo;
        //    //if (roln != null)
        //    //{
        //    //    var rol = roln.RollNo;
        //        if (rol != null)
        //        {
        //        //  string[] roll = rol.Split('-');

        //        //rollno = (Convert.ToInt32(roll[1]) + Convert.ToInt32(1)).ToString();
        //        rollno = "Jan19-001";
        //    }
        //        else
        //        {
        //            rollno = "Jan19-001";
        //        }
        //    //}
        //    return rollno;
        //}
        //[HttpPost]
        //public ActionResult updateExcel(StudentReg student, HttpPostedFileBase FileUpload)
        //{

        //    //  EmployeeDBEntities objEntity = new EmployeeDBEntities();
        //    string data = "";
        //    if (FileUpload != null)
        //    {
        //        if (FileUpload.ContentType == "application/vnd.ms-excel" || FileUpload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        //        {
        //            string filename = FileUpload.FileName;

        //            if (filename.EndsWith(".xlsx"))
        //            {
        //                string targetpath = Server.MapPath("~/DetailFormatInExcel/");
        //                FileUpload.SaveAs(targetpath + filename);
        //                string pathToExcelFile = targetpath + filename;

        //                string sheetName = "Sheet1";

        //                var excelFile = new ExcelQueryFactory(pathToExcelFile);
        //                var empDetails = from a in excelFile.Worksheet<StudentReg>(sheetName) select a;
        //                foreach (var a in empDetails)
        //                {
        //                    if (a.RollNo != null)
        //                    {

        //                        //DateTime? myBirthdate = null;


        //                        if (a.Contact.Length > 10)
        //                        {
        //                            data = "Phone number should be 10 to 12 disit";
        //                            ViewBag.Message = data;

        //                        }

        //                        //  myBirthdate = Convert.ToDateTime(a.DateOfBirth);


        //                        int resullt = PostExcelData(a.RollNo, a.StudentName, a.FatherName, a.Address, a.Contact, a.Laststudy, a.Medical, a.Refusal, a.Email, a.Password);
        //                        if (resullt <= 0)
        //                        {
        //                            data = "Hello User, Found some duplicate values! Only unique employee number has inserted and duplicate values(s) are not inserted";
        //                            ViewBag.Message = data;
        //                            continue;

        //                        }
        //                        else
        //                        {
        //                            data = "Successful upload records";
        //                            ViewBag.Message = data;
        //                        }
        //                    }

        //                    else
        //                    {
        //                        data = a.RollNo + "Some fields are null, Please check your excel sheet";
        //                        ViewBag.Message = data;
        //                        return View("updateExcel");
        //                    }

        //                }
        //            }

        //            else
        //            {
        //                data = "This file is not valid format";
        //                ViewBag.Message = data;
        //            }
        //            return View("updateExcel");
        //        }
        //        else
        //        {

        //            data = "Only Excel file format is allowed";

        //            ViewBag.Message = data;
        //            return View("updateExcel");

        //        }

        //    }
        //    else
        //    {

        //        if (FileUpload == null)
        //        {
        //            data = "Please choose Excel file";
        //        }

        //        ViewBag.Message = data;
        //        return View("ExcelUpload");
        //    }
        //}

        //public int PostExcelData(string RollNo, string StudentName, string FatherName, string address, string Contact, string Laststudy, string Medical, string Refusal, string Email, string Password)
        //{
        //    dbcontext DbEntity = new dbcontext();
        //    StudentReg reg = new StudentReg();
        //    // EmployeeDBEntities DbEntity = new EmployeeDBEntities();
        //    var InsertExcelData = reg(RollNo, StudentName, FatherName, address, Contact, Laststudy, Medical, Refusal, Email, Password);

        //    return InsertExcelData;
        //}
        // GET: OfficialAdmin/StudentRegs/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            StudentReg studentReg = db.StudentRegs.Find(id);
            if (studentReg == null)
            {
                return HttpNotFound();
            }
            return View(studentReg);
        }

        // GET: OfficialAdmin/StudentRegs/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/StudentRegs/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Studentid,StudentName,FatherName,Address,Contact,Laststudy,Medical,Refusal,Email,Password,RollNo")] StudentReg studentReg, Helper Help)
        {
            if (ModelState.IsValid)
            {
                string rollno = null;
                //string reg = Regno(rollno);
                studentReg.RollNo = Help.Regno(rollno);
                db.StudentRegs.Add(studentReg);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(studentReg);
        }

        // GET: OfficialAdmin/StudentRegs/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            StudentReg studentReg = db.StudentRegs.Find(id);
            if (studentReg == null)
            {
                return HttpNotFound();
            }
            return View(studentReg);
        }

        // POST: OfficialAdmin/StudentRegs/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Studentid,StudentName,FatherName,Address,Contact,Laststudy,Medical,Refusal,Email,Password,RollNo")] StudentReg studentReg)
        {
            if (ModelState.IsValid)
            {
                db.Entry(studentReg).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(studentReg);
        }

        // GET: OfficialAdmin/StudentRegs/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            StudentReg studentReg = db.StudentRegs.Find(id);
            if (studentReg == null)
            {
                return HttpNotFound();
            }
            return View(studentReg);
        }

        // POST: OfficialAdmin/StudentRegs/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            StudentReg studentReg = db.StudentRegs.Find(id);
            db.StudentRegs.Remove(studentReg);
            db.SaveChanges();
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
