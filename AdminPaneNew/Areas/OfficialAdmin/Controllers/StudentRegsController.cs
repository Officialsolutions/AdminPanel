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
        public static int regno;
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

                    foreach (DataRow a in dtable.Rows)
                    {
                        try
                        {
                            if (a["Student Name"] != "" && a["Address"] != "" && a["Contact"] != "")
                            {
                                StudentReg TU = new StudentReg();
                               StudentReg studentreg1 = db.StudentRegs.FirstOrDefault();
                                if (studentreg1 == null)
                                {
                                    TU.RollNo = 0001;
                                }
                                else
                                {
                                    var ab = db.StudentRegs.Max(x => x.RollNo);
                                    TU.RollNo =Convert.ToInt32(ab) + 1;

                                }
                                //TU.RollNo = Regno(rollno);
                                //TU.RollNo = a.RollNo;
                                TU.StudentName = a["Student Name"].ToString();
                                TU.FatherName = a["Father Name"].ToString();
                                TU.Address = a["Address"].ToString();
                                TU.Contact = a["Contact"].ToString();
                                TU.Laststudy = a["Last study"].ToString();
                                TU.Medical = a["Medical Done"].ToString();
                                TU.Refusal = a["Refusal"].ToString();
                                TU.Email = a["STUDENT E-MAIL ID"].ToString();
                                TU.Password = a["Password"].ToString();
                                TU.Fileno = "Jan/19";
                                db.StudentRegs.Add(TU);

                                db.SaveChanges();

                                fees fee = new fees();
                                fee.studentid = TU.Fileno + "" + TU.RollNo;
                                fee.Package = Convert.ToInt32(a["Package"]);
                                fee.Advance = Convert.ToInt32(a["Advance"]);
                                fee.pay = Convert.ToInt32(a["Advance"]);
                                fee.balance = fee.Package - fee.pay;
                                db.fees.Add(fee);
                                db.SaveChanges();

                                //    a.RollNo, a.StudentName, a.FatherName, a.Address, a.Contact, a.Laststudy, a.Medical, a.Refusal, a.Email, a.Password


                            }
                            else
                            {
                                //data.Add("<ul>");
                                //if (a.StudentName == "" || a.StudentName == null) data.Add("<li> name is required</li>");
                                //if (a.FatherName == "" || a.FatherName == null) data.Add("<li> Father Name is required</li>");
                                //if (a.Address == "" || a.Address == null) data.Add("<li> Address is required</li>");
                                //if (a.Contact == "" || a.Contact == null) data.Add("<li>ContactNo is required</li>");

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
                    // return RedirectToAction("Index");
                    return Json("success", JsonRequestBehavior.AllowGet);
                  //  return RedirectToAction("Index");
                    // Response.Redirect("Index");
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
            //StudentReg studentreg1 = db.StudentRegs.First();
            //if (studentreg1 != null)
            //{
            //    studentreg1.RollNo = 0001;
            //}
            //else
            //{
            //    var ab = db.StudentRegs.Max(x => x.RollNo);
            //    studentreg1.RollNo = Convert.ToInt32(ab) + 1;

            //}
            // return View(studentreg1);
            return View();
        }

        // POST: OfficialAdmin/StudentRegs/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Studentid,StudentName,FatherName,Address,Contact,Laststudy,Medical,Refusal,Email,Password,RollNo")] StudentReg studentReg, Helper Help,string package, string advance)
        {
            if (ModelState.IsValid)
            {
                StudentReg studentreg1 = db.StudentRegs.FirstOrDefault();
                if (studentreg1 == null)
                {
                    studentReg.RollNo = 0001;
                }
                else
                {
                    var ab = db.StudentRegs.Max(x => x.RollNo);
                    studentReg.RollNo = Convert.ToInt32(ab) + 1;

                }
                studentReg.Fileno = "Jan/19";
                db.StudentRegs.Add(studentReg);
                db.SaveChanges();

                fees fee = new fees();
                fee.studentid = studentReg.Fileno + "" + studentReg.RollNo;
                fee.Package = Convert.ToInt32(package);
                fee.Advance = Convert.ToInt32(advance);
                fee.pay = Convert.ToInt32(advance);
                fee.balance = fee.Package - fee.pay;
                db.fees.Add(fee);
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
            regno = studentReg.RollNo;
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
        public ActionResult Edit([Bind(Include = "Studentid,StudentName,FatherName,Address,Contact,Laststudy,Medical,Refusal,Email,Password")] StudentReg studentReg)
        {
            if (ModelState.IsValid)
            {
                studentReg.RollNo = regno;
                studentReg.Fileno = "Jan/19";
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
        public ActionResult Invoice(int id)
        {
           return RedirectToAction("Create","SingleFees",new { id=id});
        }
        public ActionResult FeeDetail(string roll)
        {
            return RedirectToAction("Index", "SingleFees", new { roll = roll });
        }
    }
}
