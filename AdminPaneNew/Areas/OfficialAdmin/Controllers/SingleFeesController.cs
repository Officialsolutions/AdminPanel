using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using AdminPaneNew.Areas.OfficialAdmin.Models;

namespace AdminPaneNew.Areas.OfficialAdmin.Controllers
{
    public class SingleFeesController : Controller
    {
        private dbcontext db = new dbcontext();
        public static string regno;
        public static int lastfee;

        // GET: OfficialAdmin/SingleFees
        public async Task<ActionResult> Index()
        {
            if(Session["user"]==null)
            {
                return RedirectToAction("Login", "Accounts");
            }
            return View(await db.SingleFees.ToListAsync());
        }

        // GET: OfficialAdmin/SingleFees/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            SingleFee singleFee = await db.SingleFees.FindAsync(id);
            if (singleFee == null)
            {
                return HttpNotFound();
            }
            return View(singleFee);
        }

        // GET: OfficialAdmin/SingleFees/Create
        public ActionResult Create(int id,SingleFee single2)
        {
            StudentReg stu = db.StudentRegs.Find(id);
            TempData["stuid"] = stu.Fileno + "" + stu.RollNo;
            TempData["id"] = stu.Studentid;
            regno= TempData["stuid"].ToString();
            var single = db.SingleFees.FirstOrDefault();
            if (single == null)
            {
                single2.Billno = Convert.ToInt32(100);
            }
            else
            {
                var single3 = db.SingleFees.Max(x=>x.Billno);
                single2.Billno = single3 + 1;

            }
          //  single2.studentid = TempData["stuid"].ToString();
            return View(single2);
        }

        // POST: OfficialAdmin/SingleFees/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "sfid,studentid,Date,Paid,Billno,Receivedby")] SingleFee singleFee,fees fee)
        {
            if (ModelState.IsValid)
            {
                singleFee.studentid = regno;
                singleFee.Receivedby = Session["user"].ToString();
                db.SingleFees.Add(singleFee);
                await db.SaveChangesAsync();
                fee = db.fees.Where(x => x.studentid == singleFee.studentid).FirstOrDefault();
                fee.pay = fee.pay + singleFee.Paid;
                fee.balance = fee.Package - fee.pay;
                db.Entry(fee).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(singleFee);
        }

        // GET: OfficialAdmin/SingleFees/Edit/5
        //[Authorize(Roles ="Admin")]
        
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            SingleFee singleFee = await db.SingleFees.FindAsync(id);
            regno = singleFee.studentid;
            lastfee = singleFee.Paid;
            if (singleFee == null)
            {
                return HttpNotFound();
            }
            return View(singleFee);
        }

        // POST: OfficialAdmin/SingleFees/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "sfid,studentid,Date,Paid,Billno,Receivedby")] SingleFee singleFee,fees fee)
        {
            if (ModelState.IsValid)
            {
               if(singleFee.Paid ==lastfee)
                {

                }
               else
                {
                    singleFee.studentid = regno;
                    fee = db.fees.Where(x => x.studentid == singleFee.studentid).FirstOrDefault();
                    fee.pay = fee.pay - lastfee;
                    fee.pay = fee.pay + singleFee.Paid;
                    fee.balance = fee.Package - fee.pay;
                    db.Entry(fee).State = EntityState.Modified;
                    db.SaveChanges();
                }
                singleFee.Receivedby = Session["user"].ToString();
                db.Entry(singleFee).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(singleFee);
        }
        // GET: OfficialAdmin/SingleFees/Delete/5
        //[Authorize(Roles = "Admin")]
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            SingleFee singleFee = await db.SingleFees.FindAsync(id);
            regno = singleFee.studentid;
            lastfee = singleFee.Paid;
            if (singleFee == null)
            {
                return HttpNotFound();
            }
            return View(singleFee);
        }

        // POST: OfficialAdmin/SingleFees/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id, fees fee)
        {
        
            SingleFee singleFee = await db.SingleFees.FindAsync(id);
            singleFee.studentid = regno;
            db.SingleFees.Remove(singleFee);
            await db.SaveChangesAsync();
            
            fee = db.fees.Where(x => x.studentid == singleFee.studentid).FirstOrDefault();
            fee.pay = fee.pay - lastfee;
          //  fee.pay = fee.pay + singleFee.Paid;
            fee.balance = fee.Package - fee.pay;
            db.Entry(fee).State = EntityState.Modified;
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
