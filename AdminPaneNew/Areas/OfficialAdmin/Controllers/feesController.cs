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
    public class feesController : Controller
    {
        private dbcontext db = new dbcontext();

        // GET: OfficialAdmin/fees
        public async Task<ActionResult> Index()
        {
            return View(await db.fees.ToListAsync());
        }

        // GET: OfficialAdmin/fees/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            fees fees = await db.fees.FindAsync(id);
            if (fees == null)
            {
                return HttpNotFound();
            }
            return View(fees);
        }

        // GET: OfficialAdmin/fees/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: OfficialAdmin/fees/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "feeid,studentid,Package,Advance,pay")] fees fees)
        {
            if (ModelState.IsValid)
            {
                db.fees.Add(fees);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(fees);
        }

        // GET: OfficialAdmin/fees/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            fees fees = await db.fees.FindAsync(id);
            if (fees == null)
            {
                return HttpNotFound();
            }
            return View(fees);
        }

        // POST: OfficialAdmin/fees/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "feeid,studentid,Package,Advance,pay")] fees fees)
        {
            if (ModelState.IsValid)
            {
                db.Entry(fees).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(fees);
        }

        // GET: OfficialAdmin/fees/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            fees fees = await db.fees.FindAsync(id);
            if (fees == null)
            {
                return HttpNotFound();
            }
            return View(fees);
        }

        // POST: OfficialAdmin/fees/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            fees fees = await db.fees.FindAsync(id);
            db.fees.Remove(fees);
            await db.SaveChangesAsync();
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
