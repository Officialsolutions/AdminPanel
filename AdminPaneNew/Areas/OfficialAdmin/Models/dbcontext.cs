using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace AdminPaneNew.Areas.OfficialAdmin.Models
{
    public class dbcontext:DbContext
    {
        public dbcontext():base("dbcontext")
        {
          //  Database.SetInitializer(new MigrateDatabaseToLatestVersion<dbcontext, LLRM.Migrations.Configuration>("dbcontext"));
            //Database.SetInitializer<dbcontext>(new CreateDatabaseIfNotExists<dbcontext>());
     
            Database.SetInitializer(new MigrateDatabaseToLatestVersion<dbcontext, AdminPaneNew.Migrations.Configuration>("dbcontext"));
        }
        public DbSet<slider> sliders { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.news> news { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Service> Services { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Testimonial> Testimonials { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Album> Albums { get; set; }
        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Gallery> Galleries { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Logo> Logoes { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Clientlogo> Clientlogoes { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Contact> Contacts { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Features> Features { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Videos> Videos { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Account> Accounts { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.StudentDetail> StudentDetails { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.IeltsTest> IeltsTests { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.AssignTest> AssignTests { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.Category> Categories { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.StudentReg> StudentRegs { get; set; }

        public System.Data.Entity.DbSet<AdminPaneNew.Areas.OfficialAdmin.Models.fees> fees { get; set; }
        //protected override void OnModelCreating(DbModelBuilder modelBuilder)
        //{
        //    modelBuilder.Entity<StudentReg>().MapToStoredProcedures();
        //}
    }
}