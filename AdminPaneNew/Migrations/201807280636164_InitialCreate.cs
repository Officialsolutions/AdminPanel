namespace AdminPaneNew.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.sliders",
                c => new
                    {
                        Sliderid = c.Int(nullable: false, identity: true),
                        Name = c.String(nullable: false),
                        Description = c.String(),
                        Image = c.String(nullable: false),
                        url = c.String(),
                    })
                .PrimaryKey(t => t.Sliderid);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.sliders");
        }
    }
}
