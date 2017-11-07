namespace PrintServices.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Jobs",
                c => new
                    {
                        EntryId = c.Int(nullable: false, identity: true),
                        FileName = c.String(),
                        PageCount = c.Int(nullable: false),
                        PrintQueue = c.String(),
                        Board = c.String(),
                        Disposition = c.String(),
                    })
                .PrimaryKey(t => t.EntryId);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Jobs");
        }
    }
}
