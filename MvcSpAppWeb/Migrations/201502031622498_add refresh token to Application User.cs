namespace MvcSpAppWeb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class addrefreshtokentoApplicationUser : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.AspNetUsers", "refreshToken", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.AspNetUsers", "refreshToken");
        }
    }
}
