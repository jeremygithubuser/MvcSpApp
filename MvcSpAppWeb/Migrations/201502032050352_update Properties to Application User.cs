namespace MvcSpAppWeb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class updatePropertiestoApplicationUser : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.AspNetUsers", "SharePointId", c => c.Int(nullable: false));
            AddColumn("dbo.AspNetUsers", "AccessToken", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.AspNetUsers", "AccessToken");
            DropColumn("dbo.AspNetUsers", "SharePointId");
        }
    }
}
