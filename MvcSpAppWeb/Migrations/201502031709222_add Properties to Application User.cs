namespace MvcSpAppWeb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class addPropertiestoApplicationUser : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.AspNetUsers", "TargetRealm", c => c.String());
            AddColumn("dbo.AspNetUsers", "Ressource", c => c.String());
            AddColumn("dbo.AspNetUsers", "SecurityTokenServiceUrl", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.AspNetUsers", "SecurityTokenServiceUrl");
            DropColumn("dbo.AspNetUsers", "Ressource");
            DropColumn("dbo.AspNetUsers", "TargetRealm");
        }
    }
}
