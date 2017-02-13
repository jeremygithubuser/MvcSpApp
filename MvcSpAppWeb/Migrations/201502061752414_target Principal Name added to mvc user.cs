namespace MvcSpAppWeb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class targetPrincipalNameaddedtomvcuser : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.AspNetUsers", "TargetPrincipalName", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.AspNetUsers", "TargetPrincipalName");
        }
    }
}
