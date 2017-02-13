namespace MvcSpAppWeb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class targetHostaddedtomvcuser : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.AspNetUsers", "TargetHost", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.AspNetUsers", "TargetHost");
        }
    }
}
