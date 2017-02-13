namespace MvcSpAppWeb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AppOnlyAccessTokenaddedtomvcuser : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.AspNetUsers", "AppOnlyAccessToken", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.AspNetUsers", "AppOnlyAccessToken");
        }
    }
}
