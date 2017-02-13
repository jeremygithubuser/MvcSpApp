using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MvcSpAppWeb.CodeHelper
{
    public class WebConfigHelper
    {
        public static string getClientIdFromWebConfig()
        {
            return System.Configuration.ConfigurationManager.AppSettings["ClientId"];
        }
        public static string getClientSecretFromWebConfig()
        {
            return System.Configuration.ConfigurationManager.AppSettings["ClientSecret"]; ;
        }

        public static string getDummyPasswordFromWebConfig()
        {
            return System.Configuration.ConfigurationManager.AppSettings["DummyPassword"]; ;
        }


    }
}