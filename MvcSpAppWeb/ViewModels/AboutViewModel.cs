using MvcSpAppWeb.Poco;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MvcSpAppWeb.ViewModels
{
    public class AboutViewModel
    {
        private PictureUrl pictureUrl;
        public PictureUrl PictureUrl
        {
            get { return pictureUrl; }
            set { pictureUrl = value; }
        }

        private string displayName;
        public string DisplayName
        {
            get { return displayName; }
            set { displayName = value; }
        }

        private string email;
        public string Email
        {
            get { return email; }
            set { email = value; }
        }


        private int sharePointId;
        public int SharePointId
        {
            get { return sharePointId; }
            set { sharePointId = value; }
        }

        private string refreshToken;
        public string RefreshToken
        {
            get { return refreshToken; }
            set { refreshToken = value; }
        }

        private string accessToken;
        public string AccessToken
        {
            get { return accessToken; }
            set { accessToken = value; }
        }

        private string appOnlyAccessToken;
        public string AppOnlyAccessToken
        {
            get { return appOnlyAccessToken; }
            set { appOnlyAccessToken = value; }
        }

        private HostWebListCount hostWebListCount;
        public HostWebListCount HostWebListCount
        {
            get { return hostWebListCount; }
            set { hostWebListCount = value; }
        }



    }
}