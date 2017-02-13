using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using MvcSpAppWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;

namespace MvcSpAppWeb.CodeHelper
{
    public class RefreshAccessTokenHelper
    {


        public static void RefreshAccessToken(ApplicationUser mvcUser, ApplicationUserManager userManager)
        {

            OAuth2AccessTokenRequest OAaccessTokenRequest = OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(string.Format("{0}@{1}", WebConfigHelper.getClientIdFromWebConfig(), mvcUser.TargetRealm), WebConfigHelper.getClientSecretFromWebConfig(), mvcUser.RefreshToken, mvcUser.Ressource);
            // Obtenir un jeton
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            oauth2Response = client.Issue(mvcUser.SecurityTokenServiceUrl, OAaccessTokenRequest) as OAuth2AccessTokenResponse;
            mvcUser.AccessToken = oauth2Response.AccessToken;
            userManager.Update(mvcUser);

        }
        public static void RefreshAppOnlyAccessToken(ApplicationUser mvcUser, ApplicationUserManager userManager)
        {

            OAuth2AccessTokenRequest OAaccessTokenRequest = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(string.Format("{0}@{1}", WebConfigHelper.getClientIdFromWebConfig(), mvcUser.TargetRealm), WebConfigHelper.getClientSecretFromWebConfig(), mvcUser.Ressource);
            OAaccessTokenRequest.Resource = mvcUser.Ressource;
            // Obtenir un jeton
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            oauth2Response = client.Issue(mvcUser.SecurityTokenServiceUrl, OAaccessTokenRequest) as OAuth2AccessTokenResponse;
            mvcUser.AppOnlyAccessToken = oauth2Response.AccessToken;
            userManager.Update(mvcUser);
        }
    }
}