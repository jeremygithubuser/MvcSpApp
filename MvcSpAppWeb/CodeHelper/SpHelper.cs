using MvcSpAppWeb.Models;
using MvcSpAppWeb.Poco;
using MvcSpAppWeb.Config;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Script.Serialization;

namespace MvcSpAppWeb.CodeHelper
{
    public class SpHelper
    {
        public static string getCurrentUserPictureUrl(ApplicationUser mvcUser, ApplicationUserManager UserManager, string spHostUrl)
        {
            try
            {
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(spHostUrl + "/_api/sp.userprofiles.peoplemanager/getuserprofilepropertyfor(accountName=@v, propertyName='PictureURL')?@v='" + HttpUtility.UrlEncode(mvcUser.UserName) + "'");
                endpointRequest.Method = "GET";
                endpointRequest.Accept = "application/json;odata=nometadata";
                endpointRequest.Headers.Add("Authorization",
                  "Bearer " + mvcUser.AppOnlyAccessToken);
                HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
                Stream dataStream = endpointResponse.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                #region Deserialiser le Json en un Objet
                var serializer = new JavaScriptSerializer();
                PictureUrl deserializedResult = serializer.Deserialize<PictureUrl>(responseFromServer);
                #endregion
                return deserializedResult.value;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public static string getHostWebListsCount(ApplicationUser mvcUser, ApplicationUserManager UserManager, string spHostUrl)
        {
            try
            {
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(spHostUrl + "/_api/web/Lists/getbytitle('Administration')/ItemCount");
                endpointRequest.Method = "GET";
                endpointRequest.Accept = "application/json;odata=nometadata";
                endpointRequest.Headers.Add("Authorization",
                  "Bearer " + mvcUser.AppOnlyAccessToken);
                HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
                Stream dataStream = endpointResponse.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                #region Deserialiser le Json en un Objet
                var serializer = new JavaScriptSerializer();
                HostWebListCount deserializedResult = serializer.Deserialize<HostWebListCount>(responseFromServer);
                #endregion
                return deserializedResult.value;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public static string getFollowersFor(ApplicationUser mvcUser, ApplicationUserManager UserManager, string spHostUrl)
        {
            try
            {
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(spHostUrl + "/_api/sp.userprofiles.peoplemanager/getfollowersfor(@v)?@v='" + HttpUtility.UrlEncode(mvcUser.UserName) + "'");
                endpointRequest.Method = "GET";
                endpointRequest.Accept = "application/json;odata=nometadata";
                endpointRequest.Headers.Add("Authorization",
                  "Bearer " + mvcUser.AppOnlyAccessToken);
                HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
                Stream dataStream = endpointResponse.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                #region Deserialiser le Json en un Objet
                var serializer = new JavaScriptSerializer();
                HostWebListCount deserializedResult = serializer.Deserialize<HostWebListCount>(responseFromServer);
                #endregion
                return deserializedResult.value;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public static List<SpListMetadataModel> getAllSpLists(ApplicationUser mvcUser, ApplicationUserManager UserManager, string spHostUrl)
        {
            try
            {
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(spHostUrl + "/_api/web/lists?format=json");
                endpointRequest.Method = "GET";
                endpointRequest.Accept = "application/json;odata=nometadata";
                endpointRequest.Headers.Add("Authorization",
                  "Bearer " + mvcUser.AppOnlyAccessToken);
                HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
                Stream dataStream = endpointResponse.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                #region Deserialiser le Json en un Objet
                var serializer = new JavaScriptSerializer();
                List<SpListMetadataModel> deserializedResult = serializer.Deserialize<SpListMetadataModelWrapper>(responseFromServer).value;

                #endregion
                return deserializedResult;
            }
            catch (Exception)
            {
                return new List<SpListMetadataModel>();
            }
        }
        public static List<Dictionary<string, string>> getSpListById(ApplicationUser mvcUser, ApplicationUserManager UserManager, string listId, string spHostUrl)
        {
            try
            {
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(spHostUrl + "/_api/web/lists(guid'" + listId + "')/Items?format=json");
                endpointRequest.Method = "GET";
                endpointRequest.Accept = "application/json;odata=nometadata";
                endpointRequest.Headers.Add("Authorization",
                  "Bearer " + mvcUser.AppOnlyAccessToken);
                HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
                Stream dataStream = endpointResponse.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                #region Deserialiser le Json en un Objet
                var serializer = new JavaScriptSerializer();
                serializer.RecursionLimit = 4;
                /*HostWebListCount deserializedResult = serializer.Deserialize<HostWebListCount>(responseFromServer);*/
                //Dictionary<string, string> dictionnary = serializer.Deserialize<SpListItemCollectionModelWrapper>(responseFromServer).value;
                var dictionnaryList = serializer.Deserialize<SpListItemCollectionModelWrapper>(responseFromServer).value;
                #endregion
                /*return deserializedResult.value;*/
                return dictionnaryList;
            }
            catch (Exception ex)
            {
                List<Dictionary<string, string>> dictionnaryList = new List<Dictionary<string, string>>();
                dictionnaryList[0] = new Dictionary<string, string>();
                dictionnaryList[0].Add("Message", ex.Message);
                return dictionnaryList;
            }
        }


    }
}