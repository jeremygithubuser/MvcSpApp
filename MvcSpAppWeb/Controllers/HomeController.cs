using Microsoft.SharePoint.Client;
using MvcSpAppWeb.Models;
using System;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using System.Text;
using MvcSpAppWeb.CodeHelper;
using System.Collections.Specialized;
using MvcSpAppWeb.Poco;
using MvcSpAppWeb.ViewModels;
using MvcSpAppWeb.Config;


namespace MvcSpAppWeb.Controllers
{
    [RequireHttps]
    public class HomeController : Controller
    {
        private ApplicationSignInManager _signInManager;
        private ApplicationUserManager _userManager;

        public ApplicationSignInManager SignInManager
        {
            get
            {
                return _signInManager ?? HttpContext.GetOwinContext().Get<ApplicationSignInManager>();
            }
            private set
            {
                _signInManager = value;
            }
        }

        public ApplicationUserManager UserManager
        {
            get
            {
                return _userManager ?? HttpContext.GetOwinContext().GetUserManager<ApplicationUserManager>();
            }
            private set
            {
                _userManager = value;
            }
        }

        [SharePointContextFilter]
        public  ActionResult Index()
        {

            #region Constantes

            /*Ici une constante designant le produit SharePoint*/
            string SharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000";

            /* Nom du paramètre SPHostUrl*/
            /*Ici la clef pour obtenir la valeur du site hote dans SharePoint cette valeur est contenue dans la requete initiale*/
            var SPHostUrlKey = "SPHostUrl";

            /*Ici les clefs ID/Secret pour acceder aux valeurs du Web.config*/
            var webConfigId = System.Configuration.ConfigurationManager.AppSettings["ClientId"];
            var webConfigSecret = System.Configuration.ConfigurationManager.AppSettings["ClientSecret"];

            #endregion

            #region Variables Declarer un objet spUser qui contient quelques propriétés a mapper avec le User MVC correspondant
            HttpContext context = System.Web.HttpContext.Current;
            User spUser = null;
            #endregion

            #region Ici recuperer le spAppToken contenu dans la requete
            string SPAppToken = context.Request.Params.Get("SPAppToken");
            #endregion

            #region Ici demander au Token helper de valider que le spAppToken provient bien de la bonne plateforme. Penser a faire manuelement la validation du Token
            SharePointContextToken contextToken = TokenHelper.ReadAndValidateContextToken(SPAppToken, context.Request.Url.Authority);
            #endregion

            #region Ici extraire le refresh token
            string refreshToken = contextToken.RefreshToken;
            #endregion

            #region Recuper le DOMAINE de SharePoint , le GUID d'office 365 l'URL du Servive ACS qui peux delivrer un Access Token 
            /*Office 365 guid exemple d540fc2e-db24-4ad7-8c51-99f00b13f134*/
            string targetRealm = contextToken.Realm;
            /*Domaine du site Host Sharepoint exemple jeremycorp.sharepoint.com */
            string targetHost = context.Request.QueryString[SPHostUrlKey];
            Uri targetUri = new Uri(targetHost);
            targetHost = targetUri.Authority;

            /*SharePoint Principal+sp domain+office 365 guid*/
            string ressource = string.Format("{0}/{1}@{2}", SharePointPrincipal, targetHost, targetRealm);
            string targetPrincipalName = contextToken.TargetPrincipalName;

            /*Adresse de l'identity provider*/
            string securityTokenServiceUrl = TokenHelper.AcsMetadataParser.GetStsUrl(targetRealm);
            #endregion



            #region Grace a L'API Sharepoint on recupere le spUser Courrant
            /*Grace a L'API Sharepoint on recupere le spUser Courrant*/
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser);
                    clientContext.ExecuteQuery();
                    ViewBag.UserName = spUser.Title;
                }
            }
            #endregion

            #region Encapsuler les informations
            /*Encapsuler les informations*/
            /*NameValueCollection */
            NameValueCollection userProperties = new NameValueCollection();
            userProperties.Add("RefreshToken", refreshToken);
            userProperties.Add("TargetRealm", targetRealm);
            userProperties.Add("Ressource", ressource);
            userProperties.Add("SecurityTokenServiceUrl", securityTokenServiceUrl);
            userProperties.Add("TargetHost", targetHost);
            userProperties.Add("TargetPrincipalName", targetPrincipalName);
            #endregion

            #region Read Tokens
            /*
            HttpWebRequest endpointRequest =(HttpWebRequest)HttpWebRequest.Create("https://jeremycorp.sharepoint.com/_api/web/siteusers?format=json");
            endpointRequest.Method = "GET";
            endpointRequest.Accept = "application/json;odata=verbose";
            endpointRequest.Headers.Add("Authorization",
              "Bearer " + accesToken);
            HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();

            HttpWebRequest currentUserRequest = (HttpWebRequest)HttpWebRequest.Create("https://jeremycorp.sharepoint.com/_api/web/siteusers?$filter=UserId/NameId eq '10037ffe8c9efa88'&$format=json");
            currentUserRequest.Method = "GET";
            currentUserRequest.Accept = "application/json;odata=verbose";
            currentUserRequest.Headers.Add("Authorization",
              "Bearer " + accesToken);
            HttpWebResponse currentUserResponse = (HttpWebResponse)currentUserRequest.GetResponse();

            
            var data = context;

            string[] SPAppTokenTokenArray = SPAppToken.Split('.');
            byte[] SPAppTokenHeaderBytes = B64Helper.base64urldecode(SPAppTokenTokenArray[0]);
            byte[] SPAppTokenClaimsBytes = B64Helper.base64urldecode(SPAppTokenTokenArray[1]);
            SPAppToken = Encoding.UTF8.GetString(SPAppTokenHeaderBytes) + Encoding.UTF8.GetString(SPAppTokenClaimsBytes);
            
            string[] accesTokenArray = accesToken.Split('.');
            byte[] accesTokenHeaderBytes = B64Helper.base64urldecode(accesTokenArray[0]);
            byte[] accesTokenClaimsBytes = B64Helper.base64urldecode(accesTokenArray[1]);
            accesToken = Encoding.UTF8.GetString(accesTokenHeaderBytes) + Encoding.UTF8.GetString(accesTokenClaimsBytes);

            
            Stream dataStream = endpointResponse.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string responseFromServer = reader.ReadToEnd();


            Stream dataUserStream = currentUserResponse.GetResponseStream();
            StreamReader userReader = new StreamReader(dataUserStream);
            string userResponseFromServer = userReader.ReadToEnd();

            

            ViewBag.SPAppToken = SPAppToken;
            ViewBag.contextToken = contextToken;
            ViewBag.refreshToken = refreshToken;
            ViewBag.accesToken = accesToken;
            ViewBag.endpointResponse = endpointResponse;
            ViewBag.responseFromServer = responseFromServer;
            ViewBag.userResponseFromServer = userResponseFromServer;      
            return View(data);*/

            #endregion

            #region Storer le SPuser ainsi que les propriétés du futur MVC User dans TempData
            TempData["currentUser"] = spUser;
            TempData["userProperties"] = userProperties;
            #endregion

            //#region Rediriger vers le Controller d'Inscription/Gestion des Utilisateurs
            //return RedirectToAction("Register");
            //#endregion
            #region Si l'Utilisateur n'est pas inscrit.           
            var mvcUser = UserManager.FindByName(spUser.LoginName);
            if (mvcUser == null)
            {
                mvcUser = new ApplicationUser { UserName = spUser.LoginName, Email = spUser.Email };
                var result =  UserManager.Create(mvcUser, WebConfigHelper.getDummyPasswordFromWebConfig());
                if (result.Succeeded)
                {
                    mvcUser.SharePointId = spUser.Id;
                    mvcUser.DisplayName = spUser.Title;
                    mvcUser.RefreshToken = userProperties.Get("RefreshToken");
                    mvcUser.TargetRealm = userProperties.Get("TargetRealm");
                    mvcUser.Ressource = userProperties.Get("Ressource");
                    mvcUser.SecurityTokenServiceUrl = userProperties.Get("SecurityTokenServiceUrl");
                    mvcUser.TargetHost = userProperties.Get("TargetHost");
                    mvcUser.TargetPrincipalName = userProperties.Get("TargetPrincipalName");
                    UserManager.Update(mvcUser);
                    SignInManager.SignIn(mvcUser, isPersistent: false, rememberBrowser: false);

                    // Pour plus d'informations sur l'activation de la confirmation du compte et la réinitialisation du mot de passe, consultez http://go.microsoft.com/fwlink/?LinkID=320771
                    // Envoyer un message électronique avec ce lien
                    // string code = await UserManager.GenerateEmailConfirmationTokenAsync(user.Id);
                    // var callbackUrl = Url.Action("ConfirmEmail", "Account", new { userId = user.Id, code = code }, protocol: Request.Url.Scheme);
                    // await UserManager.SendEmailAsync(user.Id, "Confirmez votre compte", "Confirmez votre compte en cliquant <a href=\"" + callbackUrl + "\">ici</a>");
                    return RedirectToAction("About");
                }
            }
            #endregion

            #region Si l'Utilisateur est déja inscrit.
            else
            {
                mvcUser.SharePointId = spUser.Id;
                mvcUser.DisplayName = spUser.Title;
                mvcUser.Email = spUser.Email;
                mvcUser.RefreshToken = userProperties.Get("RefreshToken");
                mvcUser.TargetRealm = userProperties.Get("TargetRealm");
                mvcUser.Ressource = userProperties.Get("Ressource");
                mvcUser.SecurityTokenServiceUrl = userProperties.Get("SecurityTokenServiceUrl");
                mvcUser.TargetHost = userProperties.Get("TargetHost");
                mvcUser.TargetPrincipalName = userProperties.Get("TargetPrincipalName");
                UserManager.Update(mvcUser);
                SignInManager.SignIn(mvcUser, isPersistent: false, rememberBrowser: false);
                return RedirectToAction("About");
            }
            #endregion

            #region Rediriger vers La page About
            return RedirectToAction("About");
            #endregion

        }

        //public async Task<ActionResult> Register()
        //{
        //    #region Recuper l'objet spUser + la collection de propriétés avec laquelle on definiera les propriétés de l'utilisateur MVC
        //    /*Recuper l'objet spUser + la collection de propriétés avec laquelle on definiera les propriétés de l'utilisateur MVC*/
        //    NameValueCollection userProperties = TempData["userProperties"] as NameValueCollection;
        //    User spUser = TempData["currentUser"] as User;
        //    #endregion

        //    #region Si l'Utilisateur n'est pas inscrit.           
        //    var mvcUser = UserManager.FindByName(spUser.LoginName);
        //    if (mvcUser == null)
        //    {
        //        mvcUser = new ApplicationUser { UserName = spUser.LoginName, Email = spUser.Email };
        //        var result = await UserManager.CreateAsync(mvcUser, WebConfigHelper.getDummyPasswordFromWebConfig());
        //        if (result.Succeeded)
        //        {
        //            mvcUser.SharePointId = spUser.Id;
        //            mvcUser.DisplayName = spUser.Title;
        //            mvcUser.RefreshToken = userProperties.Get("RefreshToken");
        //            mvcUser.TargetRealm = userProperties.Get("TargetRealm");
        //            mvcUser.Ressource = userProperties.Get("Ressource");
        //            mvcUser.SecurityTokenServiceUrl = userProperties.Get("SecurityTokenServiceUrl");
        //            mvcUser.TargetHost = userProperties.Get("TargetHost");
        //            mvcUser.TargetPrincipalName = userProperties.Get("TargetPrincipalName");
        //            UserManager.Update(mvcUser);
        //            await SignInManager.SignInAsync(mvcUser, isPersistent: false, rememberBrowser: false);

        //            // Pour plus d'informations sur l'activation de la confirmation du compte et la réinitialisation du mot de passe, consultez http://go.microsoft.com/fwlink/?LinkID=320771
        //            // Envoyer un message électronique avec ce lien
        //            // string code = await UserManager.GenerateEmailConfirmationTokenAsync(user.Id);
        //            // var callbackUrl = Url.Action("ConfirmEmail", "Account", new { userId = user.Id, code = code }, protocol: Request.Url.Scheme);
        //            // await UserManager.SendEmailAsync(user.Id, "Confirmez votre compte", "Confirmez votre compte en cliquant <a href=\"" + callbackUrl + "\">ici</a>");
        //            return RedirectToAction("About");
        //        }
        //    }
        //    #endregion

        //    #region Si l'Utilisateur est déja inscrit.
        //    else
        //    {
        //        mvcUser.SharePointId = spUser.Id;
        //        mvcUser.DisplayName = spUser.Title;
        //        mvcUser.Email = spUser.Email;
        //        mvcUser.RefreshToken = userProperties.Get("RefreshToken");
        //        mvcUser.TargetRealm = userProperties.Get("TargetRealm");
        //        mvcUser.Ressource = userProperties.Get("Ressource");
        //        mvcUser.SecurityTokenServiceUrl = userProperties.Get("SecurityTokenServiceUrl");
        //        mvcUser.TargetHost = userProperties.Get("TargetHost");
        //        mvcUser.TargetPrincipalName = userProperties.Get("TargetPrincipalName");
        //        UserManager.Update(mvcUser);
        //        await SignInManager.SignInAsync(mvcUser, isPersistent: false, rememberBrowser: false);
        //        return RedirectToAction("About");
        //    }
        //    #endregion

        //    #region Rediriger vers La page About
        //    return RedirectToAction("About");
        //    #endregion

        //}

        [Authorize]
        public ActionResult About()
        {
            #region Instancier le Modele de la vue
            AboutViewModel viewModel = new AboutViewModel();
            viewModel.PictureUrl = new PictureUrl();
            viewModel.HostWebListCount = new HostWebListCount();
            #endregion

            #region Recuperer le current User
            ApplicationUser currentUser = UserManager.FindById(User.Identity.GetUserId());
            RefreshAccessTokenHelper.RefreshAccessToken(currentUser, UserManager);
            RefreshAccessTokenHelper.RefreshAppOnlyAccessToken(currentUser, UserManager);
            #endregion

            #region Recuperer L'adresse de l'image du profil
            viewModel.PictureUrl.value = SpHelper.getCurrentUserPictureUrl(currentUser, UserManager, Configuration.hostwebUrl);
            #endregion

            #region Recuperer le nombre de Listes dans le site
            viewModel.HostWebListCount.value = SpHelper.getHostWebListsCount(currentUser, UserManager, Configuration.hostwebUrl);
            #endregion

            #region Convertir l'App only Access Token de B64 vers UTF8
            string[] accesTokenArray = currentUser.AppOnlyAccessToken.Split('.');
            byte[] accesTokenHeaderBytes = B64Helper.base64urldecode(accesTokenArray[0]);
            byte[] accesTokenClaimsBytes = B64Helper.base64urldecode(accesTokenArray[1]);
            string accessToken = Encoding.UTF8.GetString(accesTokenHeaderBytes) + Encoding.UTF8.GetString(accesTokenClaimsBytes);
            #endregion

            #region Definir les propriétés du viewModel
            viewModel.AccessToken = accessToken;
            viewModel.DisplayName = currentUser.DisplayName;
            viewModel.Email = currentUser.Email;
            viewModel.RefreshToken = currentUser.RefreshToken;
            viewModel.SharePointId = currentUser.SharePointId;
            viewModel.AppOnlyAccessToken = currentUser.AppOnlyAccessToken;
            #endregion

            #region Try to get PictureUrl Bytes
            /* string responseFromServer = "";
            try
            {
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create("https://jeremycorp-my.sharepoint.com:443/_api/web/getfilebyserverrelativeurl('/User%20Photos/Images%20du%20profil/jeremy_jeremycorp_onmicrosoft_com_MThumb.jpg')/$value");
                endpointRequest.Method = "GET";
                // endpointRequest.Accept = "application/json;odata=nometadata";
                endpointRequest.Headers.Add("Authorization",
                "Bearer " + UserManager.FindById(User.Identity.GetUserId()).AppOnlyAccessToken);
                HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
                Stream dataStream = endpointResponse.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                 responseFromServer = reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                
                 responseFromServer = ex.Message;
            }
            ViewBag.responseFromServer = responseFromServer;*/

            #endregion

            return View(viewModel);
        }

        [Authorize]
        public ActionResult Contact()
        {
            ViewBag.Message = "Page de contact pour MVC SP APP";
            return View();
        }
    }
}
