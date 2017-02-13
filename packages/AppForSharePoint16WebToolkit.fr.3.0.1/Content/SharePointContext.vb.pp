Imports Microsoft.IdentityModel.S2S.Protocols.OAuth2
Imports Microsoft.IdentityModel.Tokens
Imports Microsoft.SharePoint.Client
Imports System
Imports System.Net
Imports System.Security.Principal
Imports System.Web
Imports System.Web.Configuration

''' <summary>
''' Encapsule toutes les informations de SharePoint.
''' </summary>
Public MustInherit Class SharePointContext
    Public Const SPHostUrlKey As String = "SPHostUrl"
    Public Const SPAppWebUrlKey As String = "SPAppWebUrl"
    Public Const SPLanguageKey As String = "SPLanguage"
    Public Const SPClientTagKey As String = "SPClientTag"
    Public Const SPProductNumberKey As String = "SPProductNumber"

    Protected Shared ReadOnly AccessTokenLifetimeTolerance As TimeSpan = TimeSpan.FromMinutes(5.0)

    Private ReadOnly m_spHostUrl As Uri
    Private ReadOnly m_spAppWebUrl As Uri
    Private ReadOnly m_spLanguage As String
    Private ReadOnly m_spClientTag As String
    Private ReadOnly m_spProductNumber As String

    ' <AccessTokenString, UtcExpiresOn>
    Protected m_userAccessTokenForSPHost As Tuple(Of String, DateTime)
    Protected m_userAccessTokenForSPAppWeb As Tuple(Of String, DateTime)
    Protected m_appOnlyAccessTokenForSPHost As Tuple(Of String, DateTime)
    Protected m_appOnlyAccessTokenForSPAppWeb As Tuple(Of String, DateTime)

    ''' <summary>
    ''' Obtient l'URL de l'hôte SharePoint à partir du paramètre QueryString de la requête HTTP spécifiée.
    ''' </summary>
    ''' <param name="httpRequest">Requête HTTP spécifiée.</param>
    ''' <returns>URL de l'hôte SharePoint. Retourne <c>Nothing</c> si la requête HTTP ne contient pas l'URL de l'hôte SharePoint.</returns>
    Public Shared Function GetSPHostUrl(httpRequest As HttpRequestBase) As Uri
        If httpRequest Is Nothing Then
            Throw New ArgumentNullException("httpRequest")
        End If

        Dim spHostUrlString As String = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString(SPHostUrlKey))
        Dim spHostUrl As Uri = Nothing
        If Uri.TryCreate(spHostUrlString, UriKind.Absolute, spHostUrl) AndAlso
           (spHostUrl.Scheme = Uri.UriSchemeHttp OrElse spHostUrl.Scheme = Uri.UriSchemeHttps) Then
            Return spHostUrl
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Obtient l'URL de l'hôte SharePoint à partir du paramètre QueryString de la requête HTTP spécifiée.
    ''' </summary>
    ''' <param name="httpRequest">Requête HTTP spécifiée.</param>
    ''' <returns>URL de l'hôte SharePoint. Retourne <c>Nothing</c> si la requête HTTP ne contient pas l'URL de l'hôte SharePoint.</returns>
    Public Shared Function GetSPHostUrl(httpRequest As HttpRequest) As Uri
        Return GetSPHostUrl(New HttpRequestWrapper(httpRequest))
    End Function

    ''' <summary>
    ''' URL de l'hôte SharePoint.
    ''' </summary>
    Public ReadOnly Property SPHostUrl() As Uri
        Get
            Return Me.m_spHostUrl
        End Get
    End Property

    ''' <summary>
    ''' URL de l'application Web SharePoint.
    ''' </summary>
    Public ReadOnly Property SPAppWebUrl() As Uri
        Get
            Return Me.m_spAppWebUrl
        End Get
    End Property

    ''' <summary>
    ''' Langue SharePoint.
    ''' </summary>
    Public ReadOnly Property SPLanguage() As String
        Get
            Return Me.m_spLanguage
        End Get
    End Property

    ''' <summary>
    ''' Balise cliente SharePoint.
    ''' </summary>
    Public ReadOnly Property SPClientTag() As String
        Get
            Return Me.m_spClientTag
        End Get
    End Property

    ''' <summary>
    ''' Numéro de produit SharePoint.
    ''' </summary>
    Public ReadOnly Property SPProductNumber() As String
        Get
            Return Me.m_spProductNumber
        End Get
    End Property

    ''' <summary>
    ''' Jeton d'accès utilisateur de l'hôte SharePoint.
    ''' </summary>
    Public MustOverride ReadOnly Property UserAccessTokenForSPHost() As String

    ''' <summary>
    ''' Jeton d'accès utilisateur de l'application Web SharePoint.
    ''' </summary>
    Public MustOverride ReadOnly Property UserAccessTokenForSPAppWeb() As String

    ''' <summary>
    ''' Jeton d'accès pour l'application uniquement de l'hôte SharePoint.
    ''' </summary>
    Public MustOverride ReadOnly Property AppOnlyAccessTokenForSPHost() As String

    ''' <summary>
    ''' Jeton d'accès pour l'application uniquement pour l'application Web SharePoint.
    ''' </summary>
    Public MustOverride ReadOnly Property AppOnlyAccessTokenForSPAppWeb() As String

    ''' <summary>
    ''' Constructeur.
    ''' </summary>
    ''' <param name="spHostUrl">URL de l'hôte SharePoint.</param>
    ''' <param name="spAppWebUrl">URL de l'application Web SharePoint.</param>
    ''' <param name="spLanguage">Langue SharePoint.</param>
    ''' <param name="spClientTag">Balise cliente SharePoint.</param>
    ''' <param name="spProductNumber">Numéro de produit SharePoint.</param>
    Protected Sub New(spHostUrl As Uri, spAppWebUrl As Uri, spLanguage As String, spClientTag As String, spProductNumber As String)
        If spHostUrl Is Nothing Then
            Throw New ArgumentNullException("spHostUrl")
        End If

        If String.IsNullOrEmpty(spLanguage) Then
            Throw New ArgumentNullException("spLanguage")
        End If

        If String.IsNullOrEmpty(spClientTag) Then
            Throw New ArgumentNullException("spClientTag")
        End If

        If String.IsNullOrEmpty(spProductNumber) Then
            Throw New ArgumentNullException("spProductNumber")
        End If

        Me.m_spHostUrl = spHostUrl
        Me.m_spAppWebUrl = spAppWebUrl
        Me.m_spLanguage = spLanguage
        Me.m_spClientTag = spClientTag
        Me.m_spProductNumber = spProductNumber
    End Sub

    ''' <summary>
    ''' Crée un utilisateur ClientContext pour l'hôte SharePoint.
    ''' </summary>
    ''' <returns>Instance ClientContext.</returns>
    Public Function CreateUserClientContextForSPHost() As ClientContext
        Return CreateClientContext(Me.SPHostUrl, Me.UserAccessTokenForSPHost)
    End Function

    ''' <summary>
    ''' Crée un utilisateur ClientContext pour l'application Web SharePoint.
    ''' </summary>
    ''' <returns>Instance ClientContext.</returns>
    Public Function CreateUserClientContextForSPAppWeb() As ClientContext
        Return CreateClientContext(Me.SPAppWebUrl, Me.UserAccessTokenForSPAppWeb)
    End Function

    ''' <summary>
    ''' Crée un ClientContext pour l'application uniquement pour l'hôte SharePoint.
    ''' </summary>
    ''' <returns>Instance ClientContext.</returns>
    Public Function CreateAppOnlyClientContextForSPHost() As ClientContext
        Return CreateClientContext(Me.SPHostUrl, Me.AppOnlyAccessTokenForSPHost)
    End Function

    ''' <summary>
    ''' Crée un ClientContext pour l'application uniquement pour l'application Web SharePoint.
    ''' </summary>
    ''' <returns>Instance ClientContext.</returns>
    Public Function CreateAppOnlyClientContextForSPAppWeb() As ClientContext
        Return CreateClientContext(Me.SPAppWebUrl, Me.AppOnlyAccessTokenForSPAppWeb)
    End Function

    ''' <summary>
    ''' Obtient la chaîne de connexion de la base de données de SharePoint pour l'application hébergée.
    ''' </summary>
    ''' <returns>Chaîne de connexion de la base de données. Retourne <c>null</c> si l'application n'est pas autohébergée ou qu'il n'y a aucune base de données.</returns>
    Public Function GetDatabaseConnectionString() As String
        Dim dbConnectionString As String = Nothing

        Using clientContext As ClientContext = CreateAppOnlyClientContextForSPHost()
            If clientContext IsNot Nothing Then
                Dim result = AppInstance.RetrieveAppDatabaseConnectionString(clientContext)

                clientContext.ExecuteQuery()

                dbConnectionString = result.Value
            End If
        End Using

        If dbConnectionString Is Nothing Then
            Const LocalDBInstanceForDebuggingKey As String = "LocalDBInstanceForDebugging"

            Dim dbConnectionStringSettings = WebConfigurationManager.ConnectionStrings(LocalDBInstanceForDebuggingKey)

            dbConnectionString = If(dbConnectionStringSettings IsNot Nothing, dbConnectionStringSettings.ConnectionString, Nothing)
        End If

        Return dbConnectionString
    End Function

    ''' <summary>
    ''' Détermine si le jeton d'accès spécifié est valide.
    ''' Un jeton d'accès est considéré comme non valide s'il a la valeur Nothing, ou qu'il a expiré.
    ''' </summary>
    ''' <param name="accessToken">Jeton d'accès à vérifier.</param>
    ''' <returns>True si le jeton d'accès est valide.</returns>
    Protected Shared Function IsAccessTokenValid(accessToken As Tuple(Of String, DateTime)) As Boolean
        Return accessToken IsNot Nothing AndAlso
               Not String.IsNullOrEmpty(accessToken.Item1) AndAlso
               accessToken.Item2 > DateTime.UtcNow
    End Function

    ''' <summary>
    ''' Crée un ClientContext avec l'URL du site SharePoint spécifié et le jeton d'accès.
    ''' </summary>
    ''' <param name="spSiteUrl">URL du site.</param>
    ''' <param name="accessToken">Jeton d'accès.</param>
    ''' <returns>Instance ClientContext.</returns>
    Private Shared Function CreateClientContext(spSiteUrl As Uri, accessToken As String) As ClientContext
        If spSiteUrl IsNot Nothing AndAlso Not String.IsNullOrEmpty(accessToken) Then
            Return TokenHelper.GetClientContextWithAccessToken(spSiteUrl.AbsoluteUri, accessToken)
        End If

        Return Nothing
    End Function
End Class

''' <summary>
''' Statut de redirection.
''' </summary>
Public Enum RedirectionStatus
    Ok
    ShouldRedirect
    CanNotRedirect
End Enum

''' <summary>
''' Fournit les instances SharePointContext.
''' </summary>
Public MustInherit Class SharePointContextProvider
    Private Shared s_current As SharePointContextProvider

    ''' <summary>
    ''' Instance SharePointContextProvider actuelle.
    ''' </summary>
    Public Shared ReadOnly Property Current() As SharePointContextProvider
        Get
            Return SharePointContextProvider.s_current
        End Get
    End Property

    ''' <summary>
    ''' Initialise l'instance par défaut SharePointContextProvider.
    ''' </summary>
    Shared Sub New()
        If Not TokenHelper.IsHighTrustApp() Then
            SharePointContextProvider.s_current = New SharePointAcsContextProvider()
        Else
            SharePointContextProvider.s_current = New SharePointHighTrustContextProvider()
        End If
    End Sub

    ''' <summary>
    ''' Inscrit l'instance SharePointContextProvider spécifiée comme l'instance actuelle.
    ''' Doit être appelé par Application_Start() dans Global.asax.
    ''' </summary>
    ''' <param name="provider">SharePointContextProvider à définir comme actif.</param>
    Public Shared Sub Register(provider As SharePointContextProvider)
        If provider Is Nothing Then
            Throw New ArgumentNullException("provider")
        End If

        SharePointContextProvider.s_current = provider
    End Sub

    ''' <summary>
    ''' Vérifie s'il est nécessaire de rediriger vers SharePoint pour l'authentification de l'utilisateur.
    ''' </summary>
    ''' <param name="httpContext">Contexte HTTP.</param>
    ''' <param name="redirectUri">URL de redirection vers SharePoint si le statut est ShouldRedirect. <c>Null</c> si le statut est Ok ou CanNotRedirect.</param>
    ''' <returns>Statut de redirection.</returns>
    Public Shared Function CheckRedirectionStatus(httpContext As HttpContextBase, ByRef redirectUrl As Uri) As RedirectionStatus
        If httpContext Is Nothing Then
            Throw New ArgumentNullException("httpContext")
        End If

        redirectUrl = Nothing

        If SharePointContextProvider.Current.GetSharePointContext(httpContext) IsNot Nothing Then
            Return RedirectionStatus.Ok
        End If

        Const SPHasRedirectedToSharePointKey As String = "SPHasRedirectedToSharePoint"

        If Not String.IsNullOrEmpty(httpContext.Request.QueryString(SPHasRedirectedToSharePointKey)) Then
            Return RedirectionStatus.CanNotRedirect
        End If

        Dim spHostUrl As Uri = SharePointContext.GetSPHostUrl(httpContext.Request)

        If spHostUrl Is Nothing Then
            Return RedirectionStatus.CanNotRedirect
        End If

        If StringComparer.OrdinalIgnoreCase.Equals(httpContext.Request.HttpMethod, "POST") Then
            Return RedirectionStatus.CanNotRedirect
        End If

        Dim requestUrl As Uri = httpContext.Request.Url

        Dim queryNameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query)

        ' Supprime les valeurs incluses dans {StandardTokens}, car {StandardTokens} sera inséré au début de la chaîne de requête.
        queryNameValueCollection.Remove(SharePointContext.SPHostUrlKey)
        queryNameValueCollection.Remove(SharePointContext.SPAppWebUrlKey)
        queryNameValueCollection.Remove(SharePointContext.SPLanguageKey)
        queryNameValueCollection.Remove(SharePointContext.SPClientTagKey)
        queryNameValueCollection.Remove(SharePointContext.SPProductNumberKey)

        ' Ajoute SPHasRedirectedToSharePoint=1.
        queryNameValueCollection.Add(SPHasRedirectedToSharePointKey, "1")

        Dim returnUrlBuilder As New UriBuilder(requestUrl)
        returnUrlBuilder.Query = queryNameValueCollection.ToString()

        ' Insère StandardTokens.
        Const StandardTokens As String = "{StandardTokens}"
        Dim returnUrlString As String = returnUrlBuilder.Uri.AbsoluteUri
        returnUrlString = returnUrlString.Insert(returnUrlString.IndexOf("?") + 1, StandardTokens + "&")

        ' Construit une URL de redirection.
        Dim redirectUrlString As String = TokenHelper.GetAppContextTokenRequestUrl(spHostUrl.AbsoluteUri, Uri.EscapeDataString(returnUrlString))

        redirectUrl = New Uri(redirectUrlString, UriKind.Absolute)

        Return RedirectionStatus.ShouldRedirect
    End Function

    ''' <summary>
    ''' Vérifie s'il est nécessaire de rediriger vers SharePoint pour l'authentification de l'utilisateur.
    ''' </summary>
    ''' <param name="httpContext">Contexte HTTP.</param>
    ''' <param name="redirectUri">URL de redirection vers SharePoint si le statut est ShouldRedirect. <c>Null</c> si le statut est Ok ou CanNotRedirect.</param>
    ''' <returns>Statut de redirection.</returns>
    Public Shared Function CheckRedirectionStatus(httpContext As HttpContext, ByRef redirectUrl As Uri) As RedirectionStatus
        Return CheckRedirectionStatus(New HttpContextWrapper(httpContext), redirectUrl)
    End Function

    ''' <summary>
    ''' Crée une instance SharePointContext avec la requête HTTP spécifiée.
    ''' </summary>
    ''' <param name="httpRequest">Requête HTTP.</param>
    ''' <returns>Instance SharePointContext. Retourne <c>Nothing</c> en cas d'erreur.</returns>
    Public Function CreateSharePointContext(httpRequest As HttpRequestBase) As SharePointContext
        If httpRequest Is Nothing Then
            Throw New ArgumentNullException("httpRequest")
        End If

        ' SPHostUrl
        Dim spHostUrl As Uri = SharePointContext.GetSPHostUrl(httpRequest)
        If spHostUrl Is Nothing Then
            Return Nothing
        End If

        ' SPAppWebUrl
        Dim spAppWebUrlString As String = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString(SharePointContext.SPAppWebUrlKey))
        Dim spAppWebUrl As Uri = Nothing
        If Not Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, spAppWebUrl) OrElse
           Not (spAppWebUrl.Scheme = Uri.UriSchemeHttp OrElse spAppWebUrl.Scheme = Uri.UriSchemeHttps) Then
            spAppWebUrl = Nothing
        End If

        ' SPLanguage
        Dim spLanguage As String = httpRequest.QueryString(SharePointContext.SPLanguageKey)
        If String.IsNullOrEmpty(spLanguage) Then
            Return Nothing
        End If

        ' SPClientTag
        Dim spClientTag As String = httpRequest.QueryString(SharePointContext.SPClientTagKey)
        If String.IsNullOrEmpty(spClientTag) Then
            Return Nothing
        End If

        ' SPProductNumber
        Dim spProductNumber As String = httpRequest.QueryString(SharePointContext.SPProductNumberKey)
        If String.IsNullOrEmpty(spProductNumber) Then
            Return Nothing
        End If

        Return CreateSharePointContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, httpRequest)
    End Function

    ''' <summary>
    ''' Crée une instance SharePointContext avec la requête HTTP spécifiée.
    ''' </summary>
    ''' <param name="httpRequest">Requête HTTP.</param>
    ''' <returns>Instance SharePointContext. Retourne <c>Nothing</c> en cas d'erreur.</returns>
    Public Function CreateSharePointContext(httpRequest As HttpRequest) As SharePointContext
        Return CreateSharePointContext(New HttpRequestWrapper(httpRequest))
    End Function

    ''' <summary>
    ''' Obtient une instance SharePointContext associée au contexte HTTP spécifié.
    ''' </summary>
    ''' <param name="httpContext">Contexte HTTP.</param>
    ''' <returns>Instance SharePointContext. Retourne <c>Nothing</c> si aucune instance n'est trouvée et qu'une nouvelle instance ne peut pas être créée.</returns>
    Public Function GetSharePointContext(httpContext As HttpContextBase) As SharePointContext
        If httpContext Is Nothing Then
            Throw New ArgumentNullException("httpContext")
        End If

        Dim spHostUrl As Uri = SharePointContext.GetSPHostUrl(httpContext.Request)
        If spHostUrl Is Nothing Then
            Return Nothing
        End If

        Dim spContext As SharePointContext = LoadSharePointContext(httpContext)

        If spContext Is Nothing Or Not ValidateSharePointContext(spContext, httpContext) Then
            spContext = CreateSharePointContext(httpContext.Request)

            If spContext IsNot Nothing Then
                SaveSharePointContext(spContext, httpContext)
            End If
        End If

        Return spContext
    End Function

    ''' <summary>
    ''' Obtient une instance SharePointContext associée au contexte HTTP spécifié.
    ''' </summary>
    ''' <param name="httpContext">Contexte HTTP.</param>
    ''' <returns>Instance SharePointContext. Retourne <c>Nothing</c> si aucune instance n'est trouvée et qu'une nouvelle instance ne peut pas être créée.</returns>
    Public Function GetSharePointContext(httpContext As HttpContext) As SharePointContext
        Return GetSharePointContext(New HttpContextWrapper(httpContext))
    End Function

    ''' <summary>
    ''' Crée une instance SharePointContext.
    ''' </summary>
    ''' <param name="spHostUrl">URL de l'hôte SharePoint.</param>
    ''' <param name="spAppWebUrl">URL de l'application Web SharePoint.</param>
    ''' <param name="spLanguage">Langue SharePoint.</param>
    ''' <param name="spClientTag">Balise cliente SharePoint.</param>
    ''' <param name="spProductNumber">Numéro de produit SharePoint.</param>
    ''' <param name="httpRequest">Requête HTTP.</param>
    ''' <returns>Instance SharePointContext. Retourne <c>Nothing</c> en cas d'erreur.</returns>
    Protected MustOverride Function CreateSharePointContext(spHostUrl As Uri, spAppWebUrl As Uri, spLanguage As String, spClientTag As String, spProductNumber As String, httpRequest As HttpRequestBase) As SharePointContext

    ''' <summary>
    ''' Valide si l'objet SharePointContext donné peut être utilisé avec le contexte HTTP spécifié.
    ''' </summary>
    ''' <param name="spContext">SharePointContext.</param>
    ''' <param name="httpContext">Contexte HTTP.</param>
    ''' <returns>True si l'objet SharePointContext donné peut être utilisé avec le contexte HTTP spécifié.</returns>
    Protected MustOverride Function ValidateSharePointContext(spContext As SharePointContext, httpContext As HttpContextBase) As Boolean

    ''' <summary>
    ''' Charge l'instance SharePointContext associée au contexte HTTP spécifié.
    ''' </summary>
    ''' <param name="httpContext">Contexte HTTP.</param>
    ''' <returns>Instance SharePointContext. Retourne <c>Nothing</c> si aucune instance n'est trouvée.</returns>
    Protected MustOverride Function LoadSharePointContext(httpContext As HttpContextBase) As SharePointContext

    ''' <summary>
    ''' Enregistre l'instance SharePointContext spécifiée associée au contexte HTTP spécifié.
    ''' <c>Nothing</c> est acceptée pour effacer l'instance SharePointContext associée au contexte HTTP.
    ''' </summary>
    ''' <param name="spContext">Instance SharePointContext à enregistrer, ou <c>Nothing</c>.</param>
    ''' <param name="httpContext">Contexte HTTP.</param>
    Protected MustOverride Sub SaveSharePointContext(spContext As SharePointContext, httpContext As HttpContextBase)
End Class

#Region "ACS"

''' <summary>
''' Encapsule toutes les informations de SharePoint en mode ACS.
''' </summary>
Public Class SharePointAcsContext
    Inherits SharePointContext
    Private ReadOnly m_contextToken As String
    Private ReadOnly m_contextTokenObj As SharePointContextToken

    ''' <summary>
    ''' Jeton de contexte.
    ''' </summary>
    Public ReadOnly Property ContextToken() As String
        Get
            Return If(Me.m_contextTokenObj.ValidTo > DateTime.UtcNow, Me.m_contextToken, Nothing)
        End Get
    End Property

    ''' <summary>
    ''' Demande « CacheKey » du jeton de contexte.
    ''' </summary>
    Public ReadOnly Property CacheKey() As String
        Get
            Return If(Me.m_contextTokenObj.ValidTo > DateTime.UtcNow, Me.m_contextTokenObj.CacheKey, Nothing)
        End Get
    End Property

    ''' <summary>
    ''' Demande « refreshtoken » du jeton de contexte.
    ''' </summary>
    Public ReadOnly Property RefreshToken() As String
        Get
            Return If(Me.m_contextTokenObj.ValidTo > DateTime.UtcNow, Me.m_contextTokenObj.RefreshToken, Nothing)
        End Get
    End Property

    Public Overrides ReadOnly Property UserAccessTokenForSPHost() As String
        Get
            Return GetAccessTokenString(Me.m_userAccessTokenForSPHost, Function() TokenHelper.GetAccessToken(Me.m_contextTokenObj, Me.SPHostUrl.Authority))
        End Get
    End Property

    Public Overrides ReadOnly Property UserAccessTokenForSPAppWeb() As String
        Get
            If Me.SPAppWebUrl Is Nothing Then
                Return Nothing
            End If

            Return GetAccessTokenString(Me.m_userAccessTokenForSPAppWeb, Function() TokenHelper.GetAccessToken(Me.m_contextTokenObj, Me.SPAppWebUrl.Authority))
        End Get
    End Property

    Public Overrides ReadOnly Property AppOnlyAccessTokenForSPHost() As String
        Get
            Return GetAccessTokenString(Me.m_appOnlyAccessTokenForSPHost, Function() TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, Me.SPHostUrl.Authority, TokenHelper.GetRealmFromTargetUrl(Me.SPHostUrl)))
        End Get
    End Property

    Public Overrides ReadOnly Property AppOnlyAccessTokenForSPAppWeb() As String
        Get
            If Me.SPAppWebUrl Is Nothing Then
                Return Nothing
            End If

            Return GetAccessTokenString(Me.m_appOnlyAccessTokenForSPAppWeb, Function() TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, Me.SPAppWebUrl.Authority, TokenHelper.GetRealmFromTargetUrl(Me.SPAppWebUrl)))
        End Get
    End Property

    Public Sub New(spHostUrl As Uri, spAppWebUrl As Uri, spLanguage As String, spClientTag As String, spProductNumber As String, contextToken As String, contextTokenObj As SharePointContextToken)
        MyBase.New(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        If String.IsNullOrEmpty(contextToken) Then
            Throw New ArgumentNullException("contextToken")
        End If

        If contextTokenObj Is Nothing Then
            Throw New ArgumentNullException("contextTokenObj")
        End If

        Me.m_contextToken = contextToken
        Me.m_contextTokenObj = contextTokenObj
    End Sub

    ''' <summary>
    ''' Garantit que le jeton d'accès est valide et le retourne.
    ''' </summary>
    ''' <param name="accessToken">Jeton d'accès à vérifier.</param>
    ''' <param name="tokenRenewalHandler">Gestionnaire de renouvellement du jeton.</param>
    ''' <returns>Chaîne du jeton d'accès.</returns>
    Private Shared Function GetAccessTokenString(ByRef accessToken As Tuple(Of String, DateTime), tokenRenewalHandler As Func(Of OAuth2AccessTokenResponse)) As String
        RenewAccessTokenIfNeeded(accessToken, tokenRenewalHandler)

        Return If(IsAccessTokenValid(accessToken), accessToken.Item1, Nothing)
    End Function

    ''' <summary>
    ''' Renouvelle le jeton d'accès s'il n'est pas valide.
    ''' </summary>
    ''' <param name="accessToken">Jeton d'accès à renouveler.</param>
    ''' <param name="tokenRenewalHandler">Gestionnaire de renouvellement du jeton.</param>
    Private Shared Sub RenewAccessTokenIfNeeded(ByRef accessToken As Tuple(Of String, DateTime), tokenRenewalHandler As Func(Of OAuth2AccessTokenResponse))
        If IsAccessTokenValid(accessToken) Then
            Return
        End If

        Try
            Dim oAuth2AccessTokenResponse As OAuth2AccessTokenResponse = tokenRenewalHandler()

            Dim expiresOn As DateTime = oAuth2AccessTokenResponse.ExpiresOn

            If (expiresOn - oAuth2AccessTokenResponse.NotBefore) > AccessTokenLifetimeTolerance Then
                ' Entraîne un renouvellement du jeton d'accès légèrement plus tôt que la date d'expiration fixée
                ' de telle sorte que les appels à SharePoint associés aient assez de temps pour se terminer avec succès.
                expiresOn -= AccessTokenLifetimeTolerance
            End If

            accessToken = Tuple.Create(oAuth2AccessTokenResponse.AccessToken, expiresOn)
        Catch ex As WebException
        End Try
    End Sub
End Class

''' <summary>
''' Fournisseur par défaut pour SharePointAcsContext.
''' </summary>
Public Class SharePointAcsContextProvider
    Inherits SharePointContextProvider
    Private Const SPContextKey As String = "SPContext"
    Private Const SPCacheKeyKey As String = "SPCacheKey"

    Protected Overrides Function CreateSharePointContext(spHostUrl As Uri, spAppWebUrl As Uri, spLanguage As String, spClientTag As String, spProductNumber As String, httpRequest As HttpRequestBase) As SharePointContext
        Dim contextTokenString As String = TokenHelper.GetContextTokenFromRequest(httpRequest)
        If String.IsNullOrEmpty(contextTokenString) Then
            Return Nothing
        End If

        Dim contextToken As SharePointContextToken = Nothing
        Try
            contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, httpRequest.Url.Authority)
        Catch ex As WebException
            Return Nothing
        Catch ex As AudienceUriValidationFailedException
            Return Nothing
        End Try

        Return New SharePointAcsContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, contextTokenString, contextToken)
    End Function

    Protected Overrides Function ValidateSharePointContext(spContext As SharePointContext, httpContext As HttpContextBase) As Boolean
        Dim spAcsContext As SharePointAcsContext = TryCast(spContext, SharePointAcsContext)

        If spAcsContext IsNot Nothing Then
            Dim spHostUrl As Uri = SharePointContext.GetSPHostUrl(httpContext.Request)
            Dim contextToken As String = TokenHelper.GetContextTokenFromRequest(httpContext.Request)
            Dim spCacheKeyCookie As HttpCookie = httpContext.Request.Cookies(SPCacheKeyKey)
            Dim spCacheKey As String = If(spCacheKeyCookie IsNot Nothing, spCacheKeyCookie.Value, Nothing)

            Return spHostUrl = spAcsContext.SPHostUrl AndAlso
                   Not String.IsNullOrEmpty(spAcsContext.CacheKey) AndAlso
                   spCacheKey = spAcsContext.CacheKey AndAlso
                   Not String.IsNullOrEmpty(spAcsContext.ContextToken) AndAlso
                   (String.IsNullOrEmpty(contextToken) OrElse contextToken = spAcsContext.ContextToken)
        End If

        Return False
    End Function

    Protected Overrides Function LoadSharePointContext(httpContext As HttpContextBase) As SharePointContext
        Return TryCast(httpContext.Session(SPContextKey), SharePointAcsContext)
    End Function

    Protected Overrides Sub SaveSharePointContext(spContext As SharePointContext, httpContext As HttpContextBase)
        Dim spAcsContext As SharePointAcsContext = TryCast(spContext, SharePointAcsContext)

        If spAcsContext IsNot Nothing Then
            Dim spCacheKeyCookie As New HttpCookie(SPCacheKeyKey) With
            {
                .Value = spAcsContext.CacheKey,
                .Secure = True,
                .HttpOnly = True
            }

            httpContext.Response.AppendCookie(spCacheKeyCookie)
        End If

        httpContext.Session(SPContextKey) = spAcsContext
    End Sub
End Class

#End Region

#Region "HighTrust"

''' <summary>
''' Encapsule toutes les informations de SharePoint en mode HighTrust.
''' </summary>
Public Class SharePointHighTrustContext
    Inherits SharePointContext
    Private ReadOnly m_logonUserIdentity As WindowsIdentity

    ''' <summary>
    ''' Identité Windows de l'utilisateur actuel.
    ''' </summary>
    Public ReadOnly Property LogonUserIdentity() As WindowsIdentity
        Get
            Return Me.m_logonUserIdentity
        End Get
    End Property

    Public Overrides ReadOnly Property UserAccessTokenForSPHost() As String
        Get
            Return GetAccessTokenString(Me.m_userAccessTokenForSPHost, Function() TokenHelper.GetS2SAccessTokenWithWindowsIdentity(Me.SPHostUrl, Me.LogonUserIdentity))
        End Get
    End Property

    Public Overrides ReadOnly Property UserAccessTokenForSPAppWeb() As String
        Get
            If Me.SPAppWebUrl Is Nothing Then
                Return Nothing
            End If

            Return GetAccessTokenString(Me.m_userAccessTokenForSPAppWeb, Function() TokenHelper.GetS2SAccessTokenWithWindowsIdentity(Me.SPAppWebUrl, Me.LogonUserIdentity))
        End Get
    End Property

    Public Overrides ReadOnly Property AppOnlyAccessTokenForSPHost() As String
        Get
            Return GetAccessTokenString(Me.m_appOnlyAccessTokenForSPHost, Function() TokenHelper.GetS2SAccessTokenWithWindowsIdentity(Me.SPHostUrl, Nothing))
        End Get
    End Property

    Public Overrides ReadOnly Property AppOnlyAccessTokenForSPAppWeb() As String
        Get
            If Me.SPAppWebUrl Is Nothing Then
                Return Nothing
            End If

            Return GetAccessTokenString(Me.m_appOnlyAccessTokenForSPAppWeb, Function() TokenHelper.GetS2SAccessTokenWithWindowsIdentity(Me.SPAppWebUrl, Nothing))
        End Get
    End Property

    Public Sub New(spHostUrl As Uri, spAppWebUrl As Uri, spLanguage As String, spClientTag As String, spProductNumber As String, logonUserIdentity As WindowsIdentity)
        MyBase.New(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        If logonUserIdentity Is Nothing Then
            Throw New ArgumentNullException("logonUserIdentity")
        End If

        Me.m_logonUserIdentity = logonUserIdentity
    End Sub

    ''' <summary>
    ''' Garantit que le jeton d'accès est valide et le retourne.
    ''' </summary>
    ''' <param name="accessToken">Jeton d'accès à vérifier.</param>
    ''' <param name="tokenRenewalHandler">Gestionnaire de renouvellement du jeton.</param>
    ''' <returns>Chaîne du jeton d'accès.</returns>
    Private Shared Function GetAccessTokenString(ByRef accessToken As Tuple(Of String, DateTime), tokenRenewalHandler As Func(Of String)) As String
        RenewAccessTokenIfNeeded(accessToken, tokenRenewalHandler)

        Return If(IsAccessTokenValid(accessToken), accessToken.Item1, Nothing)
    End Function

    ''' <summary>
    ''' Renouvelle le jeton d'accès s'il n'est pas valide.
    ''' </summary>
    ''' <param name="accessToken">Jeton d'accès à renouveler.</param>
    ''' <param name="tokenRenewalHandler">Gestionnaire de renouvellement du jeton.</param>
    Private Shared Sub RenewAccessTokenIfNeeded(ByRef accessToken As Tuple(Of String, DateTime), tokenRenewalHandler As Func(Of String))
        If IsAccessTokenValid(accessToken) Then
            Return
        End If

        Dim expiresOn As DateTime = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime)

        If TokenHelper.HighTrustAccessTokenLifetime > AccessTokenLifetimeTolerance Then
            ' Entraîne un renouvellement du jeton d'accès légèrement plus tôt que la date d'expiration fixée
            ' de telle sorte que les appels à SharePoint associés aient assez de temps pour se terminer avec succès.
            expiresOn -= AccessTokenLifetimeTolerance
        End If

        accessToken = Tuple.Create(tokenRenewalHandler(), expiresOn)
    End Sub
End Class

''' <summary>
''' Fournisseur par défaut pour SharePointHighTrustContext.
''' </summary>
Public Class SharePointHighTrustContextProvider
    Inherits SharePointContextProvider
    Private Const SPContextKey As String = "SPContext"

    Protected Overrides Function CreateSharePointContext(spHostUrl As Uri, spAppWebUrl As Uri, spLanguage As String, spClientTag As String, spProductNumber As String, httpRequest As HttpRequestBase) As SharePointContext
        Dim logonUserIdentity As WindowsIdentity = httpRequest.LogonUserIdentity
        If logonUserIdentity Is Nothing Or Not logonUserIdentity.IsAuthenticated Or logonUserIdentity.IsGuest Or logonUserIdentity.User Is Nothing Then
            Return Nothing
        End If

        Return New SharePointHighTrustContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, logonUserIdentity)
    End Function

    Protected Overrides Function ValidateSharePointContext(spContext As SharePointContext, httpContext As HttpContextBase) As Boolean
        Dim spHighTrustContext As SharePointHighTrustContext = TryCast(spContext, SharePointHighTrustContext)

        If spHighTrustContext IsNot Nothing Then
            Dim spHostUrl As Uri = SharePointContext.GetSPHostUrl(httpContext.Request)
            Dim logonUserIdentity As WindowsIdentity = httpContext.Request.LogonUserIdentity

            Return spHostUrl = spHighTrustContext.SPHostUrl AndAlso
                   logonUserIdentity IsNot Nothing AndAlso
                   logonUserIdentity.IsAuthenticated AndAlso
                   Not logonUserIdentity.IsGuest AndAlso
                   logonUserIdentity.User = spHighTrustContext.LogonUserIdentity.User
        End If

        Return False
    End Function

    Protected Overrides Function LoadSharePointContext(httpContext As HttpContextBase) As SharePointContext
        Return TryCast(httpContext.Session(SPContextKey), SharePointHighTrustContext)
    End Function

    Protected Overrides Sub SaveSharePointContext(spContext As SharePointContext, httpContext As HttpContextBase)
        httpContext.Session(SPContextKey) = TryCast(spContext, SharePointHighTrustContext)
    End Sub
End Class

#End Region
