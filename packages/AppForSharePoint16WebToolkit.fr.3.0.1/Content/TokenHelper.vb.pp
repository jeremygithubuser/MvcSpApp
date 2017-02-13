Imports Microsoft.IdentityModel
Imports Microsoft.IdentityModel.S2S.Protocols.OAuth2
Imports Microsoft.IdentityModel.S2S.Tokens
Imports Microsoft.SharePoint.Client
Imports Microsoft.SharePoint.Client.EventReceivers
Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.IdentityModel.Selectors
Imports System.IdentityModel.Tokens
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Principal
Imports System.ServiceModel
Imports System.Text
Imports System.Web
Imports System.Web.Configuration
Imports System.Web.Script.Serialization
Imports AudienceRestriction = Microsoft.IdentityModel.Tokens.AudienceRestriction
Imports AudienceUriValidationFailedException = Microsoft.IdentityModel.Tokens.AudienceUriValidationFailedException
Imports SecurityTokenHandlerConfiguration = Microsoft.IdentityModel.Tokens.SecurityTokenHandlerConfiguration
Imports X509SigningCredentials = Microsoft.IdentityModel.SecurityTokenService.X509SigningCredentials

Public NotInheritable Class TokenHelper

#Region "champs publics"

    ''' <summary>
    ''' Principal SharePoint.
    ''' </summary>
    Public Const SharePointPrincipal As String = "00000003-0000-0ff1-ce00-000000000000"

    ''' <summary>
    ''' Durée de vie du jeton d'accès HighTrust, 12 heures.
    ''' </summary>
    Public Shared ReadOnly HighTrustAccessTokenLifetime As TimeSpan = TimeSpan.FromHours(12.0)

#End Region

#Region "méthodes publiques"

    ''' <summary>
    ''' Extrait la chaîne du jeton de contexte de la demande spécifiée en recherchant des noms de paramètre connus dans les 
    ''' paramètres de formulaire POSTed et la querystring. Retourne Nothing si aucun jeton de contexte n'est trouvé.
    ''' </summary>
    ''' <param name="request">HttpRequest dans laquelle rechercher un jeton de contexte</param>
    ''' <returns>Chaîne du jeton de contexte</returns>
    Public Shared Function GetContextTokenFromRequest(request As HttpRequest) As String
        Return GetContextTokenFromRequest(New HttpRequestWrapper(request))
    End Function

    ''' <summary>
    ''' Extrait la chaîne du jeton de contexte de la demande spécifiée en recherchant des noms de paramètre connus dans les 
    ''' paramètres de formulaire POSTed et la querystring. Retourne Nothing si aucun jeton de contexte n'est trouvé.
    ''' </summary>
    ''' <param name="request">HttpRequest dans laquelle rechercher un jeton de contexte</param>
    ''' <returns>Chaîne du jeton de contexte</returns>
    Public Shared Function GetContextTokenFromRequest(request As HttpRequestBase) As String
        Dim paramNames As String() = {"AppContext", "AppContextToken", "AccessToken", "SPAppToken"}
        For Each paramName As String In paramNames
            If Not String.IsNullOrEmpty(request.Form(paramName)) Then
                Return request.Form(paramName)
            End If
            If Not String.IsNullOrEmpty(request.QueryString(paramName)) Then
                Return request.QueryString(paramName)
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Valider le fait qu'une chaîne de jeton de contexte spécifiée est destinée à cette application en fonction des paramètres 
    ''' spécifiés dans web.config. Les paramètres utilisés depuis web.config pour la validation comprennent ClientId, 
    ''' HostedAppHostNameOverride, HostedAppHostName, ClientSecret et Realm (s'ils sont spécifiés). Si HostedAppHostNameOverride est présent,
    ''' il sera utilisé pour la validation. Sinon, si <paramref name="appHostName"/> n'est pas
    ''' Si la valeur est Nothing, il est utilisé pour la validation au lieu du HostedAppHostName de web.config. Si le jeton n'est pas valide, une 
    ''' exception est levée. Si le jeton est valide, l'URL de métadonnées STS statique de TokenHelper est mise à jour en fonction du contenu du jeton
    ''' et un JsonWebSecurityToken reposant sur le jeton de contexte est retourné.
    ''' </summary>
    ''' <param name="contextTokenString">Jeton de contexte à valider</param>
    ''' <param name="appHostName">L'autorité de L'URL, composée du nom d'hôte du système de nom de domaine (DNS) ou de l'adresse IP et du numéro de port, à utiliser pour la validation de l'audience du jeton.
    ''' Si la valeur est Nothing, le paramètre web.config HostedAppHostName est utilisé à la place. S'il est présent, le paramètre web.config HostedAppHostNameOverride sera utilisé
    ''' pour validation au lieu de <paramref name="appHostName"/> .</param>
    ''' <returns>JsonWebSecurityToken reposant sur le jeton de contexte.</returns>
    Public Shared Function ReadAndValidateContextToken(contextTokenString As String, Optional appHostName As String = Nothing) As SharePointContextToken
        Dim tokenHandler As JsonWebSecurityTokenHandler = CreateJsonWebSecurityTokenHandler()
        Dim securityToken As SecurityToken = tokenHandler.ReadToken(contextTokenString)
        Dim jsonToken As JsonWebSecurityToken = TryCast(securityToken, JsonWebSecurityToken)
        Dim token As SharePointContextToken = SharePointContextToken.Create(jsonToken)

        Dim stsAuthority As String = (New Uri(token.SecurityTokenServiceUri)).Authority
        Dim firstDot As Integer = stsAuthority.IndexOf("."c)

        GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot)
        AcsHostUrl = stsAuthority.Substring(firstDot + 1)

        tokenHandler.ValidateToken(jsonToken)

        Dim acceptableAudiences As String()
        If Not [String].IsNullOrEmpty(HostedAppHostNameOverride) Then
            acceptableAudiences = HostedAppHostNameOverride.Split(";"c)
        ElseIf appHostName Is Nothing Then
            acceptableAudiences = {HostedAppHostName}
        Else
            acceptableAudiences = {appHostName}
        End If

        Dim validationSuccessful As Boolean
        Dim definedRealm As String = If(Realm, token.Realm)
        For Each audience In acceptableAudiences
            Dim principal As String = GetFormattedPrincipal(ClientId, audience, definedRealm)
            If StringComparer.OrdinalIgnoreCase.Equals(token.Audience, principal) Then
                validationSuccessful = True
                Exit For
            End If
        Next

        If Not validationSuccessful Then
            Throw New AudienceUriValidationFailedException([String].Format(CultureInfo.CurrentCulture, """{0}"" is not the intended audience ""{1}""", [String].Join(";", acceptableAudiences), token.Audience))
        End If

        Return token
    End Function

    ''' <summary>
    ''' Extrait un jeton d'accès d'ACS afin d'appeler la source du jeton de contexte spécifié sur le 
    ''' targetHost spécifié. Le targetHost doit être inscrit pour le principal qui a envoyé le jeton de contexte.
    ''' </summary>
    ''' <param name="contextToken">Jeton de contexte émis par l'audience de jeton d'accès ciblée</param>
    ''' <param name="targetHost">Autorité de l'URL du principal cible</param>
    ''' <returns>Un jeton d'accès avec une audience correspondant à la source du jeton de contexte</returns>
    Public Shared Function GetAccessToken(contextToken As SharePointContextToken, targetHost As String) As OAuth2AccessTokenResponse

        Dim targetPrincipalName As String = contextToken.TargetPrincipalName

        ' Extraire le refreshtoken du jeton de contexte
        Dim refreshToken As String = contextToken.RefreshToken

        If [String].IsNullOrEmpty(refreshToken) Then
            Return Nothing
        End If

        Dim targetRealm As String = If(Realm, contextToken.Realm)

        Return GetAccessToken(refreshToken, targetPrincipalName, targetHost, targetRealm)
    End Function

    ''' <summary>
    ''' Utilise le code d'autorisation spécifié pour extraire un jeton d'accès d'ACS afin d'appeler le principal spécifié 
    ''' sur le targetHost spécifié. Le targetHost doit être inscrit pour le principal cible.  Si le domaine spécifié est 
    ''' Si la valeur est Nothing, le paramètre "Realm" de web.config est utilisé à la place.
    ''' </summary>
    ''' <param name="authorizationCode">Code d'autorisation à échanger contre le jeton d'accès</param>
    ''' <param name="targetPrincipalName">Nom du principal cible pour extraire un jeton d'accès pour</param>
    ''' <param name="targetHost">Autorité de l'URL du principal cible</param>
    ''' <param name="targetRealm">Domaine à utiliser pour les ID de nom et audience du jeton d'accès</param>
    ''' <param name="redirectUri">URI de redirection enregistré pour cette application</param>
    ''' <returns>Un jeton d'accès avec une audience du principal cible</returns>
    Public Shared Function GetAccessToken(authorizationCode As String, targetPrincipalName As String, targetHost As String, targetRealm As String, redirectUri As Uri) As OAuth2AccessTokenResponse

        If targetRealm Is Nothing Then
            targetRealm = Realm
        End If

        Dim resource As String = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm)
        Dim formattedClientId As String = GetFormattedPrincipal(ClientId, Nothing, targetRealm)

        ' Créer une demande de jeton. RedirectUri est égal à Nothing ici.  Échec si l'URI de redirection est enregistré
        Dim oauth2Request As OAuth2AccessTokenRequest = OAuth2MessageFactory.CreateAccessTokenRequestWithAuthorizationCode(formattedClientId, ClientSecret, authorizationCode, redirectUri, resource)

        ' Obtenir un jeton
        Dim client As New OAuth2S2SClient()
        Dim oauth2Response As OAuth2AccessTokenResponse
        Try
            oauth2Response = TryCast(client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request), OAuth2AccessTokenResponse)
        Catch ex As WebException
            Using sr As New StreamReader(ex.Response.GetResponseStream())
                Dim responseText As String = sr.ReadToEnd()
                Throw New WebException(ex.Message + " - " + responseText, ex)
            End Using
        End Try

        Return oauth2Response
    End Function

    ''' <summary>
    ''' Utilise le jeton d'actualisation spécifié pour extraire un jeton d'accès d'ACS afin d'appeler le principal spécifié 
    ''' sur le targetHost spécifié. Le targetHost doit être inscrit pour le principal cible.  Si le domaine spécifié est 
    ''' Si la valeur est Nothing, le paramètre "Realm" de web.config est utilisé à la place.
    ''' </summary>
    ''' <param name="refreshToken">Jeton d'actualisation à échanger contre le jeton d'accès</param>
    ''' <param name="targetPrincipalName">Nom du principal cible pour extraire un jeton d'accès pour</param>
    ''' <param name="targetHost">Autorité de l'URL du principal cible</param>
    ''' <param name="targetRealm">Domaine à utiliser pour les ID de nom et audience du jeton d'accès</param>
    ''' <returns>Un jeton d'accès avec une audience du principal cible</returns>
    Public Shared Function GetAccessToken(refreshToken As String, targetPrincipalName As String, targetHost As String, targetRealm As String) As OAuth2AccessTokenResponse

        If targetRealm Is Nothing Then
            targetRealm = Realm
        End If

        Dim resource As String = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm)
        Dim formattedClientId As String = GetFormattedPrincipal(ClientId, Nothing, targetRealm)

        Dim oauth2Request As OAuth2AccessTokenRequest = OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(formattedClientId, ClientSecret, refreshToken, resource)

        ' Obtenir un jeton
        Dim client As New OAuth2S2SClient()
        Dim oauth2Response As OAuth2AccessTokenResponse
        Try
            oauth2Response = TryCast(client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request), OAuth2AccessTokenResponse)
        Catch wex As WebException
            Using sr As New StreamReader(wex.Response.GetResponseStream())
                Dim responseText As String = sr.ReadToEnd()
                Throw New WebException(wex.Message + " - " & responseText, wex)
            End Using
        End Try

        Return oauth2Response
    End Function

    ''' <summary>
    ''' Extrait d'ACS un jeton d'accès pour l'application uniquement afin d'appeler le principal spécifié
    ''' sur le targetHost spécifié. Le targetHost doit être inscrit pour le principal cible.  Si le domaine spécifié est 
    ''' Si la valeur est Nothing, le paramètre "Realm" de web.config est utilisé à la place.
    ''' </summary>
    ''' <param name="targetPrincipalName">Nom du principal cible pour extraire un jeton d'accès pour</param>
    ''' <param name="targetHost">Autorité de l'URL du principal cible</param>
    ''' <param name="targetRealm">Domaine à utiliser pour les ID de nom et audience du jeton d'accès</param>
    ''' <returns>Un jeton d'accès avec une audience du principal cible</returns>
    Public Shared Function GetAppOnlyAccessToken(targetPrincipalName As String, targetHost As String, targetRealm As String) As OAuth2AccessTokenResponse

        If targetRealm Is Nothing Then
            targetRealm = Realm
        End If

        Dim resource As String = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm)
        Dim formattedClientId As String = GetFormattedPrincipal(ClientId, HostedAppHostName, targetRealm)

        Dim oauth2Request As OAuth2AccessTokenRequest = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(formattedClientId, ClientSecret, resource)
        oauth2Request.Resource = resource

        ' Obtenir un jeton
        Dim client As New OAuth2S2SClient()

        Dim oauth2Response As OAuth2AccessTokenResponse
        Try
            oauth2Response = TryCast(client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request), OAuth2AccessTokenResponse)
        Catch wex As WebException
            Using sr As New StreamReader(wex.Response.GetResponseStream())
                Dim responseText As String = sr.ReadToEnd()
                Throw New WebException(wex.Message + " - " & responseText, wex)
            End Using
        End Try

        Return oauth2Response
    End Function

    ''' <summary>
    ''' Crée un contexte client en fonction des propriétés d'un récepteur d'événements distant
    ''' </summary>
    ''' <param name="properties">Propriétés d'un récepteur d'événements distant</param>
    ''' <returns>Un ClientContext prêt à appeler le site Web dont provient l'événement</returns>
    Public Shared Function CreateRemoteEventReceiverClientContext(properties As SPRemoteEventProperties) As ClientContext
        Dim sharepointUrl As Uri
        If properties.ListEventProperties IsNot Nothing Then
            sharepointUrl = New Uri(properties.ListEventProperties.WebUrl)
        ElseIf properties.ItemEventProperties IsNot Nothing Then
            sharepointUrl = New Uri(properties.ItemEventProperties.WebUrl)
        ElseIf properties.WebEventProperties IsNot Nothing Then
            sharepointUrl = New Uri(properties.WebEventProperties.FullUrl)
        Else
            Return Nothing
        End If

        If IsHighTrustApp() Then
            Return GetS2SClientContextWithWindowsIdentity(sharepointUrl, Nothing)
        End If

        Return CreateAcsClientContextForUrl(properties, sharepointUrl)

    End Function

    ''' <summary>
    ''' Crée un contexte client en fonction des propriétés d'un événement d'application
    ''' </summary>
    ''' <param name="properties">Propriétés d'un événement d'application</param>
    ''' <param name="useAppWeb">True pour cibler le site Web de l'application, false pour cibler le site Web hôte</param>
    ''' <returns>Un ClientContext prêt à appeler le site Web parent ou celui de l'application</returns>
    Public Shared Function CreateAppEventClientContext(properties As SPRemoteEventProperties, useAppWeb As Boolean) As ClientContext
        If properties.AppEventProperties Is Nothing Then
            Return Nothing
        End If

        Dim sharepointUrl As Uri = If(useAppWeb, properties.AppEventProperties.AppWebFullUrl, properties.AppEventProperties.HostWebFullUrl)
        If IsHighTrustApp() Then
            Return GetS2SClientContextWithWindowsIdentity(sharepointUrl, Nothing)
        End If

        Return CreateAcsClientContextForUrl(properties, sharepointUrl)
    End Function

    ''' <summary>
    ''' Extrait un jeton d'accès d'ACS à l'aide du code d'autorisation spécifié et utilise ce jeton d'accès pour 
    ''' créer un contexte client
    ''' </summary>
    ''' <param name="targetUrl">URL du site SharePoint cible</param>
    ''' <param name="authorizationCode">Code d'autorisation à utiliser lors de l'extraction du jeton d'accès d'ACS</param>
    ''' <param name="redirectUri">URI de redirection enregistré pour cette application</param>
    ''' <returns>Un ClientContext prêt à appeler targetUrl avec un jeton d'accès valide</returns>
    Public Shared Function GetClientContextWithAuthorizationCode(targetUrl As String, authorizationCode As String, redirectUri As Uri) As ClientContext
        Return GetClientContextWithAuthorizationCode(targetUrl, SharePointPrincipal, authorizationCode, GetRealmFromTargetUrl(New Uri(targetUrl)), redirectUri)
    End Function

    ''' <summary>
    ''' Extrait un jeton d'accès d'ACS à l'aide du code d'autorisation spécifié et utilise ce jeton d'accès pour 
    ''' créer un contexte client
    ''' </summary>
    ''' <param name="targetUrl">URL du site SharePoint cible</param>
    ''' <param name="targetPrincipalName">Nom du principal SharePoint cible</param>
    ''' <param name="authorizationCode">Code d'autorisation à utiliser lors de l'extraction du jeton d'accès d'ACS</param>
    ''' <param name="targetRealm">Domaine à utiliser pour les ID de nom et audience du jeton d'accès</param>
    ''' <param name="redirectUri">URI de redirection enregistré pour cette application</param>
    ''' <returns>Un ClientContext prêt à appeler targetUrl avec un jeton d'accès valide</returns>
    Public Shared Function GetClientContextWithAuthorizationCode(targetUrl As String, targetPrincipalName As String, authorizationCode As String, targetRealm As String, redirectUri As Uri) As ClientContext
        Dim targetUri As New Uri(targetUrl)

        Dim accessToken As String = GetAccessToken(authorizationCode, targetPrincipalName, targetUri.Authority, targetRealm, redirectUri).AccessToken

        Return GetClientContextWithAccessToken(targetUrl, accessToken)
    End Function

    ''' <summary>
    ''' Utilise le jeton d'accès spécifié pour créer un contexte client
    ''' </summary>
    ''' <param name="targetUrl">URL du site SharePoint cible</param>
    ''' <param name="accessToken">Jeton d'accès à utiliser lors de l'appel de la targetUrl spécifiée</param>
    ''' <returns>Un ClientContext prêt à appeler targetUrl avec le jeton d'accès spécifié</returns>
    Public Shared Function GetClientContextWithAccessToken(targetUrl As String, accessToken As String) As ClientContext
        Dim clientContext As New ClientContext(targetUrl)

        clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous
        clientContext.FormDigestHandlingEnabled = False

        AddHandler clientContext.ExecutingWebRequest, Sub(oSender As Object, webRequestEventArgs As WebRequestEventArgs)
                                                          webRequestEventArgs.WebRequestExecutor.RequestHeaders("Authorization") = "Bearer " & accessToken
                                                      End Sub
        Return clientContext
    End Function

    ''' <summary>
    ''' Extrait un jeton d'accès d'ACS à l'aide du jeton de contexte spécifié et utilise ce jeton d'accès pour créer
    ''' un contexte client
    ''' </summary>
    ''' <param name="targetUrl">URL du site SharePoint cible</param>
    ''' <param name="contextTokenString">Jeton de contexte reçu du site SharePoint cible</param>
    ''' <param name="appHostUrl">Autorité de l'URL de l'application hébergée.  Si la valeur est Nothing, la valeur de HostedAppHostName
    ''' de web.config sera utilisée à la place</param>
    ''' <returns>Un ClientContext prêt à appeler targetUrl avec un jeton d'accès valide</returns>
    Public Shared Function GetClientContextWithContextToken(targetUrl As String, contextTokenString As String, appHostUrl As String) As ClientContext
        Dim contextToken As SharePointContextToken = ReadAndValidateContextToken(contextTokenString, appHostUrl)

        Dim targetUri As New Uri(targetUrl)

        Dim accessToken As String = GetAccessToken(contextToken, targetUri.Authority).AccessToken

        Return GetClientContextWithAccessToken(targetUrl, accessToken)
    End Function

    ''' <summary>
    ''' Retourne l'URL SharePoint vers laquelle l'application doit rediriger le navigateur pour demander le consentement et rapporter
    ''' un code d'autorisation.
    ''' </summary>
    ''' <param name="contextUrl">URL absolue du site SharePoint</param>
    ''' <param name="scope">Autorisations délimitées par des espaces à demander auprès du site SharePoint au format abrégé 
    ''' (par exemple, "Web.Read Site.Write")</param>
    ''' <returns>URL de la page d'autorisation OAuth du site SharePoint</returns>
    Public Shared Function GetAuthorizationUrl(contextUrl As String, scope As String) As String
        Return String.Format("{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code", EnsureTrailingSlash(contextUrl), AuthorizationPage, ClientId, scope)
    End Function

    ''' <summary>
    ''' Retourne l'URL SharePoint vers laquelle l'application doit rediriger le navigateur pour demander le consentement et rapporter
    ''' un code d'autorisation.
    ''' </summary>
    ''' <param name="contextUrl">URL absolue du site SharePoint</param>
    ''' <param name="scope">Autorisations délimitées par des espaces à demander auprès du site SharePoint au format abrégé
    ''' (par exemple, "Web.Read Site.Write")</param>
    ''' <param name="redirectUri">URI vers lequel SharePoint doit rediriger le navigateur une fois le consentement 
    ''' donné</param>
    ''' <returns>URL de la page d'autorisation OAuth du site SharePoint</returns>
    Public Shared Function GetAuthorizationUrl(contextUrl As String, scope As String, redirectUri As String) As String
        Return String.Format("{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code&redirect_uri={4}", EnsureTrailingSlash(contextUrl), AuthorizationPage, ClientId, scope, redirectUri)
    End Function

    ''' <summary>
    ''' Retourne l'URL SharePoint vers laquelle l'application doit rediriger le navigateur pour demander un nouveau jeton de contexte.
    ''' </summary>
    ''' <param name="contextUrl">URL absolue du site SharePoint</param>
    ''' <param name="redirectUri">URI vers lequel SharePoint doit rediriger le navigateur avec un jeton de contexte</param>
    ''' <returns>URL de la page de redirection du jeton de contexte du site SharePoint</returns>
    Public Shared Function GetAppContextTokenRequestUrl(contextUrl As String, redirectUri As String) As String
        Return String.Format("{0}{1}?client_id={2}&redirect_uri={3}", EnsureTrailingSlash(contextUrl), RedirectPage, ClientId, redirectUri)
    End Function

    ''' <summary>
    ''' Extrait un jeton d'accès S2S signé par le certificat privé de l'application au nom de la 
    ''' WindowsIdentity spécifiée et destiné à SharePoint pour targetApplicationUri. Si aucun domaine n'est spécifié dans 
    ''' web.config, une demande d'authentification sera émise sur targetApplicationUri pour la détection.
    ''' </summary>
    ''' <param name="targetApplicationUri">URL du site SharePoint cible</param>
    ''' <param name="identity">Identité Windows de l'utilisateur au nom duquel le jeton d'accès est créé</param>
    ''' <returns>Un jeton d'accès avec une audience du principal cible</returns>
    Public Shared Function GetS2SAccessTokenWithWindowsIdentity(targetApplicationUri As Uri, identity As WindowsIdentity) As String
        Dim targetRealm As String = If(String.IsNullOrEmpty(Realm), GetRealmFromTargetUrl(targetApplicationUri), Realm)

        Dim claims As JsonWebTokenClaim() = If(identity IsNot Nothing, GetClaimsWithWindowsIdentity(identity), Nothing)

        Return GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, targetRealm, claims)
    End Function

    ''' <summary>
    ''' Extrait un contexte de client S2S avec un jeton d'accès signé par le certificat privé de l'application 
    ''' au nom de la WindowsIdentity spécifiée et destiné à l'application pour targetApplicationUri à l'aide du 
    ''' targetRealm. Si aucun domaine n'est spécifié dans web.config, une demande d'authentification sera émise sur le 
    ''' targetApplicationUri pour la détection.
    ''' </summary>
    ''' <param name="targetApplicationUri">URL du site SharePoint cible</param>
    ''' <param name="identity">Identité Windows de l'utilisateur au nom duquel le jeton d'accès est créé</param>
    ''' <returns>Un ClientContext utilisant un jeton d'accès avec une audience de l'application cible</returns>
    Public Shared Function GetS2SClientContextWithWindowsIdentity(targetApplicationUri As Uri, identity As WindowsIdentity) As ClientContext
        Dim targetRealm As String = If(String.IsNullOrEmpty(Realm), GetRealmFromTargetUrl(targetApplicationUri), Realm)

        Dim claims As JsonWebTokenClaim() = If(identity IsNot Nothing, GetClaimsWithWindowsIdentity(identity), Nothing)

        Dim accessToken As String = GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, targetRealm, claims)

        Return GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken)
    End Function

    ''' <summary>
    ''' Obtenir le domaine d'identification de SharePoint
    ''' </summary>
    ''' <param name="targetApplicationUri">URL du site SharePoint cible</param>
    ''' <returns> Représentation sous forme de chaîne du GUID de domaine</returns>
    Public Shared Function GetRealmFromTargetUrl(targetApplicationUri As Uri) As String
        Dim request As WebRequest = HttpWebRequest.Create(targetApplicationUri.ToString() & "/_vti_bin/client.svc")
        request.Headers.Add("Authorization: Bearer ")

        Try
            request.GetResponse().Close()
        Catch e As WebException
            If e.Response Is Nothing Then
                Return Nothing
            End If

            Dim bearerResponseHeader As String = e.Response.Headers("WWW-Authenticate")
            If String.IsNullOrEmpty(bearerResponseHeader) Then
                Return Nothing
            End If

            Const bearer As String = "Bearer realm="""
            Dim bearerIndex As Integer = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal)
            If bearerIndex < 0 Then
                Return Nothing
            End If

            Dim realmIndex As Integer = bearerIndex + bearer.Length

            If bearerResponseHeader.Length >= realmIndex + 36 Then
                Dim targetRealm As String = bearerResponseHeader.Substring(realmIndex, 36)

                Dim realmGuid As Guid

                If Guid.TryParse(targetRealm, realmGuid) Then
                    Return targetRealm
                End If
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Permet de déterminer s'il s'agit d'une application HighTrust.
    ''' </summary>
    ''' <returns>True s'il s'agit d'une application HighTrust.</returns>
    Public Shared Function IsHighTrustApp() As Boolean
        Return SigningCredentials IsNot Nothing
    End Function

    ''' <summary>
    ''' Garantit que l'URL spécifiée se termine par '/' si elle n'est ni Null ni vide.
    ''' </summary>
    ''' <param name="url">URL.</param>
    ''' <returns>URL se terminant par '/' si elle n'est ni Null ni vide.</returns>
    Public Shared Function EnsureTrailingSlash(url As String) As String
        If Not String.IsNullOrEmpty(url) AndAlso url(url.Length - 1) <> "/"c Then
            Return url + "/"
        End If

        Return url
    End Function

#End Region

#Region "champs privés"

    '
    ' Constantes de configuration
    '        

    Private Const AuthorizationPage As String = "_layouts/15/OAuthAuthorize.aspx"
    Private Const RedirectPage As String = "_layouts/15/AppRedirect.aspx"
    Private Const AcsPrincipalName As String = "00000001-0000-0000-c000-000000000000"
    Private Const AcsMetadataEndPointRelativeUrl As String = "metadata/json/1"
    Private Const S2SProtocol As String = "OAuth2"
    Private Const DelegationIssuance As String = "DelegationIssuance1.0"
    Private Const NameIdentifierClaimType As String = JsonWebTokenConstants.ReservedClaims.NameIdentifier
    Private Const TrustedForImpersonationClaimType As String = "trustedfordelegation"
    Private Const ActorTokenClaimType As String = JsonWebTokenConstants.ReservedClaims.ActorToken

    '
    ' Constantes d'environnement
    '

    Private Shared GlobalEndPointPrefix As String = "accounts"
    Private Shared AcsHostUrl As String = "accesscontrol.windows.net"

    '
    ' Configuration de l'application hébergée
    '
    Private Shared ReadOnly ClientId As String = If(String.IsNullOrEmpty(WebConfigurationManager.AppSettings.[Get]("ClientId")), WebConfigurationManager.AppSettings.[Get]("HostedAppName"), WebConfigurationManager.AppSettings.[Get]("ClientId"))

    Private Shared ReadOnly IssuerId As String = If(String.IsNullOrEmpty(WebConfigurationManager.AppSettings.[Get]("IssuerId")), ClientId, WebConfigurationManager.AppSettings.[Get]("IssuerId"))

    Private Shared ReadOnly HostedAppHostName As String = WebConfigurationManager.AppSettings.[Get]("HostedAppHostName")

    Private Shared ReadOnly HostedAppHostNameOverride As String = WebConfigurationManager.AppSettings.[Get]("HostedAppHostNameOverride")

    Private Shared ReadOnly ClientSecret As String = If(String.IsNullOrEmpty(WebConfigurationManager.AppSettings.[Get]("ClientSecret")), WebConfigurationManager.AppSettings.[Get]("HostedAppSigningKey"), WebConfigurationManager.AppSettings.[Get]("ClientSecret"))

    Private Shared ReadOnly SecondaryClientSecret As String = WebConfigurationManager.AppSettings.[Get]("SecondaryClientSecret")

    Private Shared ReadOnly Realm As String = WebConfigurationManager.AppSettings.[Get]("Realm")

    Private Shared ReadOnly ServiceNamespace As String = WebConfigurationManager.AppSettings.[Get]("Realm")

    Private Shared ReadOnly ClientSigningCertificatePath As String = WebConfigurationManager.AppSettings.[Get]("ClientSigningCertificatePath")

    Private Shared ReadOnly ClientSigningCertificatePassword As String = WebConfigurationManager.AppSettings.[Get]("ClientSigningCertificatePassword")

    Private Shared ReadOnly ClientCertificate As X509Certificate2 = If((String.IsNullOrEmpty(ClientSigningCertificatePath) OrElse String.IsNullOrEmpty(ClientSigningCertificatePassword)), Nothing, New X509Certificate2(ClientSigningCertificatePath, ClientSigningCertificatePassword))

    Private Shared ReadOnly SigningCredentials As X509SigningCredentials = If((ClientCertificate Is Nothing), Nothing, New X509SigningCredentials(ClientCertificate, SecurityAlgorithms.RsaSha256Signature, SecurityAlgorithms.Sha256Digest))

#End Region

#Region "méthodes privées"

    Private Shared Function CreateAcsClientContextForUrl(properties As SPRemoteEventProperties, sharepointUrl As Uri) As ClientContext
        Dim contextTokenString As String = properties.ContextToken

        If [String].IsNullOrEmpty(contextTokenString) Then
            Return Nothing
        End If

        Dim contextToken As SharePointContextToken = ReadAndValidateContextToken(contextTokenString, OperationContext.Current.IncomingMessageHeaders.To.Host)

        Dim accessToken As String = GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken
        Return GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken)
    End Function

    Private Shared Function GetAcsMetadataEndpointUrl() As String
        Return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl)
    End Function

    Private Shared Function GetFormattedPrincipal(principalName As String, hostName As String, targetRealm As String) As String
        If Not [String].IsNullOrEmpty(hostName) Then
            Return [String].Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, targetRealm)
        End If

        Return [String].Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, targetRealm)
    End Function

    Private Shared Function GetAcsPrincipalName(targetRealm As String) As String
        Return GetFormattedPrincipal(AcsPrincipalName, New Uri(GetAcsGlobalEndpointUrl()).Host, targetRealm)
    End Function

    Private Shared Function GetAcsGlobalEndpointUrl() As String
        Return [String].Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl)
    End Function

    Private Shared Function CreateJsonWebSecurityTokenHandler() As JsonWebSecurityTokenHandler
        Dim handler As New JsonWebSecurityTokenHandler()
        handler.Configuration = New SecurityTokenHandlerConfiguration()
        handler.Configuration.AudienceRestriction = New AudienceRestriction(AudienceUriMode.Never)
        handler.Configuration.CertificateValidator = X509CertificateValidator.None

        Dim securityKeys As New List(Of Byte())()
        securityKeys.Add(Convert.FromBase64String(ClientSecret))
        If Not String.IsNullOrEmpty(SecondaryClientSecret) Then
            securityKeys.Add(Convert.FromBase64String(SecondaryClientSecret))
        End If

        Dim securityTokens As New List(Of SecurityToken)()
        securityTokens.Add(New MultipleSymmetricKeySecurityToken(securityKeys))

        handler.Configuration.IssuerTokenResolver =
            SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
                New ReadOnlyCollection(Of SecurityToken)(securityTokens), False)
        Dim issuerNameRegistry As New SymmetricKeyIssuerNameRegistry()
        For Each securityKey As Byte() In securityKeys
            issuerNameRegistry.AddTrustedIssuer(securityKey, GetAcsPrincipalName(ServiceNamespace))
        Next

        handler.Configuration.IssuerNameRegistry = issuerNameRegistry
        Return handler
    End Function

    Private Shared Function GetS2SAccessTokenWithClaims(targetApplicationHostName As String, targetRealm As String, claims As IEnumerable(Of JsonWebTokenClaim)) As String
        Return IssueToken(ClientId, IssuerId, targetRealm, SharePointPrincipal, targetRealm, targetApplicationHostName, True,
                          claims, claims Is Nothing)
    End Function

    Private Shared Function GetClaimsWithWindowsIdentity(identity As WindowsIdentity) As JsonWebTokenClaim()
        Dim claims As JsonWebTokenClaim() = New JsonWebTokenClaim() _
                {New JsonWebTokenClaim(NameIdentifierClaimType, identity.User.Value.ToLower()),
                 New JsonWebTokenClaim("nii", "urn:office:idp:activedirectory")}
        Return claims
    End Function

    Private Shared Function IssueToken(sourceApplication As String, issuerApplication As String, sourceRealm As String, targetApplication As String, targetRealm As String, targetApplicationHostName As String, trustedForDelegation As Boolean, _
                                       claims As IEnumerable(Of JsonWebTokenClaim), Optional appOnly As Boolean = False) As String
        If SigningCredentials Is Nothing Then
            Throw New InvalidOperationException("SigningCredentials was not initialized")
        End If

        '#Region "Jeton acteur"

        Dim issuer As String = If(String.IsNullOrEmpty(sourceRealm), issuerApplication, String.Format("{0}@{1}", issuerApplication, sourceRealm))
        Dim nameid As String = If(String.IsNullOrEmpty(sourceRealm), sourceApplication, String.Format("{0}@{1}", sourceApplication, sourceRealm))
        Dim audience As String = String.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm)

        Dim actorClaims As New List(Of JsonWebTokenClaim)()
        actorClaims.Add(New JsonWebTokenClaim(JsonWebTokenConstants.ReservedClaims.NameIdentifier, nameid))
        If trustedForDelegation AndAlso Not appOnly Then
            actorClaims.Add(New JsonWebTokenClaim(TrustedForImpersonationClaimType, "true"))
        End If

        ' Créer un jeton
        Dim actorToken As New JsonWebSecurityToken(issuer:=issuer, audience:=audience, validFrom:=DateTime.UtcNow, validTo:=DateTime.UtcNow.Add(HighTrustAccessTokenLifetime), signingCredentials:=SigningCredentials, claims:=actorClaims)

        Dim actorTokenString As String = New JsonWebSecurityTokenHandler().WriteTokenAsString(actorToken)

        If appOnly Then
            ' Le jeton d'application uniquement est identique au jeton d'acteur pour le cas délégué
            Return actorTokenString
        End If

        '#End Region

        '#Region "Jeton externe"

        Dim outerClaims As List(Of JsonWebTokenClaim) = If(claims Is Nothing, New List(Of JsonWebTokenClaim)(), New List(Of JsonWebTokenClaim)(claims))
        outerClaims.Add(New JsonWebTokenClaim(ActorTokenClaimType, actorTokenString))

        ' l'émetteur de jeton externe doit correspondre à l'ID de nom du jeton acteur
        Dim jsonToken As New JsonWebSecurityToken(nameid, audience, DateTime.UtcNow, DateTime.UtcNow.Add(HighTrustAccessTokenLifetime), outerClaims)

        Dim accessToken As String = New JsonWebSecurityTokenHandler().WriteTokenAsString(jsonToken)

        '#End Region

        Return accessToken
    End Function

#End Region

#Region "AcsMetadataParser"

    ' Cette classe est utilisée pour obtenir le document MetaData du point de terminaison STS global. Elle contient
    ' des méthodes pour analyser le document MetaData et obtenir des points de terminaison ainsi qu'un certificat STS.
    Public NotInheritable Class AcsMetadataParser
        Private Sub New()
        End Sub

        Public Shared Function GetAcsSigningCert(realm As String) As X509Certificate2
            Dim document As JsonMetadataDocument = GetMetadataDocument(realm)

            If document.keys IsNot Nothing AndAlso document.keys.Count > 0 Then
                Dim signingKey As JsonKey = document.keys(0)

                If signingKey IsNot Nothing AndAlso signingKey.keyValue IsNot Nothing Then
                    Return New X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value))
                End If
            End If

            Throw New Exception("Metadata document does not contain ACS signing certificate.")
        End Function

        Public Shared Function GetDelegationServiceUrl(realm As String) As String
            Dim document As JsonMetadataDocument = GetMetadataDocument(realm)

            Dim delegationEndpoint As JsonEndpoint = document.endpoints.SingleOrDefault(Function(e) e.protocol = DelegationIssuance)

            If delegationEndpoint IsNot Nothing Then
                Return delegationEndpoint.location
            End If

            Throw New Exception("Metadata document does not contain Delegation Service endpoint Url")
        End Function

        Private Shared Function GetMetadataDocument(realm As String) As JsonMetadataDocument
            Dim acsMetadataEndpointUrlWithRealm As String = [String].Format(CultureInfo.InvariantCulture, "{0}?realm={1}", GetAcsMetadataEndpointUrl(), realm)
            Dim acsMetadata As Byte()
            Using webClient As New WebClient()
                acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm)
            End Using
            Dim jsonResponseString As String = Encoding.UTF8.GetString(acsMetadata)

            Dim serializer As New JavaScriptSerializer()
            Dim document As JsonMetadataDocument = serializer.Deserialize(Of JsonMetadataDocument)(jsonResponseString)

            If document Is Nothing Then
                Throw New Exception("No metadata document found at the global endpoint " & acsMetadataEndpointUrlWithRealm)
            End If

            Return document
        End Function

        Public Shared Function GetStsUrl(realm As String) As String
            Dim document As JsonMetadataDocument = GetMetadataDocument(realm)

            Dim s2sEndpoint As JsonEndpoint = document.endpoints.SingleOrDefault(Function(e) e.protocol = S2SProtocol)

            If s2sEndpoint IsNot Nothing Then
                Return s2sEndpoint.location
            End If

            Throw New Exception("Metadata document does not contain STS endpoint url")
        End Function

        Private Class JsonMetadataDocument
            Public Property serviceName() As String
                Get
                    Return m_serviceName
                End Get
                Set(value As String)
                    m_serviceName = value
                End Set
            End Property

            Private m_serviceName As String

            Public Property endpoints() As List(Of JsonEndpoint)
                Get
                    Return m_endpoints
                End Get
                Set(value As List(Of JsonEndpoint))
                    m_endpoints = value
                End Set
            End Property

            Private m_endpoints As List(Of JsonEndpoint)

            Public Property keys() As List(Of JsonKey)
                Get
                    Return m_keys
                End Get
                Set(value As List(Of JsonKey))
                    m_keys = value
                End Set
            End Property

            Private m_keys As List(Of JsonKey)
        End Class

        Private Class JsonEndpoint
            Public Property location() As String
                Get
                    Return m_location
                End Get
                Set(value As String)
                    m_location = value
                End Set
            End Property

            Private m_location As String

            Public Property protocol() As String
                Get
                    Return m_protocol
                End Get
                Set(value As String)
                    m_protocol = value
                End Set
            End Property

            Private m_protocol As String

            Public Property usage() As String
                Get
                    Return m_usage
                End Get
                Set(value As String)
                    m_usage = value
                End Set
            End Property

            Private m_usage As String
        End Class

        Private Class JsonKeyValue
            Public Property type() As String
                Get
                    Return m_type
                End Get
                Set(value As String)
                    m_type = value
                End Set
            End Property

            Private m_type As String

            Public Property value() As String
                Get
                    Return m_value
                End Get
                Set(value As String)
                    m_value = value
                End Set
            End Property

            Private m_value As String
        End Class

        Private Class JsonKey
            Public Property usage() As String
                Get
                    Return m_usage
                End Get
                Set(value As String)
                    m_usage = value
                End Set
            End Property

            Private m_usage As String

            Public Property keyValue() As JsonKeyValue
                Get
                    Return m_keyValue
                End Get
                Set(value As JsonKeyValue)
                    m_keyValue = value
                End Set
            End Property

            Private m_keyValue As JsonKeyValue
        End Class
    End Class

#End Region
End Class

''' <summary>
''' JsonWebSecurityToken généré par SharePoint pour authentifier une application tierce et permettre les rappels à l'aide d'un jeton d'actualisation
''' </summary>
Public Class SharePointContextToken
    Inherits JsonWebSecurityToken

    Public Shared Function Create(contextToken As JsonWebSecurityToken) As SharePointContextToken
        Return New SharePointContextToken(contextToken.Issuer, contextToken.Audience, contextToken.ValidFrom, contextToken.ValidTo, contextToken.Claims)
    End Function

    Public Sub New(issuer As String, audience As String, validFrom As DateTime, validTo As DateTime, claims As IEnumerable(Of JsonWebTokenClaim))
        MyBase.New(issuer, audience, validFrom, validTo, claims)
    End Sub

    Public Sub New(issuer As String, audience As String, validFrom As DateTime, validTo As DateTime, claims As IEnumerable(Of JsonWebTokenClaim), issuerToken As SecurityToken, _
                   actorToken As JsonWebSecurityToken)
        MyBase.New(issuer, audience, validFrom, validTo, claims, issuerToken, _
                   actorToken)
    End Sub

    Public Sub New(issuer As String, audience As String, validFrom As DateTime, validTo As DateTime, claims As IEnumerable(Of JsonWebTokenClaim), signingCredentials As SigningCredentials)
        MyBase.New(issuer, audience, validFrom, validTo, claims, signingCredentials)
    End Sub

    Public ReadOnly Property NameId() As String
        Get
            Return GetClaimValue(Me, "nameid")
        End Get
    End Property

    ''' <summary>
    ''' Partie nom du principal de la demande appctxsender du jeton de contexte
    ''' </summary>
    Public ReadOnly Property TargetPrincipalName() As String
        Get
            Dim appctxsender As String = GetClaimValue(Me, "appctxsender")

            If appctxsender Is Nothing Then
                Return Nothing
            End If

            Return appctxsender.Split("@"c)(0)
        End Get
    End Property

    ''' <summary>
    ''' Demande refreshtoken du jeton de contexte
    ''' </summary>
    Public ReadOnly Property RefreshToken() As String
        Get
            Return GetClaimValue(Me, "refreshtoken")
        End Get
    End Property

    ''' <summary>
    ''' Demande CacheKey du jeton de contexte
    ''' </summary>
    Public ReadOnly Property CacheKey() As String
        Get
            Dim appctx As String = GetClaimValue(Me, "appctx")
            If appctx Is Nothing Then
                Return Nothing
            End If

            Dim ctx As New ClientContext("http://tempuri.org")
            Dim dict As Dictionary(Of String, Object) = DirectCast(ctx.ParseObjectFromJsonString(appctx), Dictionary(Of String, Object))
            Dim cacheKeyString As String = DirectCast(dict("CacheKey"), String)

            Return cacheKeyString
        End Get
    End Property

    ''' <summary>
    ''' Demande SecurityTokenServiceUri du jeton de contexte
    ''' </summary>
    Public ReadOnly Property SecurityTokenServiceUri() As String
        Get
            Dim appctx As String = GetClaimValue(Me, "appctx")
            If appctx Is Nothing Then
                Return Nothing
            End If

            Dim ctx As New ClientContext("http://tempuri.org")
            Dim dict As Dictionary(Of String, Object) = DirectCast(ctx.ParseObjectFromJsonString(appctx), Dictionary(Of String, Object))
            Dim securityTokenServiceUriString As String = DirectCast(dict("SecurityTokenServiceUri"), String)

            Return securityTokenServiceUriString
        End Get
    End Property

    ''' <summary>
    ''' Partie de domaine de la demande audience du jeton de contexte
    ''' </summary>
    Public ReadOnly Property Realm() As String
        Get
            Dim aud As String = Audience
            If aud Is Nothing Then
                Return Nothing
            End If

            Dim tokenRealm As String = aud.Substring(aud.IndexOf("@"c) + 1)

            Return tokenRealm
        End Get
    End Property

    Private Shared Function GetClaimValue(token As JsonWebSecurityToken, claimType As String) As String
        If token Is Nothing Then
            Throw New ArgumentNullException("token")
        End If

        For Each claim As JsonWebTokenClaim In token.Claims
            If StringComparer.Ordinal.Equals(claim.ClaimType, claimType) Then
                Return claim.Value
            End If
        Next

        Return Nothing
    End Function
End Class

''' <summary>
''' Représente un jeton de sécurité qui contient plusieurs clés de sécurité générées à l'aide d'algorithmes symétriques.
''' </summary>
Public Class MultipleSymmetricKeySecurityToken
    Inherits SecurityToken

    ''' <summary>
    ''' Initialise une nouvelle instance de la classe MultipleSymmetricKeySecurityToken.
    ''' </summary>
    ''' <param name="keys">Énumération des tableaux d'octets qui contiennent les clés symétriques.</param>
    Public Sub New(keys As IEnumerable(Of Byte()))
        Me.New(UniqueId.CreateUniqueId(), keys)
    End Sub

    ''' <summary>
    ''' Initialise une nouvelle instance de la classe MultipleSymmetricKeySecurityToken.
    ''' </summary>
    ''' <param name="tokenId">Identificateur unique du jeton de sécurité.</param>
    ''' <param name="keys">Énumération des tableaux d'octets qui contiennent les clés symétriques.</param>
    Public Sub New(tokenId As String, keys As IEnumerable(Of Byte()))
        If keys Is Nothing Then
            Throw New ArgumentNullException("keys")
        End If

        If String.IsNullOrEmpty(tokenId) Then
            Throw New ArgumentException("Value cannot be a null or empty string.", "tokenId")
        End If

        For Each key As Byte() In keys
            If key.Length <= 0 Then
                Throw New ArgumentException("The key length must be greater then zero.", "keys")
            End If
        Next

        m_id = tokenId
        m_effectiveTime = DateTime.UtcNow
        m_securityKeys = CreateSymmetricSecurityKeys(keys)
    End Sub

    ''' <summary>
    ''' Obtient un identificateur unique du jeton de sécurité.
    ''' </summary>
    Public Overrides ReadOnly Property Id As String
        Get
            Return m_id
        End Get
    End Property

    ''' <summary>
    ''' Obtient les clés de chiffrement associées au jeton de sécurité.
    ''' </summary>
    Public Overrides ReadOnly Property SecurityKeys() As ReadOnlyCollection(Of SecurityKey)
        Get
            Return m_securityKeys.AsReadOnly()
        End Get
    End Property

    ''' <summary>
    ''' Obtient le premier moment de validité de ce jeton de sécurité.
    ''' </summary>
    Public Overrides ReadOnly Property ValidFrom As DateTime
        Get
            Return m_effectiveTime
        End Get
    End Property

    ''' <summary>
    ''' Obtient le dernier moment de validité de ce jeton de sécurité.
    ''' </summary>
    Public Overrides ReadOnly Property ValidTo As DateTime
        Get
            ' Ne jamais expirer
            Return Date.MaxValue
        End Get
    End Property

    ''' <summary>
    ''' Retourne une valeur qui indique si l'identificateur de clé pour cette instance peut être résolu sur l'identificateur de clé spécifié.
    ''' </summary>
    ''' <param name="keyIdentifierClause">Une SecurityKeyIdentifierClause à comparer à cette instance.</param>
    ''' <returns>true si keyIdentifierClause est une SecurityKeyIdentifierClause et a le même identificateur unique que la propriété ID ; sinon, false.</returns>
    Public Overrides Function MatchesKeyIdentifierClause(keyIdentifierClause As SecurityKeyIdentifierClause) As Boolean
        If keyIdentifierClause Is Nothing Then
            Throw New ArgumentNullException("keyIdentifierClause")
        End If

        ' Étant donné qu'il s'agit d'un jeton symétrique et que nous n'avons pas d'ID pour différencier les jetons, nous recherchons simplement la
        ' présence d'un SymmetricIssuerKeyIdentifier. Le mappage réel avec l'émetteur a lieu ultérieurement
        ' lorsque la clé est mise en correspondance avec l'émetteur.
        If TypeOf keyIdentifierClause Is SymmetricIssuerKeyIdentifierClause Then
            Return True
        End If
        Return MyBase.MatchesKeyIdentifierClause(keyIdentifierClause)
    End Function

#Region "membres privés"

    Private Function CreateSymmetricSecurityKeys(keys As IEnumerable(Of Byte())) As List(Of SecurityKey)
        Dim symmetricKeys As New List(Of SecurityKey)()
        For Each key As Byte() In keys
            symmetricKeys.Add(New InMemorySymmetricSecurityKey(key))
        Next
        Return symmetricKeys
    End Function

    Private m_id As String
    Private m_effectiveTime As DateTime
    Private m_securityKeys As List(Of SecurityKey)

#End Region
End Class
