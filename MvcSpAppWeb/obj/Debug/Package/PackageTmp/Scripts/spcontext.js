(function (window, undefined) {

    "use strict";

    var $ = window.jQuery;
    var document = window.document;

    // Nom du paramètre SPHostUrl
    var SPHostUrlKey = "SPHostUrl";

    // Obtient SPHostUrl à partir de l'URL actuelle et l'ajoute en tant que chaîne de requête à chaque lien qui pointe vers le domaine actuel de la page.
    $(document).ready(function () {
        ensureSPHasRedirectedToSharePointRemoved();

        var spHostUrl = getSPHostUrlFromQueryString(window.location.search);
        var currentAuthority = getAuthorityFromUrl(window.location.href).toUpperCase();

        if (spHostUrl && currentAuthority) {
            appendSPHostUrlToLinks(spHostUrl, currentAuthority);
        }
    });

    // Ajoute SPHostUrl en tant que chaîne de requête à tous les liens qui pointent vers le domaine actuel.
    function appendSPHostUrlToLinks(spHostUrl, currentAuthority) {
        $("a")
            .filter(function () {
                var authority = getAuthorityFromUrl(this.href);
                if (!authority && /^#|:/.test(this.href)) {
                    // Filtre les ancres et les URL qui comportent d'autres protocoles non pris en charge.
                    return false;
                }
                return authority.toUpperCase() == currentAuthority;
            })
            .each(function () {
                if (!getSPHostUrlFromQueryString(this.search)) {
                    if (this.search.length > 0) {
                        this.search += "&" + SPHostUrlKey + "=" + spHostUrl;
                    }
                    else {
                        this.search = "?" + SPHostUrlKey + "=" + spHostUrl;
                    }
                }
            });
    }

    // Obtient SPHostUrl à partir de la chaîne de requête spécifiée.
    function getSPHostUrlFromQueryString(queryString) {
        if (queryString) {
            if (queryString[0] === "?") {
                queryString = queryString.substring(1);
            }

            var keyValuePairArray = queryString.split("&");

            for (var i = 0; i < keyValuePairArray.length; i++) {
                var currentKeyValuePair = keyValuePairArray[i].split("=");

                if (currentKeyValuePair.length > 1 && currentKeyValuePair[0] == SPHostUrlKey) {
                    return currentKeyValuePair[1];
                }
            }
        }

        return null;
    }

    // Obtient l'autorité de l'URL spécifiée lorsqu'il s'agit d'une URL absolue qui contient le protocole http/https ou une URL relative de protocole.
    function getAuthorityFromUrl(url) {
        if (url) {
            var match = /^(?:https:\/\/|http:\/\/|\/\/)([^\/\?#]+)(?:\/|#|$|\?)/i.exec(url);
            if (match) {
                return match[1];
            }
        }
        return null;
    }

    // Si la chaîne de requête contient SPHasRedirectedToSharePoint, supprimez-le.
    // Ainsi, lorsque l'utilisateur crée un signet de l'URL, SPHasRedirectedToSharePoint n'est pas inclus.
    // Notez que la modification de window.location.search entraînera une requête supplémentaire au serveur.
    function ensureSPHasRedirectedToSharePointRemoved() {
        var SPHasRedirectedToSharePointParam = "&SPHasRedirectedToSharePoint=1";

        var queryString = window.location.search;

        if (queryString.indexOf(SPHasRedirectedToSharePointParam) >= 0) {
            window.location.search = queryString.replace(SPHasRedirectedToSharePointParam, "");
        }
    }

})(window);
