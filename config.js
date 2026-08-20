(function () {
  "use strict";

  window.NineNetConfig = Object.freeze({
    version: "2.0.0",
    clientId: "CLIENT_ID_PENDENTE",
    tenantId: "TENANT_ID_PENDENTE",
    baseUrl: "https://suporte9net.github.io/9net-assinatura-outlook",
    graphScopes: Object.freeze(["User.Read"]),
    graphUserUrl: "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail,userPrincipalName,businessPhones,faxNumber,mobilePhone,streetAddress,city,state,postalCode,country,department,officeLocation",
    websiteUrl: "https://9net.com.br/",
    linkedinUrl: "https://www.linkedin.com/company/9net-tecnologia/",
    instagramUrl: "https://www.instagram.com/9nettecnologia/",
    fallbackPhone: "(11) 2574-9600",
    fallbackAddress: Object.freeze([
      "Av. Eng. Luís Carlos Berrini, 1645",
      "8º andar, Conj. 82 Cidade Monções,",
      "São Paulo/SP"
    ])
  });
})();
