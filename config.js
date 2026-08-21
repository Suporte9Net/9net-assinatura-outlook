(function () {
  "use strict";

  window.NineNetConfig = Object.freeze({
    version: "2.0.2",
    clientId: "eaaf2251-f42e-4705-861f-66dd100ca29b",
    tenantId: "e66ee0b3-9cf5-40f5-bc5a-a46efc7a8063",
    baseUrl: "https://9net-assinatura-outlook.pages.dev",
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
