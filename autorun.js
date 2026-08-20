(function () {
  "use strict";

  async function checkSignature(event) {
    try {
      if (NineNetSignature.isClassicWindowsOutlook()) {
        console.log("9Net: Outlook classico detectado; a GPO permanece responsavel pela assinatura.");
        return;
      }

      var profile;
      try {
        profile = await NineNetSignature.getGraphProfile(false, NineNetConfig);
      } catch (graphError) {
        console.error("9Net: consulta silenciosa ao Microsoft Graph falhou; usando dados basicos.", graphError);
        profile = NineNetSignature.getOfficeFallbackProfile(NineNetConfig);
      }

      var html = NineNetSignature.buildHtml(profile, NineNetConfig);
      await NineNetSignature.setSignature(html);
      console.log("9Net: assinatura aplicada com sucesso.");
    } catch (error) {
      console.error("9Net: nao foi possivel aplicar a assinatura.", error);
    } finally {
      event.completed();
    }
  }

  Office.actions.associate("checkSignature", checkSignature);
})();

