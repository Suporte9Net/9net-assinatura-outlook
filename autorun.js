(function () {
  "use strict";

  // Inicializa o runtime de navegador usado pelo Outlook na Web e pelo Novo Outlook.
  // O Outlook classico usa classic-skip.js pelo Override do manifest.
  Office.onReady();

  var GRAPH_TIMEOUT_MS = 4500;

  function withTimeout(promise, milliseconds) {
    return new Promise(function (resolve, reject) {
      var finished = false;
      var timer = setTimeout(function () {
        if (!finished) {
          finished = true;
          reject(new Error("Tempo limite ao consultar o Microsoft Graph no runtime automatico."));
        }
      }, milliseconds);

      Promise.resolve(promise).then(function (value) {
        if (!finished) {
          finished = true;
          clearTimeout(timer);
          resolve(value);
        }
      }, function (error) {
        if (!finished) {
          finished = true;
          clearTimeout(timer);
          reject(error);
        }
      });
    });
  }

  async function applyProfile(profile) {
    var html = NineNetSignature.buildHtml(profile, NineNetConfig);
    await NineNetSignature.setSignature(html);
  }

  async function runAutomaticSignature() {
    // Primeiro aplica imediatamente um perfil conhecido. Isso evita que a assinatura
    // dependa da autenticacao do Graph para aparecer.
    var cachedProfile = NineNetSignature.getCachedProfile(NineNetConfig);
    var initialProfile = cachedProfile || NineNetSignature.getOfficeFallbackProfile(NineNetConfig);
    await applyProfile(initialProfile);
    console.log("9Net: assinatura inicial aplicada automaticamente.");

    // Em seguida tenta atualizar com os dados completos do Microsoft 365.
    // O timeout impede que uma autenticacao em segundo plano travada impeça a assinatura.
    try {
      var graphProfile = await withTimeout(
        NineNetSignature.getGraphProfile(false, NineNetConfig),
        GRAPH_TIMEOUT_MS
      );
      await applyProfile(graphProfile);
      try {
        await NineNetSignature.saveCachedProfile(graphProfile, NineNetConfig);
      } catch (cacheError) {
        console.warn("9Net: assinatura atualizada, mas o perfil nao pode ser salvo.", cacheError);
      }
      console.log("9Net: assinatura atualizada com dados completos do Microsoft Graph.");
    } catch (graphError) {
      console.warn("9Net: Graph silencioso indisponivel; mantendo perfil salvo/fallback.", graphError);
    }
  }

  function checkSignature(event) {
    Promise.resolve()
      .then(runAutomaticSignature)
      .catch(function (error) {
        console.error("9Net: nao foi possivel aplicar a assinatura automaticamente.", error);
      })
      .finally(function () {
        try {
          event.completed();
        } catch (completeError) {
          console.error("9Net: falha ao finalizar o evento automatico.", completeError);
        }
      });
  }

  Office.actions.associate("checkSignature", checkSignature);
})();
