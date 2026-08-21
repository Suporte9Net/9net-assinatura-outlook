(function () {
  "use strict";

  var currentProfile = null;
  var connectButton;
  var applyButton;
  var statusBox;

  function setStatus(message, type) {
    statusBox.textContent = message;
    statusBox.className = "status" + (type ? " " + type : "");
  }

  function setBusy(busy) {
    connectButton.disabled = busy;
    applyButton.disabled = busy || !currentProfile;
  }

  function text(id, value) {
    document.getElementById(id).textContent = value || "—";
  }

  function showProfile(profile) {
    text("profileName", profile.displayName);
    text("profileTitle", profile.jobTitle);
    text("profileEmail", profile.email);
    text("profilePhone", profile.businessPhone);
    text("profileFax", profile.faxNumber);
    text("profileMobile", profile.mobilePhone);
    text("profileAddress", (profile.addressLines || []).join(" · "));
    document.getElementById("profile").classList.remove("hidden");
    document.getElementById("preview").innerHTML = NineNetSignature.buildHtml(profile, NineNetConfig);
    document.getElementById("previewSection").classList.remove("hidden");
  }

  async function connectAndTest() {
    setBusy(true);
    setStatus("Consultando seu perfil no Microsoft 365...", "");
    try {
      if (NineNetSignature.isClassicWindowsOutlook()) {
        throw new Error("No Outlook classico, a assinatura continua sendo administrada pela GPO. Teste este suplemento no Novo Outlook.");
      }
      currentProfile = await NineNetSignature.getGraphProfile(true, NineNetConfig);
      showProfile(currentProfile);
      try {
        await NineNetSignature.saveCachedProfile(currentProfile, NineNetConfig);
        setStatus("Dados consultados e salvos para aplicacao automatica.", "success");
      } catch (cacheError) {
        console.warn("9Net: nao foi possivel salvar o perfil para o modo automatico.", cacheError);
        setStatus("Dados consultados com sucesso, mas o perfil nao pode ser salvo para o modo automatico.", "error");
      }
    } catch (error) {
      currentProfile = null;
      setStatus(error && error.message ? error.message : String(error), "error");
    } finally {
      setBusy(false);
    }
  }

  async function applyCurrent() {
    if (!currentProfile) {
      return;
    }
    setBusy(true);
    setStatus("Aplicando a assinatura nesta mensagem...", "");
    try {
      await NineNetSignature.setSignature(NineNetSignature.buildHtml(currentProfile, NineNetConfig));
      setStatus("Assinatura aplicada com sucesso.", "success");
    } catch (error) {
      setStatus(error && error.message ? error.message : String(error), "error");
    } finally {
      setBusy(false);
    }
  }

  Office.onReady(function () {
    connectButton = document.getElementById("connectButton");
    applyButton = document.getElementById("applyButton");
    statusBox = document.getElementById("status");
    connectButton.addEventListener("click", connectAndTest);
    applyButton.addEventListener("click", applyCurrent);

    if (NineNetSignature.isPendingConfig(NineNetConfig)) {
      setStatus("A configuracao do Microsoft Entra ainda esta pendente.", "error");
      connectButton.disabled = true;
      return;
    }

    var cachedProfile = NineNetSignature.getCachedProfile(NineNetConfig);
    if (cachedProfile) {
      currentProfile = cachedProfile;
      showProfile(currentProfile);
      applyButton.disabled = false;
      setStatus("Perfil salvo encontrado. O modo automatico ja pode usa-lo; use Entrar e testar para atualizar os dados.", "success");
    }
  });
})();

