(function (global) {
  "use strict";

  var msalInstancePromise = null;

  function clean(value) {
    return value === null || value === undefined ? "" : String(value).trim();
  }

  function escapeHtml(value) {
    return clean(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function toTelHref(value) {
    var phone = clean(value).replace(/[^0-9+]/g, "");
    return phone ? "tel:" + phone : "#";
  }

  function joinCityState(city, state) {
    city = clean(city);
    state = clean(state);
    if (city && state) {
      return city + "/" + state;
    }
    return city || state;
  }

  function buildAddressLines(raw, config) {
    var street = clean(raw.streetAddress);
    var cityState = joinCityState(raw.city, raw.state);
    var postalCode = clean(raw.postalCode);
    var country = clean(raw.country);
    var hasGraphAddress = Boolean(street || cityState || postalCode || country);

    if (!hasGraphAddress) {
      return Array.prototype.slice.call(config.fallbackAddress || []);
    }

    return [
      street,
      cityState,
      [postalCode, country].filter(Boolean).join(" · ")
    ].filter(Boolean);
  }

  function normalizeGraphProfile(raw, fallback, config) {
    raw = raw || {};
    fallback = fallback || {};
    config = config || global.NineNetConfig;
    var phones = Array.isArray(raw.businessPhones) ? raw.businessPhones : [];

    return {
      displayName: clean(raw.displayName) || clean(fallback.displayName) || "Colaborador 9Net",
      jobTitle: clean(raw.jobTitle),
      email: clean(raw.mail) || clean(raw.userPrincipalName) || clean(fallback.email),
      businessPhone: clean(phones[0]) || clean(config.fallbackPhone),
      faxNumber: clean(raw.faxNumber),
      mobilePhone: clean(raw.mobilePhone),
      department: clean(raw.department),
      officeLocation: clean(raw.officeLocation),
      addressLines: buildAddressLines(raw, config)
    };
  }

  function getOfficeFallbackProfile(config) {
    config = config || global.NineNetConfig;
    var userProfile = {};
    try {
      userProfile = Office.context.mailbox.userProfile || {};
    } catch (error) {
      userProfile = {};
    }

    return {
      displayName: clean(userProfile.displayName) || "Colaborador 9Net",
      jobTitle: "",
      email: clean(userProfile.emailAddress),
      businessPhone: clean(config.fallbackPhone),
      faxNumber: "",
      mobilePhone: "",
      department: "",
      officeLocation: "",
      addressLines: Array.prototype.slice.call(config.fallbackAddress || [])
    };
  }

  function isPendingConfig(config) {
    return !config ||
      !clean(config.clientId) ||
      !clean(config.tenantId) ||
      /PENDENTE/i.test(config.clientId) ||
      /PENDENTE/i.test(config.tenantId);
  }

  function supportsNaa() {
    try {
      return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
    } catch (error) {
      return false;
    }
  }

  async function getMsalInstance(config) {
    config = config || global.NineNetConfig;
    if (isPendingConfig(config)) {
      throw new Error("A configuracao do Microsoft Entra ainda nao foi gerada.");
    }
    if (!supportsNaa()) {
      throw new Error("Este cliente do Outlook nao oferece Nested App Authentication 1.1.");
    }
    if (!global.msal || typeof global.msal.createNestablePublicClientApplication !== "function") {
      throw new Error("A biblioteca de autenticacao MSAL nao foi carregada.");
    }

    if (!msalInstancePromise) {
      msalInstancePromise = global.msal.createNestablePublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: "https://login.microsoftonline.com/" + config.tenantId
        },
        cache: {
          cacheLocation: "localStorage"
        }
      });
    }
    return msalInstancePromise;
  }

  async function acquireAccessToken(interactive, config) {
    config = config || global.NineNetConfig;
    var client = await getMsalInstance(config);
    var request = { scopes: Array.prototype.slice.call(config.graphScopes || ["User.Read"]) };

    try {
      var silentResult = await client.acquireTokenSilent(request);
      return silentResult.accessToken;
    } catch (silentError) {
      if (!interactive) {
        throw silentError;
      }
      var popupResult = await client.acquireTokenPopup(request);
      return popupResult.accessToken;
    }
  }

  async function getGraphProfile(interactive, config) {
    config = config || global.NineNetConfig;
    var accessToken = await acquireAccessToken(Boolean(interactive), config);
    var response = await fetch(config.graphUserUrl, {
      method: "GET",
      headers: {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json"
      }
    });

    if (!response.ok) {
      throw new Error("O Microsoft Graph respondeu HTTP " + response.status + ".");
    }

    var graphUser = await response.json();
    return normalizeGraphProfile(graphUser, getOfficeFallbackProfile(config), config);
  }

  function isClassicWindowsOutlook() {
    try {
      var hostName = clean(Office.context.diagnostics.hostName).toLowerCase();
      return hostName === "outlook";
    } catch (error) {
      return false;
    }
  }

  function buildHtml(profile, config) {
    profile = profile || {};
    config = config || global.NineNetConfig;

    var displayName = escapeHtml(profile.displayName || "Colaborador 9Net");
    var jobTitle = clean(profile.jobTitle);
    var email = clean(profile.email);
    var phone = clean(profile.businessPhone) || clean(config.fallbackPhone);
    var fax = clean(profile.faxNumber);
    var mobile = clean(profile.mobilePhone);
    var addressLines = Array.isArray(profile.addressLines) && profile.addressLines.length
      ? profile.addressLines
      : Array.prototype.slice.call(config.fallbackAddress || []);
    var assets = config.baseUrl.replace(/\/$/, "") + "/assets";

    var jobTitleBlock = jobTitle
      ? '<div class="theme-text-secondary" style="margin:2px 0 0 0;color:#555555;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:23px;font-weight:600;mso-line-height-rule:exactly;">' + escapeHtml(jobTitle) + "</div>"
      : "";

    var phoneCells = fax
      ? '<td width="160" valign="middle" nowrap style="width:160px;padding:8px 0;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;mso-line-height-rule:exactly;white-space:nowrap;"><a class="theme-link" href="' + escapeHtml(toTelHref(phone)) + '" style="color:#222222;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;text-decoration:none;white-space:nowrap;">' + escapeHtml(phone) + '</a></td><td class="theme-text-primary" width="29" valign="middle" nowrap style="width:29px;padding:8px 0;color:#222222;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;mso-line-height-rule:exactly;white-space:nowrap;">|</td><td class="theme-text-primary" width="171" valign="middle" nowrap style="width:171px;padding:8px 0;color:#222222;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;mso-line-height-rule:exactly;white-space:nowrap;">Ramal ' + escapeHtml(fax) + "</td>"
      : '<td colspan="3" width="360" valign="middle" nowrap style="width:360px;padding:8px 0;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;mso-line-height-rule:exactly;white-space:nowrap;"><a class="theme-link" href="' + escapeHtml(toTelHref(phone)) + '" style="color:#222222;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;text-decoration:none;white-space:nowrap;">' + escapeHtml(phone) + "</a></td>";

    var mobileRow = mobile
      ? '<tr><td width="52" valign="middle" style="width:52px;padding:8px 0;"><img class="theme-icon" src="' + assets + '/icon-mobile.png" width="24" height="24" alt="Celular" style="display:block;width:24px;height:24px;border:0;outline:none;background-color:#ffffff;"></td><td colspan="3" width="360" valign="middle" style="width:360px;padding:8px 0;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;"><a class="theme-link" href="' + escapeHtml(toTelHref(mobile)) + '" style="color:#222222;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;text-decoration:none;white-space:nowrap;">' + escapeHtml(mobile) + "</a></td></tr>"
      : "";

    var addressHtml = addressLines.map(escapeHtml).join("<br>");

    return '' +
      '<style type="text/css">' +
      'table,td{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt}img{border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic}a{text-decoration:none}.theme-icon{background-color:#fff}' +
      '@media only screen and (max-width:900px){.signature-shell{width:100%!important;max-width:880px!important}.banner-image{width:100%!important;height:auto!important}.contact-grid{width:100%!important}.stack-column{display:block!important;width:100%!important;max-width:100%!important}.address-cell{border-left:0!important;border-top:1px solid #ddd!important;padding:16px 0 0 0!important}.contact-cell{padding:0 0 16px 0!important}}' +
      '@media (prefers-color-scheme:dark){.theme-text-primary,.theme-link{color:#f4f4f4!important}.theme-text-secondary{color:#d6d6d6!important}.theme-text-muted{color:#bdbdbd!important}.theme-divider,.address-cell{border-color:#666!important}.theme-icon{filter:invert(1)!important}}' +
      '[data-ogsc] .theme-text-primary,[data-ogsc] .theme-link{color:#f4f4f4!important}[data-ogsc] .theme-text-secondary{color:#d6d6d6!important}[data-ogsc] .theme-text-muted{color:#bdbdbd!important}[data-ogsc] .theme-divider,[data-ogsc] .address-cell{border-color:#666!important}[data-ogsc] .theme-icon{filter:invert(1)!important}' +
      '</style>' +
      '<div class="WordSection1" data-9net-signature="new-outlook-v2.0.0">' +
      '<table role="presentation" class="signature-shell" width="880" cellpadding="0" cellspacing="0" border="0" style="width:880px;max-width:880px;font-family:\'Segoe UI\',Arial,sans-serif;">' +
      '<tr><td style="padding:0;"><a href="' + escapeHtml(config.websiteUrl) + '" target="_blank" style="text-decoration:none;border:0;"><img class="banner-image" src="' + assets + '/banner-9net.jpg" width="880" height="126" alt="9Net Tecnologia - Cuidamos do seu TI para que o foco e investimento da sua empresa estejam em seus negocios." style="display:block;width:880px;max-width:880px;height:126px;border:0;outline:none;text-decoration:none;"></a></td></tr>' +
      '<tr><td style="padding:31px 0 0 0;"><div class="theme-text-primary" style="margin:0;color:#111111;font-family:\'Segoe UI\',Arial,sans-serif;font-size:24px;line-height:29px;font-weight:700;mso-line-height-rule:exactly;">' + displayName + '</div>' + jobTitleBlock + '</td></tr>' +
      '<tr><td height="19" style="height:19px;font-size:1px;line-height:19px;mso-line-height-rule:exactly;">&nbsp;</td></tr>' +
      '<tr><td style="padding:0;"><table role="presentation" class="contact-grid" width="880" cellpadding="0" cellspacing="0" border="0" style="width:880px;table-layout:fixed;"><tr>' +
      '<td class="stack-column contact-cell" width="440" valign="top" style="width:412px;padding:9px 28px 0 0;mso-padding-alt:9px 28px 0 0;"><table role="presentation" width="412" cellpadding="0" cellspacing="0" border="0" style="width:412px;table-layout:fixed;">' +
      '<tr><td width="52" valign="middle" style="width:52px;padding:8px 0;"><img class="theme-icon" src="' + assets + '/icon-email.png" width="24" height="24" alt="E-mail" style="display:block;width:24px;height:24px;border:0;outline:none;background-color:#ffffff;"></td><td colspan="3" width="360" valign="middle" style="width:360px;padding:8px 0;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;"><a class="theme-link" href="mailto:' + escapeHtml(email) + '" style="color:#222222;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:24px;text-decoration:none;word-break:break-word;">' + escapeHtml(email) + '</a></td></tr>' +
      '<tr><td width="52" valign="middle" style="width:52px;padding:8px 0;"><img class="theme-icon" src="' + assets + '/icon-phone.png" width="24" height="24" alt="Telefone" style="display:block;width:24px;height:24px;border:0;outline:none;background-color:#ffffff;"></td>' + phoneCells + '</tr>' +
      mobileRow +
      '<tr><td colspan="4" width="412" valign="middle" style="width:412px;padding:8px 0;font-size:0;line-height:0;mso-padding-alt:8px 0 8px 0;"><table role="presentation" align="left" width="112" height="24" cellpadding="0" cellspacing="0" border="0" style="width:112px;height:24px;table-layout:fixed;margin:0;"><tr>' +
      '<td width="44" height="24" align="left" valign="middle" style="width:44px;height:24px;padding:0;font-size:0;line-height:0;"><a href="' + escapeHtml(config.websiteUrl) + '" target="_blank" aria-label="Site 9Net" style="text-decoration:none;border:0;"><img class="theme-icon" src="' + assets + '/icon-website.png" width="24" height="24" alt="Site" style="display:block;width:24px;height:24px;border:0;outline:none;background-color:#ffffff;"></a></td>' +
      '<td width="44" height="24" align="left" valign="middle" style="width:44px;height:24px;padding:0;font-size:0;line-height:0;"><a href="' + escapeHtml(config.linkedinUrl) + '" target="_blank" aria-label="LinkedIn 9Net" style="text-decoration:none;border:0;"><img class="theme-icon" src="' + assets + '/icon-linkedin.png" width="24" height="24" alt="LinkedIn" style="display:block;width:24px;height:24px;border:0;outline:none;background-color:#ffffff;"></a></td>' +
      '<td width="24" height="24" align="left" valign="middle" style="width:24px;height:24px;padding:0;font-size:0;line-height:0;"><a href="' + escapeHtml(config.instagramUrl) + '" target="_blank" aria-label="Instagram 9Net" style="text-decoration:none;border:0;"><img class="theme-icon" src="' + assets + '/icon-instagram.png" width="24" height="24" alt="Instagram" style="display:block;width:24px;height:24px;border:0;outline:none;background-color:#ffffff;"></a></td>' +
      '</tr></table></td></tr></table></td>' +
      '<td class="stack-column address-cell" width="440" valign="top" style="width:412px;border-left:1px solid #dddddd;padding:1px 0 0 27px;mso-padding-alt:1px 0 0 27px;"><div class="theme-text-primary" style="margin:0 0 8px 0;color:#111111;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:23px;font-weight:700;">Endereço</div><div class="theme-text-secondary" style="margin:0;color:#444444;font-family:\'Segoe UI\',Arial,sans-serif;font-size:18px;line-height:26px;">' + addressHtml + '</div></td>' +
      '</tr></table></td></tr>' +
      '<tr><td height="14" style="height:14px;font-size:1px;line-height:14px;">&nbsp;</td></tr><tr><td class="theme-divider" height="1" style="height:1px;font-size:1px;line-height:1px;border-top:1px solid #dddddd;">&nbsp;</td></tr>' +
      '<tr><td class="theme-text-muted" style="padding:15px 0 0 0;color:#666666;font-family:\'Segoe UI\',Arial,sans-serif;font-size:11px;line-height:16px;text-align:left;">O conteúdo deste e-mail é confidencial e destinado ao destinatário especificado apenas na mensagem. É estritamente proibido compartilhar qualquer parte desta mensagem com terceiros, sem o consentimento por escrito do remetente. Se você recebeu esta mensagem por engano, responda a esta mensagem e siga com sua exclusão, para que possamos garantir que tal erro não ocorra no futuro.</td></tr>' +
      '<tr><td class="theme-text-muted" style="padding:12px 0 0 0;color:#666666;font-family:\'Segoe UI\',Arial,sans-serif;font-size:11px;line-height:16px;text-align:left;">The content of this email is confidential and intended for the recipient specified in message only. It is strictly forbidden to share any part of this message with any third party, without a written consent of the sender. If you received this message by mistake, please reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future.</td></tr>' +
      '</table></div>';
  }

  function setSignature(html) {
    return new Promise(function (resolve, reject) {
      try {
        var item = Office.context.mailbox.item;
        if (!item || !item.body || typeof item.body.setSignatureAsync !== "function") {
          reject(new Error("setSignatureAsync nao esta disponivel neste item."));
          return;
        }
        item.body.setSignatureAsync(
          html,
          { coercionType: Office.CoercionType.Html },
          function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(new Error(result.error && result.error.message ? result.error.message : "Falha ao inserir a assinatura."));
            }
          }
        );
      } catch (error) {
        reject(error);
      }
    });
  }

  global.NineNetSignature = Object.freeze({
    clean: clean,
    escapeHtml: escapeHtml,
    normalizeGraphProfile: normalizeGraphProfile,
    getOfficeFallbackProfile: getOfficeFallbackProfile,
    getGraphProfile: getGraphProfile,
    isPendingConfig: isPendingConfig,
    supportsNaa: supportsNaa,
    isClassicWindowsOutlook: isClassicWindowsOutlook,
    buildHtml: buildHtml,
    setSignature: setSignature
  });
})(window);
