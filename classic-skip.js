/*
 * O Outlook classico no Windows carrega este arquivo pela secao Override do
 * manifesto. A assinatura classica continua sendo administrada pela GPO e,
 * portanto, este manipulador encerra o evento sem inserir uma segunda copia.
 */
(function () {
  "use strict";

  function checkSignature(event) {
    event.completed();
  }

  Office.actions.associate("checkSignature", checkSignature);
})();

