/* Asistente Enel para Word - versión inicial de prueba.
   Los valores se procesan en el complemento y se insertan en el documento abierto.
   Esta versión no transmite el contenido del documento ni los campos a un servidor. */

(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const officeNotice = $("officeNotice");
  const checkButton = $("checkButton");
  const applyButton = $("applyButton");
  const clearButton = $("clearButton");
  const resultsText = $("resultsText");
  const resultsList = $("resultsList");

  const fields = [
    { token: "{{NOMBRE_CLIENTE}}", input: "clientName", label: "Nombre del cliente" },
    { token: "{{NUMERO_CASO}}", input: "caseNumber", label: "Número de caso" },
    { token: "{{SUMINISTRO}}", input: "supplyNumber", label: "Número de suministro" },
    { token: "{{RESPUESTA_DE_FONDO}}", input: "finalResponse", label: "Respuesta de fondo" }
  ];

  let officeReady = false;
  let wordAvailable = false;

  function setNotice(message, type) {
    officeNotice.textContent = message;
    officeNotice.className = "notice" + (type ? " " + type : "");
  }

  function getValues() {
    const values = {};
    for (const field of fields) values[field.token] = $(field.input).value;
    return values;
  }

  function showCounts(counts) {
    resultsList.replaceChildren();
    const entries = Object.entries(counts);
    if (entries.length === 0) {
      resultsText.textContent = "No se encontraron marcadores con el formato {{NOMBRE_CAMPO}} en el texto principal del documento.";
      return;
    }
    resultsText.textContent = "Marcadores detectados en el documento:";
    for (const [token, count] of entries) {
      const li = document.createElement("li");
      const tokenSpan = document.createElement("span");
      const countSpan = document.createElement("span");
      tokenSpan.className = "result-token";
      tokenSpan.textContent = token;
      countSpan.className = "result-count";
      countSpan.textContent = count + (count === 1 ? " coincidencia" : " coincidencias");
      li.append(tokenSpan, countSpan);
      resultsList.appendChild(li);
    }
  }

  async function inspectDocument() {
    if (!wordAvailable) throw new Error("Abre este complemento desde un documento de Word para utilizar esta función.");
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      const text = body.text || "";
      const matches = text.match(/\{\{[A-Z0-9_]+\}\}/g) || [];
      const counts = {};
      for (const token of matches) counts[token] = (counts[token] || 0) + 1;
      return counts;
    });
  }

  async function replaceToken(token, value) {
    return Word.run(async (context) => {
      const matches = context.document.body.search(token, { matchCase: true, matchWildcards: false });
      matches.load("items");
      await context.sync();
      const count = matches.items.length;
      for (const range of matches.items) {
        range.insertText(value, Word.InsertLocation.replace);
      }
      if (count > 0) await context.sync();
      return count;
    });
  }

  checkButton.addEventListener("click", async () => {
    checkButton.disabled = true;
    resultsText.textContent = "Revisando el documento…";
    resultsList.replaceChildren();
    try {
      const counts = await inspectDocument();
      showCounts(counts);
      setNotice("Revisión completada.", "success");
      applyButton.disabled = !wordAvailable;
    } catch (error) {
      setNotice(error.message || "No fue posible revisar el documento.", "error");
      resultsText.textContent = "No se pudo completar la revisión.";
    } finally {
      checkButton.disabled = false;
    }
  });

  applyButton.addEventListener("click", async () => {
    const values = getValues();
    const emptyFields = fields.filter((field) => !values[field.token].trim());
    const filledFields = fields.filter((field) => values[field.token].trim());

    if (filledFields.length === 0) {
      setNotice("Diligencia al menos un campo antes de aplicar los cambios.", "warning");
      return;
    }

    applyButton.disabled = true;
    checkButton.disabled = true;
    let total = 0;
    const applied = [];
    const missing = [];

    try {
      for (const field of filledFields) {
        const count = await replaceToken(field.token, values[field.token]);
        total += count;
        if (count > 0) applied.push(field.token + " (" + count + ")");
        else missing.push(field.token);
      }

      const skipped = emptyFields.map((field) => field.token);
      let message = total + (total === 1 ? " marcador reemplazado." : " marcadores reemplazados.");
      if (applied.length) message += " Aplicados: " + applied.join(", ") + ".";
      if (missing.length) message += " No encontrados: " + missing.join(", ") + ".";
      if (skipped.length) message += " Campos vacíos omitidos: " + skipped.join(", ") + ".";

      setNotice(message, total > 0 ? "success" : "warning");
      const counts = await inspectDocument();
      showCounts(counts);
    } catch (error) {
      setNotice(error.message || "Ocurrió un error al reemplazar los marcadores.", "error");
    } finally {
      applyButton.disabled = !wordAvailable;
      checkButton.disabled = !wordAvailable;
    }
  });

  clearButton.addEventListener("click", () => {
    for (const field of fields) $(field.input).value = "";
    setNotice("Formulario limpio. El documento no ha sido modificado.");
  });

  function initialize() {
    if (typeof Office === "undefined" || !Office.onReady) {
      setNotice("La biblioteca de Office no se ha cargado. Abre este panel desde Word y verifica la conexión.", "warning");
      return;
    }

    Office.onReady((info) => {
      officeReady = true;
      wordAvailable = !!(info && info.host === Office.HostType.Word && typeof Word !== "undefined");
      if (wordAvailable) {
        setNotice("Conectado a Word. Abre una plantilla con marcadores y pulsa “Revisar etiquetas”.", "success");
        checkButton.disabled = false;
        applyButton.disabled = false;
      } else {
        setNotice("El panel cargó, pero no detecta un documento de Word. Ábrelo desde Word para probar los reemplazos.", "warning");
      }
    });
  }

  // Desactiva acciones hasta confirmar que el complemento está dentro de Word.
  checkButton.disabled = true;
  applyButton.disabled = true;
  initialize();
})();
