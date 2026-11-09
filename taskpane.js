/* Asistente Enel para Word - versión 0.2.
   Los valores del caso se procesan en el complemento y se insertan en el documento.
   Las definiciones de campos personalizados se guardan localmente; los valores no se persisten. */

(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const officeNotice = $("officeNotice");
  const checkButton = $("checkButton");
  const applyButton = $("applyButton");
  const clearButton = $("clearButton");
  const addFieldButton = $("addFieldButton");
  const addDetectedButton = $("addDetectedButton");
  const resultsText = $("resultsText");
  const resultsList = $("resultsList");
  const customFieldsContainer = $("customFields");
  const customEmpty = $("customEmpty");
  const STORAGE_KEY = "asistenteEnel.customFields.v2";

  const builtInFields = [
    { token: "{{NOMBRE_CLIENTE}}", input: "clientName", label: "Nombre del cliente" },
    { token: "{{NUMERO_CASO}}", input: "caseNumber", label: "Número de caso" },
    { token: "{{SUMINISTRO}}", input: "supplyNumber", label: "Número de suministro" },
    { token: "{{RESPUESTA_DE_FONDO}}", input: "finalResponse", label: "Respuesta de fondo" }
  ];

  const suggestedLabels = {
    NOMBRE_CAMPO: "Nombre del campo",
    FECHA_CASO: "Fecha del caso",
    PETICION_CLIENTE: "Petición del cliente",
    PERIODO_FACTURADO: "Periodo facturado",
    META_CONSUMO: "Meta de consumo",
    CONSUMO_REGISTRADO: "Consumo registrado",
    ANALISIS_CASO: "Análisis del caso",
    RESPUESTA_FONDO: "Respuesta de fondo",
    OBSERVACIONES_INTERNAS: "Observaciones internas",
    REMISION_RESPUESTA: "Remisión de respuesta"
  };

  let officeReady = false;
  let wordAvailable = false;
  let detectedTokens = [];
  let nextId = 1;
  let customFields = loadCustomFields();

  function makeId() {
    const id = "campo_" + Date.now().toString(36) + "_" + nextId;
    nextId += 1;
    return id;
  }

  function loadCustomFields() {
    try {
      const saved = window.localStorage.getItem(STORAGE_KEY);
      if (!saved) return [];
      const parsed = JSON.parse(saved);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((item) => item && typeof item.token === "string" && typeof item.label === "string")
        .map((item) => ({ id: String(item.id || makeId()), token: item.token, label: item.label, value: "" }));
    } catch (_) {
      return [];
    }
  }

  function saveCustomFields() {
    try {
      const definitions = customFields.map(({ id, token, label }) => ({ id, token, label }));
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(definitions));
    } catch (_) {
      setNotice("No se pudo guardar la configuración local de los campos. Puedes seguir usando el formulario mientras permanezca abierto.", "warning");
    }
  }

  function setNotice(message, type) {
    officeNotice.textContent = message;
    officeNotice.className = "notice" + (type ? " " + type : "");
  }

  function createCustomField(label, token) {
    return { id: makeId(), label: label || "", token: token || "", value: "" };
  }

  function updateCustomField(id, key, value) {
    const field = customFields.find((item) => item.id === id);
    if (!field) return;
    field[key] = value;
    saveCustomFields();
  }

  function renderCustomFields(focusId) {
    customFieldsContainer.replaceChildren();
    customEmpty.hidden = customFields.length > 0;

    customFields.forEach((field, index) => {
      const card = document.createElement("article");
      card.className = "custom-field-card";

      const heading = document.createElement("div");
      heading.className = "custom-field-heading";

      const title = document.createElement("strong");
      title.textContent = "Campo personalizado " + (index + 1);

      const removeButton = document.createElement("button");
      removeButton.className = "remove-field-button";
      removeButton.type = "button";
      removeButton.textContent = "Eliminar";
      removeButton.setAttribute("aria-label", "Eliminar " + (field.label || field.token || "campo personalizado"));
      removeButton.addEventListener("click", () => {
        customFields = customFields.filter((item) => item.id !== field.id);
        saveCustomFields();
        renderCustomFields();
      });
      heading.append(title, removeButton);

      const labelCaption = document.createElement("label");
      labelCaption.textContent = "Nombre del campo";
      const labelInput = document.createElement("input");
      labelInput.type = "text";
      labelInput.autocomplete = "off";
      labelInput.placeholder = "Ej. Periodo facturado";
      labelInput.value = field.label;
      labelInput.addEventListener("input", () => updateCustomField(field.id, "label", labelInput.value));

      const tokenCaption = document.createElement("label");
      tokenCaption.textContent = "Etiqueta que escribiste en Word";
      const tokenInput = document.createElement("input");
      tokenInput.type = "text";
      tokenInput.autocomplete = "off";
      tokenInput.spellcheck = false;
      tokenInput.placeholder = "Ej. {{PERIODO_FACTURADO}}";
      tokenInput.value = field.token;
      tokenInput.addEventListener("input", () => updateCustomField(field.id, "token", tokenInput.value));

      const valueCaption = document.createElement("label");
      valueCaption.textContent = "Contenido para insertar";
      const valueInput = document.createElement("textarea");
      valueInput.rows = 3;
      valueInput.className = "custom-value";
      valueInput.dataset.fieldId = field.id;
      valueInput.placeholder = "Escribe o pega aquí el contenido…";
      valueInput.value = field.value || "";
      valueInput.addEventListener("input", () => { field.value = valueInput.value; });

      const hint = document.createElement("p");
      hint.className = "field-hint";
      hint.textContent = "La etiqueta debe coincidir exactamente con el marcador del documento, por ejemplo {{PERIODO_FACTURADO}}.";

      card.append(heading, labelCaption, labelInput, tokenCaption, tokenInput, valueCaption, valueInput, hint);
      customFieldsContainer.appendChild(card);

      if (focusId === field.id) labelInput.focus();
    });
  }

  function getConfiguredFields() {
    const configured = builtInFields.map((field) => ({
      token: field.token,
      label: field.label,
      getValue: () => $(field.input).value
    }));

    const usedTokens = new Set(builtInFields.map((field) => field.token));
    for (const field of customFields) {
      const token = field.token.trim();
      const label = field.label.trim();
      const value = field.value || "";

      // Las tarjetas totalmente vacías se ignoran hasta que el usuario las configure.
      if (!token && !label && !value.trim()) continue;
      if (!label) throw new Error("Escribe el nombre de cada campo personalizado.");
      if (!/^\{\{[A-Z0-9_]+\}\}$/.test(token)) {
        throw new Error("La etiqueta de “" + label + "” debe tener el formato {{NOMBRE_CAMPO}}, usando mayúsculas, números y guion bajo.");
      }
      if (usedTokens.has(token)) throw new Error("La etiqueta " + token + " está repetida. Cada campo debe tener una etiqueta distinta.");
      usedTokens.add(token);
      configured.push({ token, label, getValue: () => field.value || "" });
    }
    return configured;
  }

  function showCounts(counts) {
    resultsList.replaceChildren();
    const entries = Object.entries(counts);
    if (entries.length === 0) {
      resultsText.textContent = "No se encontraron marcadores con el formato {{NOMBRE_CAMPO}} en el texto principal del documento.";
      addDetectedButton.hidden = true;
      detectedTokens = [];
      return;
    }

    resultsText.textContent = "Marcadores detectados en el documento:";
    detectedTokens = entries.map(([token]) => token);

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

    const configuredTokens = new Set(builtInFields.map((field) => field.token));
    customFields.forEach((field) => {
      if (field.token.trim()) configuredTokens.add(field.token.trim());
    });
    const unconfigured = detectedTokens.filter((token) => !configuredTokens.has(token));
    addDetectedButton.hidden = unconfigured.length === 0;
    addDetectedButton.textContent = unconfigured.length === 1
      ? "Crear campo para la etiqueta nueva"
      : "Crear campos para " + unconfigured.length + " etiquetas nuevas";
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
      for (const range of matches.items) range.insertText(value, Word.InsertLocation.replace);
      if (count > 0) await context.sync();
      return count;
    });
  }

  checkButton.addEventListener("click", async () => {
    checkButton.disabled = true;
    resultsText.textContent = "Revisando el documento…";
    resultsList.replaceChildren();
    addDetectedButton.hidden = true;
    try {
      const counts = await inspectDocument();
      showCounts(counts);
      setNotice("Revisión completada.", "success");
      applyButton.disabled = !wordAvailable;
    } catch (error) {
      setNotice(error.message || "No fue posible revisar el documento.", "error");
      resultsText.textContent = "No se pudo completar la revisión.";
    } finally {
      checkButton.disabled = !wordAvailable;
    }
  });

  addFieldButton.addEventListener("click", () => {
    const field = createCustomField("", "");
    customFields.push(field);
    saveCustomFields();
    renderCustomFields(field.id);
    setNotice("Nuevo campo agregado. Define su nombre y su etiqueta tal como aparece en Word.", "success");
  });

  addDetectedButton.addEventListener("click", () => {
    const configuredTokens = new Set(builtInFields.map((field) => field.token));
    customFields.forEach((field) => {
      if (field.token.trim()) configuredTokens.add(field.token.trim());
    });

    const newTokens = detectedTokens.filter((token) => !configuredTokens.has(token));
    if (newTokens.length === 0) {
      addDetectedButton.hidden = true;
      setNotice("No hay etiquetas nuevas para agregar.", "warning");
      return;
    }

    newTokens.forEach((token) => {
      const key = token.slice(2, -2);
      const label = suggestedLabels[key] || key.toLowerCase().replace(/_/g, " ").replace(/(^|\s)\S/g, (letter) => letter.toUpperCase());
      customFields.push(createCustomField(label, token));
    });

    saveCustomFields();
    renderCustomFields();
    addDetectedButton.hidden = true;
    setNotice(newTokens.length + (newTokens.length === 1 ? " campo creado" : " campos creados") + ". Ahora diligencia el contenido y pulsa “Aplicar al documento”.", "success");
  });

  applyButton.addEventListener("click", async () => {
    let configuredFields;
    try {
      configuredFields = getConfiguredFields();
    } catch (error) {
      setNotice(error.message || "Revisa la configuración de tus campos personalizados.", "error");
      return;
    }

    const filledFields = configuredFields.filter((field) => field.getValue().trim());
    const emptyFields = configuredFields.filter((field) => !field.getValue().trim());
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
        const count = await replaceToken(field.token, field.getValue());
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
    for (const field of builtInFields) $(field.input).value = "";
    customFields.forEach((field) => { field.value = ""; });
    customFieldsContainer.querySelectorAll(".custom-value").forEach((input) => { input.value = ""; });
    setNotice("Valores del formulario limpiados. Se conservó la configuración de tus etiquetas y el documento no se modificó.", "success");
  });

  function initialize() {
    renderCustomFields();
    if (typeof Office === "undefined" || !Office.onReady) {
      setNotice("La biblioteca de Office no se ha cargado. Abre este panel desde Word y verifica la conexión.", "warning");
      return;
    }

    Office.onReady((info) => {
      officeReady = true;
      wordAvailable = !!(info && info.host === Office.HostType.Word && typeof Word !== "undefined");
      if (wordAvailable) {
        setNotice("Conectado a Word. Revisa las etiquetas del documento para crear campos personalizados automáticamente.", "success");
        checkButton.disabled = false;
        applyButton.disabled = false;
      } else {
        setNotice("El panel cargó, pero no detecta un documento de Word. Ábrelo desde Word para probar los reemplazos.", "warning");
      }
    });
  }

  checkButton.disabled = true;
  applyButton.disabled = true;
  initialize();
})();
