/**
 * Project Plan Generator - Core Application Logic
 * Pure Client-Side JavaScript (GitHub Pages Compatible)
 */

document.addEventListener('DOMContentLoaded', () => {
  const TEMPLATE_PATH = './Project Plan template.docx';

  const wabInput = document.getElementById('wabInput');
  const generateBtn = document.getElementById('generateBtn');
  const templateStatus = document.getElementById('templateStatus');

  // Input Field References
  const fieldDeptName = document.getElementById('fieldDeptName');
  const fieldDeptAbbr = document.getElementById('fieldDeptAbbr');
  const fieldSystemName = document.getElementById('fieldSystemName');
  const fieldSystemAbbr = document.getElementById('fieldSystemAbbr');
  const fieldStartDate = document.getElementById('fieldStartDate');

  let parsedWabContext = null;
  let templateArrayBuffer = null;

  // 1. Automatically fetch the template stored in the repository on startup
  async function loadRepoTemplate() {
    try {
      const response = await fetch(TEMPLATE_PATH);
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      templateArrayBuffer = await response.arrayBuffer();
      if (templateStatus) {
        templateStatus.className = "status-badge ready";
        templateStatus.innerText = "Repository template loaded successfully.";
      }
      checkReadyState();
    } catch (err) {
      console.error("Failed to load project plan template:", err);
      if (templateStatus) {
        templateStatus.className = "status-badge error";
        templateStatus.innerText = "Error loading template from repository. Check file path.";
      }
    }
  }

  loadRepoTemplate();

  // Enable button only when WAB is uploaded AND template is loaded
  function checkReadyState() {
    generateBtn.disabled = !(wabInput.files.length > 0 && templateArrayBuffer !== null);
  }

  // WAB File Upload Handler
  wabInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (file) {
      try {
        parsedWabContext = await parseWabDocument(file);
        autoPopulateUI(parsedWabContext);
        checkReadyState();
      } catch (err) {
        console.error("Error parsing WAB Document:", err);
        alert("Failed to parse the WAB file. You can manually enter the fields below.");
      }
    } else {
      checkReadyState();
    }
  });

  // Main Generation Handler
  generateBtn.addEventListener('click', async () => {
    if (!templateArrayBuffer) {
      alert("Repository template is not loaded.");
      return;
    }

    try {
      generateBtn.disabled = true;
      generateBtn.innerText = "Generating Document...";

      // Build Template Payload by combining UI Inputs & Extracted Context
      const payload = prepareTemplatePayload(parsedWabContext);

      // Render Document using Docxtemplater with loaded template
      const zip = new PizZip(templateArrayBuffer);
      const doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });

      doc.render(payload);

      // Generate Download
      const outputBlob = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      const sysAbbr = fieldSystemAbbr.value.trim() || "PROJECT";
      const downloadUrl = URL.createObjectURL(outputBlob);
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = `Project_Plan_${sysAbbr}_${formatDateForFileName(new Date())}.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

    } catch (error) {
      console.error("Error generating Project Plan:", error);
      alert("An error occurred while generating the document. Please ensure template tags are valid.");
    } finally {
      generateBtn.disabled = false;
      generateBtn.innerText = "Generate Project Plan (.docx)";
    }
  });

  // ==========================================
  // WAB PARSING ENGINE & PATTERN EXTRACTION
  // ==========================================

  async function parseWabDocument(file) {
    const arrayBuffer = await file.arrayBuffer();
    const mammothResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
    const rawText = mammothResult.value;

    // 1. Service Type Detection
    const hasPia = /Privacy Impact Assessment|PIA\b|PIAA\b|Privacy Compliance Audit|PCA\b/i.test(rawText);
    const hasSraa = /Security Risk Assessment|SRAA\b|Security Audit|SRA\b/i.test(rawText);

    // 2. Department & Bureau Extraction
    let deptName = extractFirstMatch(rawText, [
      /(?:Civil Service Bureau|CSB)/i ? "Civil Service Bureau" : null,
      /(?:Immigration Department|ImmD)/i ? "Immigration Department" : null,
      /(?:Education Bureau|EDB)/i ? "Education Bureau" : null,
      /for\s+the\s+([A-Z][A-Za-z\s]+(?:Bureau|Department|Office))/i,
      /Services for\s+([A-Z][A-Za-z\s]+(?:Bureau|Department|Office))/i,
      /Bureaux\/Departments\s*\(\s*([A-Za-z\s]+)\s*\)/i
    ]);

    if (!deptName) {
      if (/Civil Service Bureau|CSB\b/i.test(rawText)) deptName = "Civil Service Bureau";
      else if (/Immigration Department|ImmD\b/i.test(rawText)) deptName = "Immigration Department";
      else if (/Education Bureau|EDB\b/i.test(rawText)) deptName = "Education Bureau";
    }

    let deptAbbr = extractFirstMatch(rawText, [
      /(?:Bureau|Department|Office)\s*\(\s*["'“]?([A-Z]{2,6})["'”]?\s*\)/i,
      /\b(CSB|ImmD|EDB|C&ED|HKPF|ArchSD|HD)\b/
    ]);

    if (!deptAbbr && deptName) {
      if (deptName.includes("Civil Service Bureau")) deptAbbr = "CSB";
      else if (deptName.includes("Immigration Department")) deptAbbr = "ImmD";
      else if (deptName.includes("Education Bureau")) deptAbbr = "EDB";
    }

    // 3. System Name & Abbreviation Extraction
    let systemName = "";
    let systemAbbr = "";

    const sysMatch = rawText.match(/(?:for|of)\s+(?:the\s+)?([A-Za-z0-9\s\-\(\)\/\,\.\:\;\']+?)\s*\((?:the\s+["“]?)?([A-Z0-9\-]{2,15})(?:["”]?\s*System)?\)/i) ||
                     rawText.match(/System[:\s]+([A-Za-z0-9\s\-]+?)\s*\(([A-Z0-9\-]+)\)/i);

    if (sysMatch) {
      systemName = sysMatch[1].trim();
      systemAbbr = sysMatch[2].trim();
    } else {
      systemName = extractFirstMatch(rawText, [
        /in relation to\s+["'“]([^"'”]+)["'”]/i,
        /for\s+the\s+["'“]?([A-Z0-9\s-]+System)["'”]?/i,
        /for\s+["'“]([^"'”]+)["'”]/i
      ]);

      systemAbbr = extractFirstMatch(rawText, [
        /["'“][^"'”]+["'”]\s*\(\s*["'“]?([A-Z0-9-]{2,12})["'”]?\s*\)/i,
        /\b(ICONS|AOPIS|HRMIS|LMP|BLNST|SEMIS|BSPP|IISS|QR)\b/
      ]);
    }

    // 4. Objectives and Scope Extraction (No default fallback text)
    const sraaObjectives = hasSraa ? extractSectionText(
      rawText,
      /(?:objectives of|objective).*?(?:security risk assessment|sraa|audit)/i,
      /project requirements|scope of services|user requirements|2\./i
    ) : "";

    const piaObjectives = hasPia ? extractSectionText(
      rawText,
      /(?:objectives of|objective).*?(?:privacy impact assessment|pia|piaa)/i,
      /project requirements|scope of services|user requirements|2\./i
    ) : "";

    const sraaScope = hasSraa ? extractSectionText(
      rawText,
      /(?:scope of the services|scope of services|service scope)/i,
      /background|project objectives|3\./i
    ) : "";

    // 5. Start Date Determination (Returns null if not found)
    const startDate = extractStartDateAnchor(rawText);

    return {
      rawText,
      hasPia,
      hasSraa,
      deptName,
      deptAbbr,
      systemName,
      systemAbbr,
      sraaObjectives,
      piaObjectives,
      sraaScope,
      startDate
    };
  }

  // Auto-populate UI Input Textboxes (Only if extracted values exist)
  function autoPopulateUI(ctx) {
    if (ctx.deptName) fieldDeptName.value = ctx.deptName;
    if (ctx.deptAbbr) fieldDeptAbbr.value = ctx.deptAbbr;
    if (ctx.systemName) fieldSystemName.value = ctx.systemName;
    if (ctx.systemAbbr) fieldSystemAbbr.value = ctx.systemAbbr;
    if (ctx.startDate) fieldStartDate.value = ctx.startDate.toISOString().split('T')[0];
  }

  // Construct Document Tokens
  function prepareTemplatePayload(parsedCtx) {
    const deptName = fieldDeptName.value.trim();
    const deptAbbr = fieldDeptAbbr.value.trim();
    const systemName = fieldSystemName.value.trim();
    const systemAbbr = fieldSystemAbbr.value.trim();

    return {
      DEPARTMENT_NAME: deptName,
      DEPARTMENT_ABBR: deptAbbr,
      SYSTEM_NAME: systemName,
      SYSTEM_ABBR: systemAbbr,

      HAS_SRAA: parsedCtx ? parsedCtx.hasSraa : false,
      HAS_PIA: parsedCtx ? parsedCtx.hasPia : false,

      SRAA_OBJECTIVES: parsedCtx ? parsedCtx.sraaObjectives : "",
      PIAA_OBJECTIVES: parsedCtx ? parsedCtx.piaObjectives : "",
      SRAA_SCOPE: parsedCtx ? parsedCtx.sraaScope : "",

      // Date of Creation for VERSION_DATE; rest left as blank/null
      VERSION_DATE: formatDate(new Date()),
      START_DATE_MONTH: "",

      DATE_STAGE_0: "",
      DATE_INTRO_MEETING: "",
      DATE_PROJECT_PLAN: "",
      DATE_CHECKLISTS: "",
      DATE_SRAA_COMPLETION: "",
      DATE_PRESENTATION: "",
      DATE_CLOSURE: "",
      TENTATIVE_COMPLETION_DATE: ""
    };
  }

  // ==========================================
  // HELPER FUNCTIONS
  // ==========================================

  function extractStartDateAnchor(text) {
    const match = text.match(/(?:tentative start date|commence on|conducted from|start date)[^.]*?is\s+([A-Za-z]+\s+\d{4}|\d{1,2}\s+[A-Za-z]+\s+\d{4}|[A-Za-z]+\s+\d{1,2}\,?\s+\d{4})/i);
    if (match && match[1]) {
      const parsed = new Date(match[1]);
      if (!isNaN(parsed.getTime())) return parsed;
    }
    return null;
  }

  function extractFirstMatch(text, regexArray) {
    for (const item of regexArray) {
      if (!item) continue;
      if (typeof item === 'string') return item;
      const match = text.match(item);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
    return "";
  }

  function extractSectionText(text, startRegex, endRegex) {
    const startIndex = text.search(startRegex);
    if (startIndex === -1) return "";

    const subText = text.substring(startIndex);
    const endIndex = subText.search(endRegex);

    const result = endIndex !== -1 ? subText.substring(0, endIndex) : subText;
    const cleaned = result.replace(startRegex, '').trim();
    return cleaned.length > 10 ? cleaned : "";
  }

  function formatDate(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  function formatDateForFileName(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${yyyy}-${mm}-${dd}`;
  }
});/**
 * Project Plan Generator - Core Application Logic
 * Pure Client-Side JavaScript (GitHub Pages Compatible)
 */

document.addEventListener('DOMContentLoaded', () => {
  const TEMPLATE_PATH = './Project Plan template.docx';

  const wabInput = document.getElementById('wabInput');
  const generateBtn = document.getElementById('generateBtn');
  const templateStatus = document.getElementById('templateStatus');

  // Input Field References
  const fieldDeptName = document.getElementById('fieldDeptName');
  const fieldDeptAbbr = document.getElementById('fieldDeptAbbr');
  const fieldSystemName = document.getElementById('fieldSystemName');
  const fieldSystemAbbr = document.getElementById('fieldSystemAbbr');
  const fieldStartDate = document.getElementById('fieldStartDate');

  let parsedWabContext = null;
  let templateArrayBuffer = null;

  // 1. Automatically fetch the template stored in the repository on startup
  async function loadRepoTemplate() {
    try {
      const response = await fetch(TEMPLATE_PATH);
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      templateArrayBuffer = await response.arrayBuffer();
      templateStatus.className = "status-badge ready";
      templateStatus.innerText = "Repository template loaded successfully.";
      checkReadyState();
    } catch (err) {
      console.error("Failed to load project plan template:", err);
      templateStatus.className = "status-badge error";
      templateStatus.innerText = "Error loading template from repository. Check file path.";
    }
  }

  loadRepoTemplate();

  // Enable button only when WAB is uploaded AND template is loaded
  function checkReadyState() {
    generateBtn.disabled = !(wabInput.files.length > 0 && templateArrayBuffer !== null);
  }

  // WAB File Upload Handler
  wabInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (file) {
      try {
        parsedWabContext = await parseWabDocument(file);
        autoPopulateUI(parsedWabContext);
        checkReadyState();
      } catch (err) {
        console.error("Error parsing WAB Document:", err);
        alert("Failed to parse the WAB file. You can manually enter the fields below.");
      }
    } else {
      checkReadyState();
    }
  });

  // Main Generation Handler
  generateBtn.addEventListener('click', async () => {
    if (!templateArrayBuffer) {
      alert("Repository template is not loaded.");
      return;
    }

    try {
      generateBtn.disabled = true;
      generateBtn.innerText = "Generating Document...";

      // Build Template Payload by combining UI Inputs & Extracted Context
      const payload = prepareTemplatePayload(parsedWabContext);

      // Render Document using Docxtemplater with loaded template
      const zip = new PizZip(templateArrayBuffer);
      const doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });

      doc.render(payload);

      // Generate Download
      const outputBlob = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      const sysAbbr = fieldSystemAbbr.value.trim() || "PROJECT";
      const downloadUrl = URL.createObjectURL(outputBlob);
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = `Project_Plan_${sysAbbr}_${formatDateForFileName(new Date())}.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

    } catch (error) {
      console.error("Error generating Project Plan:", error);
      alert("An error occurred while generating the document. Please ensure template tags are valid.");
    } finally {
      generateBtn.disabled = false;
      generateBtn.innerText = "Generate Project Plan (.docx)";
    }
  });

  // ==========================================
  // WAB PARSING ENGINE & PATTERN EXTRACTION
  // ==========================================

  async function parseWabDocument(file) {
    const arrayBuffer = await file.arrayBuffer();
    const mammothResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
    const rawText = mammothResult.value;

    // 1. Detect Service Types
    const hasPia = /Privacy Impact Assessment|PIA|PIAA|Privacy Compliance Audit/i.test(rawText);
    const hasSraa = /Security Risk Assessment|SRAA|Security Audit|SRA/i.test(rawText);

    // 2. Extract Metadata with Heuristic Fallback Arrays
    const deptName = extractFirstMatch(rawText, [
      /(?:for|Services for the|Services for)\s+([A-Z][A-Za-z\s]+(?:Bureau|Department|Office))/i,
      /(Education Bureau|Immigration Department|Civil Service Bureau|Customs and Excise Department)/i,
      /Bureaux\/Departments\s*\(([A-Za-z\s]+)\)/i
    ], "");

    const deptAbbr = extractFirstMatch(rawText, [
      /(?:Bureau|Department|Office)\s*\(\s*["'“]?([A-Z]{2,6})["'”]?\s*\)/i,
      /\b(EDB|ImmD|CSB|C&ED|HKPF|ArchSD|HD)\b/
    ], "");

    const systemName = extractFirstMatch(rawText, [
      /in relation to\s+["'“]([^"'”]+)["'”]/i,
      /for the\s+["'“]?([A-Z0-9\s-]+System|Immigration Control System|Coaching Programme on Executive Skills)["'”]?/i,
      /for\s+["'“]([^"'”]+)["'”]/i
    ], "");

    const systemAbbr = extractFirstMatch(rawText, [
      /["'“][^"'”]+["'”]\s*\(\s*["'“]?([A-Z0-9-]{2,12})["'”]?\s*\)/i,
      /\b(e-COPES|ICONS|Seamless|eBS)\b/
    ], "");

    // 3. Extract Scope & Objectives Sections
    const sraaObjectives = hasSraa ? extractSectionText(
      rawText,
      /objectives of.*?(?:security risk assessment|sraa)/i,
      /project requirements|scope of services|user requirements/i,
      "To evaluate security risks, verify compliance with Baseline IT Security Policy (S17) and IT Security Guidelines (G3), and recommend security safeguards."
    ) : "";

    const piaObjectives = hasPia ? extractSectionText(
      rawText,
      /objectives of.*?(?:privacy impact assessment|pia)/i,
      /project requirements|scope of services|user requirements/i,
      "To identify data privacy risks, perform data processing cycle analysis, ensure compliance with the Personal Data (Privacy) Ordinance (Cap. 486), and recommend privacy protective measures."
    ) : "";

    const sraaScope = hasSraa ? extractSectionText(
      rawText,
      /scope of the services|scope of services/i,
      /background|project objectives/i,
      "Conduct general control review, vulnerability scanning, penetration testing, configuration review, source code review, risk assessment, and security compliance audit."
    ) : "";

    // 4. Determine Start Date Anchor
    const startDate = extractStartDateAnchor(rawText);

    return {
      rawText,
      hasPia,
      hasSraa,
      deptName,
      deptAbbr,
      systemName,
      systemAbbr,
      sraaObjectives,
      piaObjectives,
      sraaScope,
      startDate
    };
  }

  // Auto-populate UI Input Textboxes
  function autoPopulateUI(ctx) {
    if (ctx.deptName) fieldDeptName.value = ctx.deptName;
    if (ctx.deptAbbr) fieldDeptAbbr.value = ctx.deptAbbr;
    if (ctx.systemName) fieldSystemName.value = ctx.systemName;
    if (ctx.systemAbbr) fieldSystemAbbr.value = ctx.systemAbbr;
    if (ctx.startDate) fieldStartDate.value = ctx.startDate.toISOString().split('T')[0];
  }

  // Construct Document Tokens
  function prepareTemplatePayload(parsedCtx) {
    const deptName = fieldDeptName.value.trim() || "Government Bureau/Department";
    const deptAbbr = fieldDeptAbbr.value.trim() || "B/D";
    const systemName = fieldSystemName.value.trim() || "Target Information System";
    const systemAbbr = fieldSystemAbbr.value.trim() || "SYSTEM";

    const baseDateInput = fieldStartDate.value ? new Date(fieldStartDate.value) : new Date();
    const baseDate = isNaN(baseDateInput.getTime()) ? new Date() : baseDateInput;

    const schedule = generateCalculatedSchedule(baseDate);

    return {
      DEPARTMENT_NAME: deptName,
      DEPARTMENT_ABBR: deptAbbr,
      SYSTEM_NAME: systemName,
      SYSTEM_ABBR: systemAbbr,

      HAS_SRAA: parsedCtx ? parsedCtx.hasSraa : true,
      HAS_PIA: parsedCtx ? parsedCtx.hasPia : true,

      SRAA_OBJECTIVES: parsedCtx ? parsedCtx.sraaObjectives : "",
      PIAA_OBJECTIVES: parsedCtx ? parsedCtx.piaObjectives : "",
      SRAA_SCOPE: parsedCtx ? parsedCtx.sraaScope : "",

      ...schedule,
      TENTATIVE_COMPLETION_DATE: schedule.DATE_CLOSURE
    };
  }

  // ==========================================
  // SCHEDULE & HELPER FUNCTIONS
  // ==========================================

  function generateCalculatedSchedule(baseDate) {
    const addDays = (d, days) => {
      const res = new Date(d);
      res.setDate(res.getDate() + days);
      return formatDate(res);
    };
    const addWeeks = (d, weeks) => addDays(d, weeks * 7);

    return {
      VERSION_DATE: formatDate(new Date()),
      START_DATE_MONTH: baseDate.toLocaleString('default', { month: 'long', year: 'numeric' }),

      DATE_STAGE_0: addWeeks(baseDate, 1),
      DATE_INTRO_MEETING: addDays(baseDate, 10),
      DATE_PROJECT_PLAN: addWeeks(baseDate, 2),
      DATE_CHECKLISTS: addWeeks(baseDate, 3),
      DATE_SRAA_COMPLETION: addWeeks(baseDate, 8),
      DATE_PRESENTATION: addWeeks(baseDate, 9),
      DATE_CLOSURE: addWeeks(baseDate, 12)
    };
  }

  function extractStartDateAnchor(text) {
    const match = text.match(/(?:tentative start date|commence on|conducted from)[^.]*?is\s+([A-Za-z]+\s+\d{4}|\d{1,2}\s+[A-Za-z]+\s+\d{4}|[A-Za-z]+\s+\d{4})/i);
    if (match && match[1]) {
      const parsed = new Date(match[1]);
      if (!isNaN(parsed.getTime())) return parsed;
    }
    return new Date();
  }

  function extractFirstMatch(text, regexArray, fallback) {
    for (const regex of regexArray) {
      const match = text.match(regex);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
    return fallback;
  }

  function extractSectionText(text, startRegex, endRegex, fallback) {
    const startIndex = text.search(startRegex);
    if (startIndex === -1) return fallback;

    const subText = text.substring(startIndex);
    const endIndex = subText.search(endRegex);

    const result = endIndex !== -1 ? subText.substring(0, endIndex) : subText;
    const cleaned = result.replace(startRegex, '').trim();
    return cleaned.length > 20 ? cleaned : fallback;
  }

  function formatDate(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  function formatDateForFileName(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${yyyy}-${mm}-${dd}`;
  }
});/**
 * Project Plan Generator - Core Application Logic
 * Pure Client-Side JavaScript (GitHub Pages Compatible)
 */

document.addEventListener('DOMContentLoaded', () => {
  const wabInput = document.getElementById('wabInput');
  const templateInput = document.getElementById('templateInput');
  const generateBtn = document.getElementById('generateBtn');

  // Input Field References
  const fieldDeptName = document.getElementById('fieldDeptName');
  const fieldDeptAbbr = document.getElementById('fieldDeptAbbr');
  const fieldSystemName = document.getElementById('fieldSystemName');
  const fieldSystemAbbr = document.getElementById('fieldSystemAbbr');
  const fieldStartDate = document.getElementById('fieldStartDate');

  let parsedWabContext = null;

  // Track File Uploads & Enable Button
  const checkFilesReady = () => {
    generateBtn.disabled = !(wabInput.files.length > 0 && templateInput.files.length > 0);
  };

  wabInput.addEventListener('change', async (event) => {
    checkFilesReady();
    const file = event.target.files[0];
    if (file) {
      try {
        parsedWabContext = await parseWabDocument(file);
        autoPopulateUI(parsedWabContext);
      } catch (err) {
        console.error("Error parsing WAB Document:", err);
        alert("Failed to parse the WAB file. You can manually enter the fields below.");
      }
    }
  });

  templateInput.addEventListener('change', checkFilesReady);

  // Main Generation Handler
  generateBtn.addEventListener('click', async () => {
    const templateFile = templateInput.files[0];
    if (!templateFile) {
      alert("Please upload a Project Plan Template (.docx)");
      return;
    }

    try {
      generateBtn.disabled = true;
      generateBtn.innerText = "Generating Document...";

      // Build Template Payload by combining UI Inputs & Extracted Context
      const payload = prepareTemplatePayload(parsedWabContext);

      // Render Document using Docxtemplater
      const templateArrayBuffer = await templateFile.arrayBuffer();
      const zip = new PizZip(templateArrayBuffer);
      const doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });

      doc.render(payload);

      // Generate Download
      const outputBlob = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      const sysAbbr = fieldSystemAbbr.value.trim() || "PROJECT";
      const downloadUrl = URL.createObjectURL(outputBlob);
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = `Project_Plan_${sysAbbr}_${formatDateForFileName(new Date())}.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

    } catch (error) {
      console.error("Error generating Project Plan:", error);
      alert("An error occurred while generating the document. Please ensure template tags are valid.");
    } finally {
      generateBtn.disabled = false;
      generateBtn.innerText = "Generate Project Plan (.docx)";
    }
  });

  // ==========================================
  // WAB PARSING ENGINE & PATTERN EXTRACTION
  // ==========================================

  async function parseWabDocument(file) {
    const arrayBuffer = await file.arrayBuffer();
    const mammothResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
    const rawText = mammothResult.value;

    // 1. Detect Service Types
    const hasPia = /Privacy Impact Assessment|PIA|PIAA|Privacy Compliance Audit/i.test(rawText);
    const hasSraa = /Security Risk Assessment|SRAA|Security Audit|SRA/i.test(rawText);

    // 2. Extract Metadata with Heuristic Fallback Arrays
    const deptName = extractFirstMatch(rawText, [
      /(?:for|Services for the|Services for)\s+([A-Z][A-Za-z\s]+(?:Bureau|Department|Office))/i,
      /(Education Bureau|Immigration Department|Civil Service Bureau|Customs and Excise Department)/i,
      /Bureaux\/Departments\s*\(([A-Za-z\s]+)\)/i
    ], "");

    const deptAbbr = extractFirstMatch(rawText, [
      /(?:Bureau|Department|Office)\s*\(\s*["'“]?([A-Z]{2,6})["'”]?\s*\)/i,
      /\b(EDB|ImmD|CSB|C&ED|HKPF|ArchSD|HD)\b/
    ], "");

    const systemName = extractFirstMatch(rawText, [
      /in relation to\s+["'“]([^"'”]+)["'”]/i,
      /for the\s+["'“]?([A-Z0-9\s-]+System|Immigration Control System|Coaching Programme on Executive Skills)["'”]?/i,
      /for\s+["'“]([^"'”]+)["'”]/i
    ], "");

    const systemAbbr = extractFirstMatch(rawText, [
      /["'“][^"'”]+["'”]\s*\(\s*["'“]?([A-Z0-9-]{2,12})["'”]?\s*\)/i,
      /\b(e-COPES|ICONS|Seamless|eBS)\b/
    ], "");

    // 3. Extract Scope & Objectives Sections
    const sraaObjectives = hasSraa ? extractSectionText(
      rawText,
      /objectives of.*?(?:security risk assessment|sraa)/i,
      /project requirements|scope of services|user requirements/i,
      "To evaluate security risks, verify compliance with Baseline IT Security Policy (S17) and IT Security Guidelines (G3), and recommend security safeguards."
    ) : "";

    const piaObjectives = hasPia ? extractSectionText(
      rawText,
      /objectives of.*?(?:privacy impact assessment|pia)/i,
      /project requirements|scope of services|user requirements/i,
      "To identify data privacy risks, perform data processing cycle analysis, ensure compliance with the Personal Data (Privacy) Ordinance (Cap. 486), and recommend privacy protective measures."
    ) : "";

    const sraaScope = hasSraa ? extractSectionText(
      rawText,
      /scope of the services|scope of services/i,
      /background|project objectives/i,
      "Conduct general control review, vulnerability scanning, penetration testing, configuration review, source code review, risk assessment, and security compliance audit."
    ) : "";

    // 4. Determine Start Date Anchor
    const startDate = extractStartDateAnchor(rawText);

    return {
      rawText,
      hasPia,
      hasSraa,
      deptName,
      deptAbbr,
      systemName,
      systemAbbr,
      sraaObjectives,
      piaObjectives,
      sraaScope,
      startDate
    };
  }

  // Auto-populate UI Input Textboxes
  function autoPopulateUI(ctx) {
    if (ctx.deptName) fieldDeptName.value = ctx.deptName;
    if (ctx.deptAbbr) fieldDeptAbbr.value = ctx.deptAbbr;
    if (ctx.systemName) fieldSystemName.value = ctx.systemName;
    if (ctx.systemAbbr) fieldSystemAbbr.value = ctx.systemAbbr;
    if (ctx.startDate) fieldStartDate.value = ctx.startDate.toISOString().split('T')[0];
  }

  // Construct Document Tokens
  function prepareTemplatePayload(parsedCtx) {
    const deptName = fieldDeptName.value.trim() || "Government Bureau/Department";
    const deptAbbr = fieldDeptAbbr.value.trim() || "B/D";
    const systemName = fieldSystemName.value.trim() || "Target Information System";
    const systemAbbr = fieldSystemAbbr.value.trim() || "SYSTEM";

    const baseDateInput = fieldStartDate.value ? new Date(fieldStartDate.value) : new Date();
    const baseDate = isNaN(baseDateInput.getTime()) ? new Date() : baseDateInput;

    const schedule = generateCalculatedSchedule(baseDate);

    return {
      // Direct UI & Extraction Tokens
      DEPARTMENT_NAME: deptName,
      DEPARTMENT_ABBR: deptAbbr,
      SYSTEM_NAME: systemName,
      SYSTEM_ABBR: systemAbbr,

      // Flags
      HAS_SRAA: parsedCtx ? parsedCtx.hasSraa : true,
      HAS_PIA: parsedCtx ? parsedCtx.hasPia : true,

      // Sections
      SRAA_OBJECTIVES: parsedCtx ? parsedCtx.sraaObjectives : "",
      PIAA_OBJECTIVES: parsedCtx ? parsedCtx.piaObjectives : "",
      SRAA_SCOPE: parsedCtx ? parsedCtx.sraaScope : "",

      // Dynamic Milestones
      ...schedule,
      TENTATIVE_COMPLETION_DATE: schedule.DATE_CLOSURE
    };
  }

  // ==========================================
  // SCHEDULE & HELPER FUNCTIONS
  // ==========================================

  function generateCalculatedSchedule(baseDate) {
    const addDays = (d, days) => {
      const res = new Date(d);
      res.setDate(res.getDate() + days);
      return formatDate(res);
    };
    const addWeeks = (d, weeks) => addDays(d, weeks * 7);

    return {
      VERSION_DATE: formatDate(new Date()),
      START_DATE_MONTH: baseDate.toLocaleString('default', { month: 'long', year: 'numeric' }),

      DATE_STAGE_0: addWeeks(baseDate, 1),
      DATE_INTRO_MEETING: addDays(baseDate, 10),
      DATE_PROJECT_PLAN: addWeeks(baseDate, 2),
      DATE_CHECKLISTS: addWeeks(baseDate, 3),
      DATE_SRAA_COMPLETION: addWeeks(baseDate, 8),
      DATE_PRESENTATION: addWeeks(baseDate, 9),
      DATE_CLOSURE: addWeeks(baseDate, 12)
    };
  }

  function extractStartDateAnchor(text) {
    const match = text.match(/(?:tentative start date|commence on|conducted from)[^.]*?is\s+([A-Za-z]+\s+\d{4}|\d{1,2}\s+[A-Za-z]+\s+\d{4}|[A-Za-z]+\s+\d{4})/i);
    if (match && match[1]) {
      const parsed = new Date(match[1]);
      if (!isNaN(parsed.getTime())) return parsed;
    }
    return new Date(); // Fallback: Current Date
  }

  function extractFirstMatch(text, regexArray, fallback) {
    for (const regex of regexArray) {
      const match = text.match(regex);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
    return fallback;
  }

  function extractSectionText(text, startRegex, endRegex, fallback) {
    const startIndex = text.search(startRegex);
    if (startIndex === -1) return fallback;

    const subText = text.substring(startIndex);
    const endIndex = subText.search(endRegex);

    const result = endIndex !== -1 ? subText.substring(0, endIndex) : subText;
    const cleaned = result.replace(startRegex, '').trim();
    return cleaned.length > 20 ? cleaned : fallback;
  }

  function formatDate(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  function formatDateForFileName(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${yyyy}-${mm}-${dd}`;
  }
});
