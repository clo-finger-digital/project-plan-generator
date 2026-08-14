/**
 * Project Plan Generator - Client Engine (GitHub Pages)
 * Extracts dynamic data from uploaded WAB .docx files and populates
 * "Project Plan template.docx".
 */

let repositoryTemplateBuffer = null;
const TEMPLATE_FILE_NAME = "Project Plan template.docx";

/**
 * Safely fetches the template file from the repository root on startup.
 */
async function loadRepositoryTemplate() {
  try {
    const response = await fetch(`./${encodeURIComponent(TEMPLATE_FILE_NAME)}`);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: Failed to fetch "${TEMPLATE_FILE_NAME}"`);
    }
    repositoryTemplateBuffer = await response.arrayBuffer();
    console.log(`[App] Successfully cached repository template: ${TEMPLATE_FILE_NAME}`);
    return true;
  } catch (error) {
    console.error("[App] Template load error:", error);
    return false;
  }
}

/**
 * Extracts raw text from an uploaded WAB .docx file using mammoth.
 */
async function extractWabDataFromFile(file) {
  if (!file) {
    throw new Error("No file provided for extraction.");
  }

  const mammothEngine = window.mammoth;
  if (!mammothEngine || typeof mammothEngine.extractRawText !== 'function') {
    throw new Error("Mammoth library is not loaded. Ensure mammoth.browser.min.js script tag is present.");
  }

  const fileBuffer = await file.arrayBuffer();
  const mammothResult = await mammothEngine.extractRawText({ arrayBuffer: fileBuffer });
  const text = mammothResult.value || "";

  // 1. Department / Testee Name & Abbreviation
  const deptMatch = text.match(/(?:Department|Bureau|Office)\s+of\s+[A-Za-z\s]+|Food and Environmental Hygiene Department|Education Bureau|Department of Health/i);
  const deptName = deptMatch ? deptMatch[0].trim() : "";

  const abbrMatch = text.match(/\((FEHD|DH|EDB|DPO|HKPF)\)/i) || text.match(/([A-Z]{2,6})\s+Work Assignment/);
  const deptAbbr = abbrMatch ? abbrMatch[1].toUpperCase() : (deptName ? (deptName.match(/\b([A-Z])/g) || []).join('') : "");

  // 2. Systems Detection
  const detectedSystems = extractSystemsFromText(text);
  const primarySystem = detectedSystems.length > 0 ? detectedSystems[0].name : "";
  const primaryAbbr = detectedSystems.length > 0 ? detectedSystems[0].abbr : "";

  // 3. Privacy Scope (HAS_PIA)
  const hasPia = /Privacy Impact Assessment|PIA|Privacy Compliance Audit|PCA/i.test(text) &&
                 !/PIA\s+(?:is\s+)?not\s+required|No\s+PIA/i.test(text);

  // 4. Dates
  const dates = text.match(/\b(?:\d{1,2}\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\b/gi) || [];
  const completionDate = dates.length > 0 ? dates[dates.length - 1] : "";
  const startDate = dates.length > 1 ? dates[0] : "";

  // 5. Dynamic Scope & Objectives Extraction directly from WAB (No generic fallbacks)
  const extractedSraaObjectives = extractSraaObjectives(text);
  const extractedPiaaObjectives = hasPia ? extractPiaaObjectives(text) : "";
  const extractedSraaScope = extractSraaScope(text);

  return {
    DEPARTMENT_NAME: deptName,
    DEPARTMENT_ABBR: deptAbbr,
    SYSTEM_NAME: primarySystem,
    SYSTEM_ABBR: primaryAbbr,
    DETECTED_SYSTEMS: detectedSystems,
    HAS_PIA: hasPia,
    DATE_START: startDate,
    TENTATIVE_COMPLETION_DATE: completionDate,
    SRAA_OBJECTIVES: extractedSraaObjectives,
    PIAA_OBJECTIVES: extractedPiaaObjectives,
    SRAA_SCOPE: extractedSraaScope
  };
}

/**
 * Robustly extracts ALL items starting with "It is to" verbatim inside Section 3 (PROJECT OBJECTIVES).
 * Stops strictly before Section 4 (PROJECT REQUIREMENTS).
 */
function extractSraaObjectives(text) {
  // Isolate Section 3 between "PROJECT OBJECTIVES" and Section 4 ("PROJECT REQUIREMENTS")
  const section3Match = text.match(/PROJECT OBJECTIVES[\s\S]*?(?=PROJECT REQUIREMENTS|4\.\s*PROJECT REQUIREMENTS|$)/i);

  if (!section3Match) return "";

  const rawSectionText = section3Match[0];

  // Split section text using boundary matching on "It is to"
  const matches = [];
  const parts = rawSectionText.split(/\b(?=It is to\b)/i);

  for (let part of parts) {
    let cleanPart = part.replace(/\s+/g, ' ').trim();

    // Clean leading bullets or numbers if present
    cleanPart = cleanPart.replace(/^(?:\([a-z0-9]+\)|\d+\.|[•\-\*])\s*/i, '');

    if (/^It is to\b/i.test(cleanPart)) {
      // Remove trailing section headers or non-objective artifacts if captured at the end
      cleanPart = cleanPart.replace(/PROJECT REQUIREMENTS.*$/i, '').trim();

      if (cleanPart.length > 15 && !matches.includes(cleanPart)) {
        matches.push(cleanPart);
      }
    }
  }

  if (matches.length > 0) {
    // Format into lettered point form verbatim: (a) It is to... \n\n (b) It is to...
    return matches.map((item, index) => {
      const letter = String.fromCharCode(97 + index); // 'a', 'b', 'c'...
      return `(${letter}) ${item}`;
    }).join('\n\n');
  }

  return "";
}

/**
 * Extracts PIAA Objectives directly from WAB text.
 */
function extractPiaaObjectives(text) {
  const match = text.match(/The objectives of PIA services are:[\s\S]*?(?=PROJECT REQUIREMENTS|4\.)/i) ||
                text.match(/Privacy Impact Assessment.*?Objectives?[\s\S]*?(?=Scope|Requirements)/i);

  if (match) {
    let rawObj = match[0]
      .replace(/The objectives of PIA services are:/i, '')
      .trim();
    if (rawObj.length > 20) return rawObj;
  }

  return "";
}

/**
 * Extracts SRAA Scope directly from WAB text.
 */
function extractSraaScope(text) {
  const match = text.match(/SCOPE OF THE SERVICES[\s\S]*?(?=BACKGROUND|PROJECT OBJECTIVES|2\.)/i) ||
                text.match(/Scope of Service[\s\S]*?(?=Approach|Objectives)/i);

  if (match) {
    let rawScope = match[0]
      .replace(/SCOPE OF THE SERVICES/i, '')
      .replace(/Scope of Service/i, '')
      .trim();
    if (rawScope.length > 30) return rawScope;
  }

  return "";
}

/**
 * Parses candidate systems from plain text.
 */
function extractSystemsFromText(text) {
  const systems = [];
  const seen = new Set();
  const regex = /([A-Z0-9\s\-\/,\(\)]+?\b(?:System|Platform|Service|Enhancement|Website)\b)/gi;
  let match;

  while ((match = regex.exec(text)) !== null) {
    const rawName = match[1].replace(/\s+/g, ' ').trim();
    if (rawName.length > 3 && rawName.length < 120 && !rawName.toLowerCase().startsWith('the scope')) {
      const abbrMatch = rawName.match(/\(([^)]+)\)/);
      const abbr = abbrMatch ? abbrMatch[1] : (rawName.match(/\b([A-Z0-9])/g) || []).join('');
      
      if (!seen.has(rawName)) {
        seen.add(rawName);
        systems.push({ name: rawName, abbr: abbr });
      }
    }
  }

  return systems;
}

/**
 * Generates and downloads the populated Word document.
 */
async function generateAndDownloadDocx(formData, originalFileName) {
  if (!repositoryTemplateBuffer) {
    throw new Error(`Repository template "${TEMPLATE_FILE_NAME}" is not loaded yet.`);
  }

  const currentDateStr = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });

  const templateContext = {
    DEPARTMENT_NAME: formData.DEPARTMENT_NAME || "",
    DEPARTMENT_ABBR: formData.DEPARTMENT_ABBR || "",
    SYSTEM_NAME: formData.SYSTEM_NAME || "",
    SYSTEM_ABBR: formData.SYSTEM_ABBR || "",
    VERSION_DATE: currentDateStr,
    TENTATIVE_COMPLETION_DATE: formData.TENTATIVE_COMPLETION_DATE || "",

    HAS_PIA: Boolean(formData.HAS_PIA),
    SRAA_OBJECTIVES: formData.SRAA_OBJECTIVES || "",
    PIAA_OBJECTIVES: formData.HAS_PIA ? (formData.PIAA_OBJECTIVES || "") : "",
    SRAA_SCOPE: formData.SRAA_SCOPE || "",

    DATE_STAGE_0: formData.DATE_START || "",
    DATE_INTRO_MEETING: formData.DATE_START || "",
    DATE_PROJECT_PLAN: formData.DATE_START || "",
    DATE_CHECKLISTS: formData.TENTATIVE_COMPLETION_DATE || "",
    DATE_SRAA_COMPLETION: formData.TENTATIVE_COMPLETION_DATE || "",
    DATE_PRESENTATION: formData.TENTATIVE_COMPLETION_DATE || "",
    DATE_CLOSURE: formData.TENTATIVE_COMPLETION_DATE || ""
  };

  const PizZipConstructor = window.PizZip || window.pizzip || (window.docxtemplater && window.docxtemplater.PizZip);
  const DocxtemplaterConstructor = window.docxtemplater || window.Docxtemplater;

  if (!PizZipConstructor) throw new Error("PizZip engine missing. Verify CDN script tags.");
  if (!DocxtemplaterConstructor) throw new Error("Docxtemplater engine missing. Verify CDN script tags.");

  const zip = new PizZipConstructor(repositoryTemplateBuffer);
  const doc = new DocxtemplaterConstructor(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{{', end: '}}' }
  });

  try {
    doc.render(templateContext);
  } catch (error) {
    if (error.properties && error.properties.errors instanceof Array) {
      const errorDetails = error.properties.errors.map(e => {
        const tag = e.properties && e.properties.id ? ` Tag: "${e.properties.id}"` : '';
        const explanation = e.properties && e.properties.explanation ? e.properties.explanation : e.message;
        return `${explanation}${tag}`;
      }).join(' | ');
      throw new Error(`Template Format Error: ${errorDetails}`);
    }
    throw error;
  }

  const outputBlob = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  const downloadUrl = URL.createObjectURL(outputBlob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  anchor.download = `Generated_Project_Plan_${originalFileName || 'WAB.docx'}`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(downloadUrl);
}

// Global Exports
window.loadRepositoryTemplate = loadRepositoryTemplate;
window.extractWabDataFromFile = extractWabDataFromFile;
window.generateAndDownloadDocx = generateAndDownloadDocx;
