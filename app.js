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
  const deptName = deptMatch ? deptMatch[0].trim() : "Government Department";

  const abbrMatch = text.match(/\((FEHD|DH|EDB|DPO|HKPF)\)/i) || text.match(/([A-Z]{2,6})\s+Work Assignment/);
  const deptAbbr = abbrMatch ? abbrMatch[1].toUpperCase() : (deptName.match(/\b([A-Z])/g) || []).join('');

  // 2. Systems Detection
  const detectedSystems = extractSystemsFromText(text);
  const primarySystem = detectedSystems.length > 0 ? detectedSystems[0].name : "Information System";
  const primaryAbbr = detectedSystems.length > 0 ? detectedSystems[0].abbr : "IS";

  // 3. Privacy Scope (HAS_PIA)
  const hasPia = /Privacy Impact Assessment|PIA|Privacy Compliance Audit|PCA/i.test(text) &&
                 !/PIA\s+(?:is\s+)?not\s+required|No\s+PIA/i.test(text);

  // 4. Dates
  const dates = text.match(/\b(?:\d{1,2}\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\b/gi) || [];
  const completionDate = dates.length > 0 ? dates[dates.length - 1] : "October 2026";
  const startDate = dates.length > 1 ? dates[0] : "July 2026";

  // 5. Dynamic Scope & Objectives Extraction directly from WAB
  const extractedSraaObjectives = extractSraaObjectives(text, primarySystem);
  const extractedPiaaObjectives = hasPia ? extractPiaaObjectives(text, primarySystem) : "";
  const extractedSraaScope = extractSraaScope(text, primarySystem);

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
 * Strictly searches ONLY inside Section 3 (PROJECT OBJECTIVES) and stops before Section 4 (PROJECT REQUIREMENTS).
 * Extracts all point-form items starting with "It is to..." or "To...".
 */
function extractSraaObjectives(text, systemName) {
  // Isolate Section 3 strictly between "PROJECT OBJECTIVES" and Section 4 ("PROJECT REQUIREMENTS")
  const section3Match = text.match(/PROJECT OBJECTIVES([\s\S]*?)(?=PROJECT REQUIREMENTS|4\.\s*PROJECT REQUIREMENTS|$)/i) ||
                        text.match(/3\.\s*PROJECT OBJECTIVES([\s\S]*?)(?=4\.\s*|$)/i);

  if (section3Match && section3Match[1]) {
    const rawObjectivesText = section3Match[1].trim();

    // Regex to split items in point form starting with "It is to..." or "To..."
    const itemsRegex = /(?:\([a-z0-9]+\)|\d+\.|[•\-\*])?\s*((?:It is\s+)?to\s+[\s\S]*?)(?=(?:\([a-z0-9]+\)|\d+\.|[•\-\*]|\b(?:It is\s+)?to\b|$))/gi;
    const matches = [];
    let match;

    while ((match = itemsRegex.exec(rawObjectivesText)) !== null) {
      let cleanItem = match[1].replace(/\s+/g, ' ').trim();

      // Normalize string so every point starts with "It is to..."
      if (!cleanItem.toLowerCase().startsWith('it is to')) {
        if (cleanItem.toLowerCase().startsWith('to ')) {
          cleanItem = 'It is ' + cleanItem;
        } else {
          cleanItem = 'It is to ' + cleanItem;
        }
      }

      // Ignore short trailing text fragments
      if (cleanItem.length > 20 && !matches.includes(cleanItem)) {
        matches.push(cleanItem);
      }
    }

    if (matches.length > 0) {
      // Return point form: (a) It is to... \n\n (b) It is to...
      return matches.map((item, index) => {
        const letter = String.fromCharCode(97 + index); // 'a', 'b', 'c'...
        return `(${letter}) ${item}`;
      }).join('\n\n');
    }
  }

  // Fallback default
  return `(a) It is to evaluate the security risks of ${systemName} and related data of the department.\n\n(b) It is to determine the level of compliance with government IT security requirements (S17, G3).\n\n(c) It is to verify after implementation of safeguards that identified risks have been mitigated or reduced to an acceptable level.`;
}

/**
 * Extracts PIAA Objectives directly from WAB text.
 */
function extractPiaaObjectives(text, systemName) {
  const match = text.match(/The objectives of PIA services are:[\s\S]*?(?=PROJECT REQUIREMENTS|4\.)/i) ||
                text.match(/Privacy Impact Assessment.*?Objectives?[\s\S]*?(?=Scope|Requirements)/i);

  if (match) {
    let rawObj = match[0]
      .replace(/The objectives of PIA services are:/i, '')
      .trim();
    if (rawObj.length > 20) return rawObj;
  }
  return `To conduct a Privacy Impact Assessment (PIA) and Privacy Compliance Audit (PCA) for ${systemName} to ensure compliance with the Personal Data (Privacy) Ordinance.`;
}

/**
 * Extracts SRAA Scope directly from WAB text.
 */
function extractSraaScope(text, systemName) {
  const match = text.match(/SCOPE OF THE SERVICES[\s\S]*?(?=BACKGROUND|PROJECT OBJECTIVES|2\.)/i) ||
                text.match(/Scope of Service[\s\S]*?(?=Approach|Objectives)/i);

  if (match) {
    let rawScope = match[0]
      .replace(/SCOPE OF THE SERVICES/i, '')
      .replace(/Scope of Service/i, '')
      .trim();
    if (rawScope.length > 30) return rawScope;
  }
  return `To assess the overall information security level by evaluating the security risks of ${systemName} and related data, identifying recommended safeguards to strengthen system protection.`;
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
    DEPARTMENT_NAME: formData.DEPARTMENT_NAME || "Government Department",
    DEPARTMENT_ABBR: formData.DEPARTMENT_ABBR || "GOV",
    SYSTEM_NAME: formData.SYSTEM_NAME || "Target System",
    SYSTEM_ABBR: formData.SYSTEM_ABBR || "TS",
    VERSION_DATE: currentDateStr,
    TENTATIVE_COMPLETION_DATE: formData.TENTATIVE_COMPLETION_DATE || "October 2026",

    HAS_PIA: Boolean(formData.HAS_PIA),
    SRAA_OBJECTIVES: formData.SRAA_OBJECTIVES,
    PIAA_OBJECTIVES: formData.HAS_PIA ? formData.PIAA_OBJECTIVES : "",
    SRAA_SCOPE: formData.SRAA_SCOPE,

    DATE_STAGE_0: formData.DATE_START || "July 2026",
    DATE_INTRO_MEETING: formData.DATE_START || "July 2026",
    DATE_PROJECT_PLAN: formData.DATE_START || "July 2026",
    DATE_CHECKLISTS: formData.TENTATIVE_COMPLETION_DATE || "October 2026",
    DATE_SRAA_COMPLETION: formData.TENTATIVE_COMPLETION_DATE || "October 2026",
    DATE_PRESENTATION: formData.TENTATIVE_COMPLETION_DATE || "October 2026",
    DATE_CLOSURE: formData.TENTATIVE_COMPLETION_DATE || "October 2026"
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
