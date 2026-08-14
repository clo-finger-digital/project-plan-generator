/**
 * Project Plan Generator - Client-Side Engine (GitHub Pages)
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
 * @param {File} file - Uploaded File object
 * @returns {Promise<Object>} Extracted metadata & detected system list
 */
async function extractWabDataFromFile(file) {
  if (!file) {
    throw new Error("No file provided for extraction.");
  }

  const mammothEngine = window.mammoth;
  if (!mammothEngine || typeof mammothEngine.extractRawText !== 'function') {
    throw new Error("Mammoth library is not available. Ensure mammoth script tag is included.");
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

  return {
    DEPARTMENT_NAME: deptName,
    DEPARTMENT_ABBR: deptAbbr,
    SYSTEM_NAME: primarySystem,
    SYSTEM_ABBR: primaryAbbr,
    DETECTED_SYSTEMS: detectedSystems,
    HAS_PIA: hasPia,
    DATE_START: startDate,
    TENTATIVE_COMPLETION_DATE: completionDate
  };
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
 * Generates and triggers browser download for the populated Word document.
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
    SRAA_OBJECTIVES: `To perform a comprehensive Security Risk Assessment and Audit (SRAA) for ${formData.SYSTEM_NAME} in accordance with baseline security requirements.`,
    PIAA_OBJECTIVES: formData.HAS_PIA ? `To conduct a Privacy Impact Assessment (PIA) and Privacy Compliance Audit (PCA) for ${formData.SYSTEM_NAME} to ensure compliance with the Personal Data (Privacy) Ordinance.` : "",
    SRAA_SCOPE: `The scope covers security risk assessment, general control reviews, technical vulnerability scanning, and penetration testing for ${formData.SYSTEM_NAME}.`,

    DATE_STAGE_0: formData.DATE_START || "July 2026",
    DATE_INTRO_MEETING: formData.DATE_START || "July 2026",
    DATE_PROJECT_PLAN: formData.DATE_START || "July 2026",
    DATE_CHECKLISTS: formData.TENTATIVE_COMPLETION_DATE || "October 2026",
    DATE_SRAA_COMPLETION: formData.TENTATIVE_COMPLETION_DATE || "October 2026",
    DATE_PRESENTATION: formData.TENTATIVE_COMPLETION_DATE || "October 2026",
    DATE_CLOSURE: formData.TENTATIVE_COMPLETION_DATE || "October 2026"
  };

  const DocxtemplaterClass = window.docxtemplater || window.Docxtemplater;
  if (!DocxtemplaterClass) {
    throw new Error("Docxtemplater library missing. Ensure docxtemplater.js is included.");
  }

  let zip;
  const PizZipClass = window.PizZip || window.pizzip;

  if (PizZipClass) {
    zip = new PizZipClass(repositoryTemplateBuffer);
  } else if (window.JSZip && typeof window.JSZip.loadAsync === 'function') {
    zip = await window.JSZip.loadAsync(repositoryTemplateBuffer);
  } else {
    throw new Error("Neither PizZip nor JSZip library could be initialized properly.");
  }

  const doc = new DocxtemplaterClass(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(templateContext);

  let outputBlob;
  if (zip.generateAsync) {
    outputBlob = await doc.getZip().generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  } else {
    outputBlob = doc.getZip().generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  const downloadUrl = URL.createObjectURL(outputBlob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  anchor.download = `Generated_Project_Plan_${originalFileName || 'WAB.docx'}`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(downloadUrl);
}

// Explicitly export to window object
window.loadRepositoryTemplate = loadRepositoryTemplate;
window.extractWabDataFromFile = extractWabDataFromFile;
window.generateAndDownloadDocx = generateAndDownloadDocx;
