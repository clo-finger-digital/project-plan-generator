/**
 * Project Plan Generator - Client Engine
 * Dynamically parses uploaded WAB .docx files, detects systems/departments,
 * populates the HTML UI form, and renders the repository template.
 */

let repositoryTemplateBuffer = null;
const TEMPLATE_FILE_NAME = "Project Plan template.docx";

/**
 * Loads the project plan template directly from the GitHub repository on page load.
 */
async function loadRepositoryTemplate() {
  try {
    const response = await fetch(`./${encodeURIComponent(TEMPLATE_FILE_NAME)}`);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: Could not retrieve "${TEMPLATE_FILE_NAME}"`);
    }
    repositoryTemplateBuffer = await response.arrayBuffer();
    console.log(`[App] Successfully loaded template: "${TEMPLATE_FILE_NAME}"`);
    return true;
  } catch (error) {
    console.error("[App] Error loading repository template:", error);
    return false;
  }
}

/**
 * Parses an uploaded WAB .docx file and returns extracted parameters + detected systems list.
 * @param {File} file - Uploaded File object
 */
async function extractWabDataFromFile(file) {
  const fileBuffer = await file.arrayBuffer();
  const mammothResult = await window.mammoth.extractRawText({ arrayBuffer: fileBuffer });
  const text = mammothResult.value;

  // 1. Department / Testee Name & Abbreviation
  const deptMatch = text.match(/(?:Department|Bureau|Office)\s+of\s+[A-Za-z\s]+|Food and Environmental Hygiene Department|Education Bureau|Department of Health/i);
  const deptName = deptMatch ? deptMatch[0].trim() : "Government Department";

  const abbrMatch = text.match(/\((FEHD|DH|EDB|DPO|HKPF)\)/i) || text.match(/([A-Z]{2,6})\s+Work Assignment/);
  const deptAbbr = abbrMatch ? abbrMatch[1].toUpperCase() : (deptName.match(/\b([A-Z])/g) || []).join('');

  // 2. Extract All Systems Found in Document for Dropdown
  const detectedSystems = extractSystemsFromText(text);

  let primarySystem = detectedSystems.length > 0 ? detectedSystems[0].name : "Information System";
  let primaryAbbr = detectedSystems.length > 0 ? detectedSystems[0].abbr : "IS";

  // 3. Dynamic PIA Detection
  const hasPia = /Privacy Impact Assessment|PIA|Privacy Compliance Audit|PCA/i.test(text) &&
                 !/PIA\s+(?:is\s+)?not\s+required|No\s+PIA/i.test(text);

  // 4. Dynamic Dates
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
    TENTATIVE_COMPLETION_DATE: completionDate,
    RAW_TEXT: text
  };
}

/**
 * Extracts candidate systems and abbreviations directly from the WAB document text.
 * @param {string} text - Raw text of WAB
 * @returns {Array<{name: string, abbr: string}>} List of detected systems
 */
function extractSystemsFromText(text) {
  const systems = [];
  const seen = new Set();

  // Pattern matching system titles ending in System, Platform, Website, Service, etc.
  const regex = /([A-Z0-9\s\-\/,\(\)]+?\b(?:System|Platform|Service|Enhancement|Website)\b)/gi;
  let match;

  while ((match = regex.exec(text)) !== null) {
    let rawName = match[1].replace(/\s+/g, ' ').trim();
    
    // Filter out generic headers
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
 * Takes user-edited form data, populates the repository template, and triggers a download.
 * @param {Object} formData - Object containing user input/edited values
 * @param {string} originalFileName - Name of uploaded WAB file
 */
async function generateAndDownloadDocx(formData, originalFileName) {
  if (!repositoryTemplateBuffer) {
    throw new Error(`Repository template "${TEMPLATE_FILE_NAME}" is not loaded.`);
  }

  const currentDateStr = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });

  // Map input fields to template placeholders
  const templateContext = {
    DEPARTMENT_NAME: formData.DEPARTMENT_NAME,
    DEPARTMENT_ABBR: formData.DEPARTMENT_ABBR,
    SYSTEM_NAME: formData.SYSTEM_NAME,
    SYSTEM_ABBR: formData.SYSTEM_ABBR,
    VERSION_DATE: currentDateStr,
    TENTATIVE_COMPLETION_DATE: formData.TENTATIVE_COMPLETION_DATE,

    HAS_PIA: formData.HAS_PIA,
    SRAA_OBJECTIVES: `To perform a comprehensive Security Risk Assessment and Audit (SRAA) for ${formData.SYSTEM_NAME} in accordance with baseline security requirements.`,
    PIAA_OBJECTIVES: formData.HAS_PIA ? `To conduct a Privacy Impact Assessment (PIA) and Privacy Compliance Audit (PCA) for ${formData.SYSTEM_NAME} to ensure compliance with the Personal Data (Privacy) Ordinance.` : "",
    SRAA_SCOPE: `The scope covers security risk assessment, general control reviews, technical vulnerability scanning, and penetration testing for ${formData.SYSTEM_NAME}.`,

    DATE_STAGE_0: formData.DATE_START,
    DATE_INTRO_MEETING: formData.DATE_START,
    DATE_PROJECT_PLAN: formData.DATE_START,
    DATE_CHECKLISTS: formData.TENTATIVE_COMPLETION_DATE,
    DATE_SRAA_COMPLETION: formData.TENTATIVE_COMPLETION_DATE,
    DATE_PRESENTATION: formData.TENTATIVE_COMPLETION_DATE,
    DATE_CLOSURE: formData.TENTATIVE_COMPLETION_DATE
  };

  // Safe constructor check for PizZip & docxtemplater
  const PizZipClass = window.PizZip || window.pizzip;
  const DocxtemplaterClass = window.docxtemplater || window.Docxtemplater;

  if (!PizZipClass) {
    throw new Error("PizZip library is not defined. Please check CDN script tags.");
  }
  if (!DocxtemplaterClass) {
    throw new Error("Docxtemplater library is not defined. Please check CDN script tags.");
  }

  const zip = new PizZipClass(repositoryTemplateBuffer);
  const doc = new DocxtemplaterClass(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(templateContext);

  // Generate output blob and trigger download
  const outputBlob = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  const downloadUrl = URL.createObjectURL(outputBlob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  anchor.download = `Generated_Project_Plan_${originalFileName}`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(downloadUrl);
}
}
