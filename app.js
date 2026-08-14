/**
 * Project Plan Generator - Client-Side Engine (GitHub Pages)
 * Synchronizes with index.html to parse uploaded WAB .docx files and 
 * populate "Project Plan template.docx" fetched directly from the repo.
 */

// Global reference for the fetched template ArrayBuffer
let repositoryTemplateBuffer = null;
const TEMPLATE_FILE_NAME = "Project Plan template.docx";

/**
 * Loads the project plan template directly from the GitHub repository on page load.
 * @returns {Promise<boolean>} True if successfully loaded, false otherwise.
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
 * Dynamically extracts parameters from raw text extracted from uploaded WAB files.
 * @param {string} text - Plain text extracted from the WAB document.
 * @returns {Object} Template data context mapping to template placeholders.
 */
function extractDynamicWabData(text) {
  // 1. Dynamic Department & Abbreviation Extraction
  const deptMatch = text.match(/(?:Department|Bureau|Office)\s+of\s+[A-Za-z\s]+|Food and Environmental Hygiene Department|Education Bureau|Department of Health/i);
  const deptName = deptMatch ? deptMatch[0].trim() : "Government Department";

  const abbrMatch = text.match(/\((FEHD|DH|EDB|DPO|HKPF)\)/i) || text.match(/([A-Z]{2,6})\s+Work Assignment/);
  const deptAbbr = abbrMatch ? abbrMatch[1].toUpperCase() : (deptName.match(/\b([A-Z])/g) || []).join('');

  // 2. Dynamic System Name & Abbreviation Extraction
  let systemName = "Information System";
  let systemAbbr = "IS";

  const systemMatch = text.match(/(?:for|of)\s+(?:the\s+)?([A-Z0-9\s\-\/,\(\)]+?\b(?:System|Platform|Service|Enhancement|Website)\b)/i);
  if (systemMatch) {
    systemName = systemMatch[1].replace(/\s+/g, ' ').trim();
    const abbrInSys = systemName.match(/\(([^)]+)\)/);
    systemAbbr = abbrInSys ? abbrInSys[1] : (systemName.match(/\b([A-Z0-9])/g) || []).join('');
  }

  // 3. Dynamic PIA Scope Detection (HAS_PIA)
  const hasPia = /Privacy Impact Assessment|PIA|Privacy Compliance Audit|PCA/i.test(text) &&
                 !/PIA\s+(?:is\s+)?not\s+required|No\s+PIA/i.test(text);

  // 4. Dynamic Schedule & Date Calculations
  const dates = text.match(/\b(?:\d{1,2}\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\b/gi) || [];
  const currentDateStr = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  const completionDate = dates.length > 0 ? dates[dates.length - 1] : "October 2026";
  const startDate = dates.length > 1 ? dates[0] : "July 2026";

  // 5. Dynamic Scope & Objectives Extraction
  const sraaObjMatch = text.match(/Objectives?[\s\S]*?(?=Scope|2\.|3\.)/i);
  const sraaScopeMatch = text.match(/Scope\s+of\s+Service[\s\S]*?(?=Approach|3\.|4\.)/i);

  const sraaObjectives = sraaObjMatch 
    ? sraaObjMatch[0].replace(/Objectives?/i, '').trim().slice(0, 500) 
    : `To perform a comprehensive Security Risk Assessment and Audit (SRAA) for ${systemName} in accordance with baseline security requirements.`;

  const sraaScope = sraaScopeMatch 
    ? sraaScopeMatch[0].replace(/Scope\s+of\s+Service/i, '').trim().slice(0, 600) 
    : `The scope covers security risk assessment, general control reviews, technical vulnerability scanning, and penetration testing for ${systemName}.`;

  const piaaObjectives = hasPia 
    ? `To conduct a Privacy Impact Assessment (PIA) and Privacy Compliance Audit (PCA) for ${systemName} to ensure compliance with the Personal Data (Privacy) Ordinance.`
    : "";

  return {
    DEPARTMENT_NAME: deptName,
    DEPARTMENT_ABBR: deptAbbr,
    SYSTEM_NAME: systemName,
    SYSTEM_ABBR: systemAbbr,
    VERSION_DATE: currentDateStr,
    TENTATIVE_COMPLETION_DATE: completionDate,

    HAS_PIA: hasPia,
    SRAA_OBJECTIVES: sraaObjectives,
    PIAA_OBJECTIVES: piaaObjectives,
    SRAA_SCOPE: sraaScope,

    DATE_STAGE_0: startDate,
    DATE_INTRO_MEETING: startDate,
    DATE_PROJECT_PLAN: startDate,
    DATE_CHECKLISTS: completionDate,
    DATE_SRAA_COMPLETION: completionDate,
    DATE_PRESENTATION: completionDate,
    DATE_CLOSURE: completionDate
  };
}

/**
 * Processes an uploaded WAB .docx file and merges it with the repository template.
 * @param {File} file - Uploaded File object from index.html input.
 */
async function processUploadedWab(file) {
  if (!repositoryTemplateBuffer) {
    throw new Error(`Repository template "${TEMPLATE_FILE_NAME}" is not loaded yet. Please ensure the file exists in the repository root.`);
  }

  // 1. Read uploaded DOCX file as ArrayBuffer
  const fileBuffer = await file.arrayBuffer();

  // 2. Extract plain text using mammoth
  const mammothResult = await window.mammoth.extractRawText({ arrayBuffer: fileBuffer });
  const rawText = mammothResult.value;

  // 3. Extract dynamic parameters from raw text
  const templateData = extractDynamicWabData(rawText);

  // 4. Populate template using PizZip and docxtemplater
  const zip = new window.PizZip(repositoryTemplateBuffer);
  const doc = new window.docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(templateData);

  // 5. Generate output binary blob and initiate direct browser download
  const outputBlob = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  const downloadUrl = URL.createObjectURL(outputBlob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  anchor.download = `Generated_Project_Plan_${file.name}`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(downloadUrl);
}
