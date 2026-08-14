/**
 * Generates and triggers browser download for the populated Word document.
 * Handles JSZip 3.0+ async parsing alongside synchronous PizZip.
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
    throw new Error("Docxtemplater library is missing. Ensure docxtemplater.js is included.");
  }

  let zip;

  // 1. Try PizZip (Synchronous Constructor)
  const PizZipClass = window.PizZip || window.pizzip;
  if (PizZipClass) {
    zip = new PizZipClass(repositoryTemplateBuffer);
  } 
  // 2. Fallback to JSZip 3.0+ (Asynchronous loadAsync)
  else if (window.JSZip && typeof window.JSZip.loadAsync === 'function') {
    zip = await window.JSZip.loadAsync(repositoryTemplateBuffer);
  } else {
    throw new Error("Neither PizZip nor JSZip library could be initialized properly.");
  }

  // Render template with loaded Zip instance
  const doc = new DocxtemplaterClass(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(templateContext);

  // Generate output blob
  let outputBlob;
  if (zip.generateAsync) {
    // JSZip 3+ async blob generation
    outputBlob = await doc.getZip().generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  } else {
    // PizZip synchronous blob generation
    outputBlob = doc.getZip().generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  // Trigger browser download
  const downloadUrl = URL.createObjectURL(outputBlob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  anchor.download = `Generated_Project_Plan_${originalFileName || 'WAB.docx'}`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(downloadUrl);
}
