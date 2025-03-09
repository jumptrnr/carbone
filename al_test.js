const carbone = require("./lib/index"); // Adjust this path if necessary
const fs = require("fs");
const path = require("path");

// Import the data
const data = require("./data.js");

// Check if data is defined
if (data && data.company_data) {
  console.log("Data found: with ", data.company_data.length, " companies");
} else {
  console.error(
    "Error: 'data' is undefined or doesn't contain expected properties"
  );
  process.exit(1);
}

// Array of render configurations
const renderConfigs = [
  {
    enabled: false,
    templatePath:
      "../../freshtracksback_s4/templates/FreshTracksFundReport.odt",
    options: { convertTo: "pdf" },
    outputName: "resultFromODT.pdf",
  },
  {
    enabled: false,
    templatePath: "../freshtracksback_s4/templates/FreshTracksFundReport.docx",
    options: { convertTo: "docx" },
    outputName: "resultFromDOCX.docx",
  },
  {
    enabled: false,
    templatePath: "../freshtracksback_s4/templates/FreshtracksAllCompanies.odt",
    options: { convertTo: "pdf" },
    outputName: "resultReportFromODT.pdf",
  },
  {
    enabled: true,
    templatePath:
      "../freshtracksback_s4/templates/FreshtracksAllCompanies.docx",
    options: { convertTo: "docx" },
    outputName: "resultReportFromDOCX.docx",
  },
  {
    enabled: false,
    templatePath:
      "../freshtracksback_s4/templates/FreshtracksNewPresentationBulletTest.odp",
    options: {
      convertTo: "pdf",
      lang: "en-us",
      reportName: "FreshTracksFundReport",
    },
    outputName: "resultPresentationFromODP.pdf",
  },
  {
    enabled: true,
    templatePath: "../freshtracksback_s4/templates/FreshtracksPresentation.odp",
    options: {
      convertTo: "pptx",
      lang: "en-us",
      reportName: "FreshTracksFundReport",
    },
    outputName: "resultPresentationFromODP.pptx",
  },
  {
    enabled: false,
    templatePath: "../freshtracksback_s4/templates/FreshtracksPresentation.odp",
    options: {
      convertTo: "odp",
      lang: "en-us",
      reportName: "FreshTracksFundReport",
    },
    outputName: "resultPresentationFromODP.odp",
  },
];

// Ensure output directory exists
const outputDir = "test/output";
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// Function to handle errors
function handleError(err) {
  console.error("Error:", err);
  process.exit(1);
}

// Function to render reports
function renderReports() {
  // Filter enabled configurations
  const enabledConfigs = renderConfigs.filter((config) => config.enabled);
  let currentIndex = 0;

  function renderNext() {
    if (currentIndex >= enabledConfigs.length) {
      console.log("All reports generated successfully.");
      process.exit(0);
      return;
    }

    const config = enabledConfigs[currentIndex];
    const outputPath = path.join(outputDir, config.outputName);

    console.log("Calling carbone.render with the following parameters:");
    console.log("Template Path:", config.templatePath);
    console.log("Options:", config.options);

    carbone.render(
      config.templatePath,
      data,
      config.options,
      function (err, result) {
        if (err) {
          handleError(err);
        }
        fs.writeFileSync(outputPath, result);
        console.log(
          `${config.options.convertTo.toUpperCase()} report generated successfully: ${outputPath}`
        );
        currentIndex++;
        renderNext();
      }
    );
  }

  renderNext();
}

// Start rendering reports
renderReports();
