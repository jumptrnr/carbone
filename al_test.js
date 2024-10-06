const carbone = require("./lib/index"); // Adjust this path if necessary
const fs = require("fs");

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
  /*
  {
    
    templatePath:
      "../../freshtracksback_s4/templates/FreshTracksFundReport.odt",
    options: { convertTo: "pdf" },
    outputName: "result.pdf",
  },
  
  {
    templatePath: "../freshtracksback_s4/templates/FreshTracksFundReport.docx",
    options: { convertTo: "docx" },
    outputName: "result.docx",
  },
  
  */
  {
    templatePath: "../freshtracksback_s4/templates/FreshtracksAllCompanies.odt",
    options: { convertTo: "pdf" },
    outputName: "result.pdf",
  },
  /*
  {
    templatePath:
      "../freshtracksback_s4/templates/FreshtracksFundPresentation.odp",
    options: {
      convertTo: "pptx",
      lang: "en-us",
      reportName: "FreshTracksFundReport",
    },
    outputName: "result.pptx",
  },*/

  // Add more configurations as needed
];

// Function to handle errors
function handleError(err) {
  console.error("Error:", err);
  process.exit(1);
}

// Function to render reports
function renderReports() {
  let currentIndex = 0;

  function renderNext() {
    if (currentIndex >= renderConfigs.length) {
      console.log("All reports generated successfully.");
      process.exit(0);
      return;
    }

    const config = renderConfigs[currentIndex];

    console.log("Calling carbone.render with the following parameters:");
    console.log("Template Path:", config.templatePath);
    //console.log("Data:", JSON.stringify(data, null, 2));
    console.log("Options:", config.options);

    carbone.render(
      config.templatePath,
      data,
      config.options,
      function (err, result) {
        if (err) {
          handleError(err);
        }
        fs.writeFileSync(config.outputName, result);
        console.log(
          `${config.options.convertTo.toUpperCase()} report generated successfully: ${
            config.outputName
          }`
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
