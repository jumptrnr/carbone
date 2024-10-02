const carbone = require("carbone");
const fs = require("fs");

function generateReport(templatePath, data, options) {
  return new Promise((resolve, reject) => {
    carbone.render(templatePath, data, options, (err, result) => {
      if (err) reject(err);
      else resolve(result);
    });
  });
}

async function main() {
  const data = {
    firstname: "John",
    lastname: "Doe",
  };

  const options = {
    convertTo: "pdf", //can be docx, txt, ...
  };

  try {
    const result = await generateReport(
      "./node_modules/carbone/examples/simple.odt",
      data,
      options
    );
    fs.writeFileSync("result.pdf", result);
    console.log("Report generated successfully: result.pdf");
  } catch (err) {
    console.error("Error generating report:", err);
  } finally {
    process.exit(); // to kill automatically LibreOffice workers
  }
}

main();
