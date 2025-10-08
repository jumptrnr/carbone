const carbone = require("./lib/index.js");
const fs = require("fs");
const path = require("path");

// Enable debug for the postprocessor
process.env.DEBUG = "carbone:*";
process.env.DEBUG_DOCX_IMAGES = "true";

async function testImageExtraction() {
  console.log("Testing image extraction and processing...");

  // Initialize Carbone first
  console.log("Initializing Carbone...");
  await new Promise((resolve, reject) => {
    carbone.set({ factories: 1 }, (err) => {
      if (err) return reject(err);
      resolve();
    });
  });

  const templatePath = path.join(__dirname, "examples", "movies.docx");

  if (!fs.existsSync(templatePath)) {
    console.error("Template file not found:", templatePath);
    return;
  }

  console.log("Template file found:", templatePath);

  const template = fs.readFileSync(templatePath);
  const data = { movies: [{ title: "Test Movie" }] };

  const options = {
    convertTo: "html",
  };

  console.log("Starting Carbone render...");

  carbone.render(template, data, options, function (err, result) {
    if (err) {
      console.error("Error:", err);
      return;
    }

    console.log("Render successful!");

    const outputPath = path.join(__dirname, "debug_result.html");
    fs.writeFileSync(outputPath, result);

    console.log("HTML written to:", outputPath);

    // Check if images were processed
    const htmlContent = result.toString();
    const localImages = (
      htmlContent.match(/src="[^"]*\.(png|jpg|jpeg|gif|svg)"/g) || []
    ).filter((src) => !src.includes("data:"));
    const dataUriImages = htmlContent.match(/src="data:[^"]*"/g) || [];

    console.log("Local image references found:", localImages.length);
    console.log("Data URI images found:", dataUriImages.length);

    if (localImages.length > 0) {
      console.log("Local images (should be 0):");
      localImages.forEach((img, i) => console.log(`  ${i + 1}: ${img}`));
    }

    if (dataUriImages.length > 0) {
      console.log("Data URI images found:", dataUriImages.length);
    }

    process.exit(0);
  });
}

testImageExtraction().catch(console.error);
