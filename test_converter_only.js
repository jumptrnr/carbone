const fs = require("fs");
const path = require("path");

// Try to convert a simple DOCX to HTML and capture the converter output
const carbone = require("./lib/index.js");

// Enable debug for converter specifically
process.env.DEBUG = "carbone:converter";

console.log("Testing DOCX to HTML conversion specifically...");

async function testConversion() {
  // Initialize Carbone
  await new Promise((resolve, reject) => {
    carbone.set({ factories: 1 }, (err) => {
      if (err) return reject(err);
      resolve();
    });
  });

  const templatePath = path.join(__dirname, "examples", "movies.docx");
  const template = fs.readFileSync(templatePath);

  console.log("Starting DOCX to HTML conversion...");

  carbone.render(
    template,
    { movies: [] },
    { convertTo: "html" },
    (err, result) => {
      if (err) {
        console.error("Error:", err);
        return;
      }

      console.log("Conversion completed successfully");

      // Write the result and analyze it
      fs.writeFileSync("/tmp/test_output.html", result);

      const htmlContent = result.toString();
      console.log("HTML length:", htmlContent.length);

      // Look for image references
      const localImages =
        htmlContent.match(/src="[^"]*\.(png|jpg|jpeg|gif|svg)"/g) || [];
      const dataUriImages = htmlContent.match(/src="data:[^"]*"/g) || [];

      console.log("Local image references:", localImages.length);
      console.log("Data URI images:", dataUriImages.length);

      if (localImages.length > 0) {
        console.log("Sample local images:");
        localImages.slice(0, 3).forEach((img) => console.log(" ", img));
      }

      process.exit(0);
    }
  );
}

testConversion().catch(console.error);
