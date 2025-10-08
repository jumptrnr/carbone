const carbone = require("./lib/index");
const fs = require("fs");

// Create a test HTML template with an image URL
const htmlTemplate = `<html>
  <body>
    <h1>Test Document</h1>
    <img src="https://raw.githubusercontent.com/jumptrnr/carbone/master/doc/carbone_icon_small.png" alt="Carbone Logo">
    <p>This is a test document with an image from {d.name}.</p>
  </body>
</html>`;

// Save the template
fs.writeFileSync("test_template.html", htmlTemplate);

// Data for the template
const data = {
  name: "Carbone Template Engine",
};

console.log("Testing HTML to HTML rendering with image processing...");

// Render the template
carbone.render("./test_template.html", data, function (err, result) {
  if (err) {
    console.error("Error rendering template:", err);
    return;
  }

  // Save the result
  fs.writeFileSync("test_result.html", result);

  console.log("✅ HTML template rendered successfully!");
  console.log("Check test_result.html to see the result with embedded image");

  // Check if the image was processed
  const resultString = result.toString();
  if (resultString.includes("data:image/")) {
    console.log("✅ SUCCESS: Image was converted to base64 data URI!");
  } else if (resultString.includes("https://")) {
    console.log("❌ FAILED: Image URL was not processed");
  } else {
    console.log("⚠️  UNCLEAR: Could not determine image processing status");
  }

  // Clean up
  fs.unlinkSync("test_template.html");
  console.log("Cleanup completed");
});
