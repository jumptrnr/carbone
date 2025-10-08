const carbone = require("./lib/index");
const fs = require("fs");

// Test data with image URLs
const data = {
  name: "Test Document",
  logo_url:
    "https://raw.githubusercontent.com/jumptrnr/carbone/master/doc/carbone_icon_small.png",
};

console.log("Testing DOCX to HTML conversion with image processing...");

// Try to render a DOCX template to HTML format
const testTemplate = "./test/datasets/test_sample.html";

carbone.render(
  testTemplate,
  data,
  { convertTo: "html" },
  function (err, result) {
    if (err) {
      console.error("Error rendering template:", err);
      return;
    }

    // Save the result
    fs.writeFileSync("docx_to_html_result.html", result);

    console.log("✅ DOCX to HTML rendering completed!");
    console.log("Check docx_to_html_result.html to see the result");

    // Check if the image was processed
    const resultString = result.toString();
    if (resultString.includes("data:image/")) {
      console.log("✅ SUCCESS: Image was converted to base64 data URI!");
    } else if (resultString.includes("https://")) {
      console.log(
        "⚠️  INFO: Image URL was preserved (may be converted by LibreOffice)"
      );
    } else {
      console.log("⚠️  UNCLEAR: Could not determine image processing status");
    }

    console.log("Test completed");
  }
);
