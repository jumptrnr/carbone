const carbone = require("./lib/index.js");
const fs = require("fs");
const path = require("path");

// Read the DOCX template that contains embedded images
const templatePath = path.join(
  __dirname,
  "test",
  "datasets",
  "test_word_with_embedded_excel.docx"
);

if (!fs.existsSync(templatePath)) {
  console.error("Test DOCX file not found:", templatePath);
  console.log("Available test files:");
  const testDir = path.join(__dirname, "test", "datasets");
  if (fs.existsSync(testDir)) {
    fs.readdirSync(testDir)
      .filter((f) => f.endsWith(".docx"))
      .forEach((f) => {
        console.log("  -", f);
      });
  }
  process.exit(1);
}

console.log("Testing enhanced image processing with DOCX template...");
console.log("Template:", templatePath);

const data = {
  test: "Enhanced Image Processing Test",
  description:
    "This test demonstrates the enhanced image processing capability that captures LibreOffice-extracted images and embeds them as base64 data URIs in the HTML output.",
};

const options = {
  convertTo: "html",
};

console.log("\nStarting conversion with DEBUG enabled...");
process.env.DEBUG = "carbone:*";

carbone.render(templatePath, data, options, function (err, result) {
  if (err) {
    console.error("Conversion failed:", err);
    return;
  }

  console.log("\n✅ Conversion completed successfully!");

  // Save the result to analyze
  const outputPath = path.join(__dirname, "test_enhanced_images_output.html");
  fs.writeFileSync(outputPath, result);

  console.log("📁 HTML output saved to:", outputPath);

  // Analyze the HTML for images
  const htmlContent = result.toString();

  // Count different types of image references
  const httpImages = (
    htmlContent.match(/<img[^>]+src=["']https?:\/\/[^"']+["'][^>]*>/gi) || []
  ).length;
  const dataUriImages = (
    htmlContent.match(/<img[^>]+src=["']data:[^"']+["'][^>]*>/gi) || []
  ).length;
  const localFileImages =
    (
      htmlContent.match(
        /<img[^>]+src=["'][^"']*\.(png|jpg|jpeg|gif|svg)["'][^>]*>/gi
      ) || []
    ).length - dataUriImages;

  console.log("\n📊 Image Analysis Results:");
  console.log(
    `  📷 Total img tags: ${(htmlContent.match(/<img[^>]*>/gi) || []).length}`
  );
  console.log(`  🌐 HTTP/HTTPS images: ${httpImages}`);
  console.log(`  📦 Data URI images (embedded): ${dataUriImages}`);
  console.log(`  📁 Local file images (not embedded): ${localFileImages}`);

  if (dataUriImages > 0) {
    console.log(
      "\n🎉 SUCCESS: Found embedded images! The enhanced image processing is working."
    );
  } else if (localFileImages > 0) {
    console.log(
      "\n⚠️  PARTIAL: Still have local file references. May need further enhancement."
    );
  } else {
    console.log("\n ℹ️  INFO: No images found in this template for testing.");
  }

  // Show first few image tags for inspection
  const imgTags = htmlContent.match(/<img[^>]*>/gi) || [];
  if (imgTags.length > 0) {
    console.log("\n🔍 Sample image tags:");
    imgTags.slice(0, 3).forEach((tag, i) => {
      const src = tag.match(/src=["']([^"']+)["']/i);
      const srcValue = src ? src[1] : "no src";
      const srcType = srcValue.startsWith("data:")
        ? "DATA_URI"
        : srcValue.startsWith("http")
        ? "HTTP_URL"
        : "LOCAL_FILE";
      console.log(
        `  ${i + 1}. ${srcType}: ${
          srcValue.length > 60 ? srcValue.substring(0, 60) + "..." : srcValue
        }`
      );
    });
  }

  console.log("\n✅ Enhanced image processing test completed!");
});
