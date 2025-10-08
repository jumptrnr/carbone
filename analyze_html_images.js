const fs = require("fs");
const path = require("path");

// Script to analyze and potentially fix HTML image references
function analyzeHtmlImages(htmlFilePath) {
  console.log(`\n🔍 Analyzing HTML file: ${htmlFilePath}`);

  if (!fs.existsSync(htmlFilePath)) {
    console.error(`❌ HTML file not found: ${htmlFilePath}`);
    return;
  }

  const htmlContent = fs.readFileSync(htmlFilePath, "utf8");

  // Find all image references
  const imgTagRegex = /<img\s+[^>]*src=["']([^"']+)["'][^>]*>/gi;
  const images = [];
  let match;

  while ((match = imgTagRegex.exec(htmlContent)) !== null) {
    const src = match[1];
    images.push({
      fullTag: match[0],
      src: src,
      isHttp: src.startsWith("http://") || src.startsWith("https://"),
      isLocal:
        !src.startsWith("http://") &&
        !src.startsWith("https://") &&
        !src.startsWith("data:"),
    });
  }

  console.log(`\n📊 Image Analysis Results:`);
  console.log(`Total images found: ${images.length}`);

  const httpImages = images.filter((img) => img.isHttp);
  const localImages = images.filter((img) => img.isLocal);
  const dataUriImages = images.filter((img) => img.src.startsWith("data:"));

  console.log(`- HTTP/HTTPS images: ${httpImages.length}`);
  console.log(`- Local file references: ${localImages.length}`);
  console.log(`- Data URI images: ${dataUriImages.length}`);

  if (localImages.length > 0) {
    console.log(`\n⚠️  Local image files that won't display in browsers:`);
    localImages.forEach((img, index) => {
      console.log(`${index + 1}. ${img.src}`);
    });

    console.log(
      `\n💡 These images were likely extracted by LibreOffice during DOCX→HTML conversion.`
    );
    console.log(
      `   They exist as separate files but aren't embedded in the HTML.`
    );
    console.log(`   Solutions:`);
    console.log(
      `   1. Host these images on a web server and update the HTML references`
    );
    console.log(
      `   2. Convert them to data URIs (base64) for inline embedding`
    );
    console.log(
      `   3. Enhance Carbone to handle this conversion automatically`
    );
  }

  if (httpImages.length > 0) {
    console.log(`\n✅ HTTP/HTTPS images (should work if URLs are accessible):`);
    httpImages.forEach((img, index) => {
      console.log(`${index + 1}. ${img.src}`);
    });
  }

  return {
    total: images.length,
    http: httpImages.length,
    local: localImages.length,
    dataUri: dataUriImages.length,
    localImagePaths: localImages.map((img) => img.src),
  };
}

// Analyze your generated HTML file
const htmlPath =
  "/Users/al/Development/carbone/test/output/resultFromDOCX.html";
const analysis = analyzeHtmlImages(htmlPath);

console.log(`\n📋 Summary for ${path.basename(htmlPath)}:`);
console.log(
  `The HTML contains ${analysis.local} local image references that won't display.`
);
console.log(
  `These images need to be converted to data URIs or hosted externally.`
);

if (analysis.local > 0) {
  console.log(`\n🔧 Next steps to fix image display:`);
  console.log(`1. The images were embedded in your original DOCX file`);
  console.log(
    `2. LibreOffice extracted them during conversion but they're not in the HTML output`
  );
  console.log(`3. You need to either:`);
  console.log(`   - Use a different conversion method that embeds images`);
  console.log(
    `   - Extract images from the DOCX manually and convert to data URIs`
  );
  console.log(
    `   - Use the enhanced Carbone version with proper DOCX→HTML image handling`
  );
}
