const postprocessor = require("./lib/postprocessor");

// Test HTML template with an image
const template = {
  isZipped: false,
  filename: "test.html",
  embeddings: [],
  extension: "html",
  files: [
    {
      name: "test.html",
      parent: "",
      data: `<html>
        <body>
          <h1>Test Document</h1>
          <img src="https://raw.githubusercontent.com/jumptrnr/carbone/master/doc/carbone_icon_small.png" alt="Carbone Logo">
          <p>This is a test document with an image.</p>
        </body>
      </html>`,
    },
  ],
};

console.log("Testing HTML image processing...");
console.log("Original HTML:");
console.log(template.files[0].data);
console.log("\n" + "=".repeat(80) + "\n");

postprocessor.process(template, {}, {}, function (err) {
  if (err) {
    console.error("Error processing HTML:", err);
    return;
  }

  console.log("Processed HTML:");
  console.log(template.files[0].data);

  // Check if the image was processed
  if (template.files[0].data.includes("data:image/")) {
    console.log("\n✅ SUCCESS: Image was converted to base64 data URI!");
  } else {
    console.log("\n❌ FAILED: Image was not processed");
  }
});
