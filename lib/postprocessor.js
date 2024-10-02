var crypto = require("crypto");
const xml2js = require("xml2js"); // Consider using an XML parser to ensure correctness

const postprocessor = {
  process: function (template, data, options) {
    console.log("Processing template");
    var fileType = template.extension;
    console.log("File type", fileType);
    const filestore = new FileStore();
    switch (fileType) {
      case "odt":
        new OdtPostProcessor(template, data, options, filestore);
        break;
      case "docx":
        new DocxPostProcessor(template, data, options, filestore);
        break;
      case "odp":
        new OdpPostProcessor(template, data, options, filestore);
        break;
      case "xlsx":
      case "ods":
      default:
        break;
    }
  },
};

class FileStore {
  cache = [];
  constructor() {}

  fileBaseName(base64) {
    const hash = crypto.createHash("sha256").update(base64).digest("hex");
    const newfile = !this.cache.includes(hash);
    if (newfile) this.cache.push(hash);
    return [hash, newfile];
  }
}

class OdpPostProcessor {
  constructor(template, data, options, filestore) {
    console.log("Processing ODP file");
    this.template = template;
    this.filestore = filestore;
    this.processContent();
    this.validateFiles();
  }

  processContent() {
    const contentXml = this.template.files.find(
      (f) => f.name === "content.xml"
    );
    if (!contentXml) {
      console.log("content.xml not found");
      return;
    }

    const manifestXml = this.template.files.find(
      (f) => f.name === "META-INF/manifest.xml"
    );
    if (!manifestXml) {
      console.log("manifest.xml not found");
      return;
    }

    // Regular expression to find all <draw:frame> elements
    const drawFrameRegex = /<draw:frame[^>]*>.*?<\/draw:frame>/gs;

    // Process each <draw:frame> element
    contentXml.data = contentXml.data.replace(drawFrameRegex, (drawFrame) => {
      console.log("Processing drawFrame");

      // Look for base64 image data or URL in <svg:title> or <svg:desc>
      const base64Match = drawFrame.match(
        /<svg:(title|desc)>data:([^;]+);base64,([^<]+)<\/svg:\1>/
      );
      const urlMatch = drawFrame.match(
        /<svg:(title|desc)>https?:\/\/[^\s<]+<\/svg:\1>/
      );

      let mime, content;

      if (base64Match) {
        console.log("Base64 image data found");
        [, , mime, content] = base64Match;
      } else if (urlMatch) {
        console.log("URL found, using placeholder image");
        // Use a placeholder image for URLs
        [, mime, content] = PLACEHOLDER_IMAGE.match(/data:([^;]+);base64,(.+)/);
      } else {
        console.log("No valid image reference found in drawFrame");
        return drawFrame;
      }

      const [basename, newfile] = this.filestore.fileBaseName(content);
      const extension = mime.split("/")[1];
      const imgFile = `Pictures/${basename}.${extension}`;

      if (newfile) {
        console.log(`Adding new image file: ${imgFile}`);
        // Add the new image file to the template
        this.template.files.push({
          name: imgFile,
          isMarked: false,
          data: Buffer.from(content, "base64"),
          parent: "",
        });

        // Update the manifest.xml
        manifestXml.data = manifestXml.data.replace(
          /(<manifest:manifest[^>]*>)/,
          `$1\n  <manifest:file-entry manifest:full-path="${imgFile}" manifest:media-type="${mime}"/>`
        );
        console.log(`Updated manifest.xml with new file entry: ${imgFile}`);
      }

      // Update the draw:frame element to reference the new image file
      drawFrame = drawFrame
        // Remove the base64 data or URL from the alt text
        .replace(/<svg:(title|desc)>.*?<\/svg:\1>/, `<svg:$1></$1>`)
        // Update the xlink:href to point to the new image file
        .replace(/xlink:href="[^"]+"/, `xlink:href="${imgFile}"`);

      // Ensure there's a draw:image element
      if (!/<draw:image[^>]*>/.test(drawFrame)) {
        drawFrame = drawFrame.replace(
          /<\/draw:frame>/,
          `<draw:image xlink:href="${imgFile}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`
        );
      } else {
        drawFrame = drawFrame.replace(
          /<draw:image[^>]*xlink:href="[^"]+"[^>]*>/,
          `<draw:image xlink:href="${imgFile}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`
        );
      }

      console.log("Updated drawFrame with new image reference");

      return drawFrame;
    });

    console.log("Processed content.xml and updated manifest.xml");
  }

  validateFiles() {
    console.log("Validating processed files");
    this.template.files.forEach((file) => {
      if (!file.data || file.data.length === 0) {
        console.warn(`Empty or invalid file: ${file.name}`);
      }
    });

    // Check if essential files are present
    const essentialFiles = [
      "content.xml",
      "styles.xml",
      "META-INF/manifest.xml",
    ];
    essentialFiles.forEach((fileName) => {
      if (!this.template.files.some((f) => f.name === fileName)) {
        console.error(`Missing essential file: ${fileName}`);
      }
    });
  }
}

const PLACEHOLDER_IMAGE =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAB7CAAAewgFu0HU+AAAAB3RJTUUH5AQUDSc6eCU9pAAAAAFzUkdCAK7OHOkAAAELSURBVGje7cEBDQAwEAOh+jedjYAA+4zIWlywjW6RiMI+BeAQF0AAEIAQMgAYQgAwgBAyABhCABCCADAABCCADAABDAEFAABCAEDIAEJIAwAAgFAAAEEwBQAAEIATAgAUQgxAADAABCAEDAABCAADAABCAEDAAChQBAAAAIoGbnMAKxXv+A4AAAAASUVORK5CYII=";

module.exports = postprocessor;
