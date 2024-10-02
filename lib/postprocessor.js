var crypto = require("crypto");

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
      case "pptx":
        new PptxPostProcessor(template, data, options, filestore);
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

const allframes = /<draw:frame (.*?)<\/draw:frame>/g;

class OdtPostProcessor {
  constructor(template, data, options, filestore) {
    const contentXml = template.files.find((f) => f.name === "content.xml");
    if (!contentXml) return;
    const manifestXml = template.files.find(
      (f) => f.name === "META-INF/manifest.xml"
    );

    // Use base64 data to create new file and update references
    contentXml.data = contentXml.data.replaceAll(
      allframes,
      function (drawFrame) {
        const [, , mime, content] =
          /<svg:title>(data:([^;]+);base64,(.*?))<\/svg:title>/.exec(
            drawFrame
          ) || [];
        if (!content || !mime) return drawFrame;
        const [, extension] = mime.split("/");
        // Add new image to Pictures folder
        const [basename, newfile] = filestore.fileBaseName(content);
        const imgFile = `Pictures/${basename}.${extension}`;
        if (newfile) {
          template.files.push({
            name: imgFile,
            isMarked: false,
            data: Buffer.from(content, "base64"),
            parent: "",
          });
          // Update manifest.xml file
          manifestXml.data = manifestXml.data.replace(
            /((.|\n)*)(<\/manifest:manifest>)/,
            function (_match, p1, _p2, p3) {
              return [
                p1,
                `<manifest:file-entry manifest:full-path="${imgFile}" manifest:media-type="${mime}"/>`,
                p3,
              ].join("");
            }
          );
        }
        return drawFrame
          .replace(/<svg:title>.*?<\/svg:title>/, "<svg:title></svg:title>")
          .replace(/draw:mime-type="[^"]+"/, `draw:mime-type="${mime}"`)
          .replace(/xlink:href="[^"]+"/, `xlink:href="${imgFile}"`);
      }
    );
  }
}

const alldrawings = /<p:pic>.*?<\/p:pic>/g;
const allrels = /<Relationship .*?\/>/g;

class PptxPostProcessor {
  constructor(template, data, options, filestore) {
    this.template = template;
    this.filestore = filestore;
    this.processSlides();
  }

  processSlides() {
    console.log("Processing slides");
    const slideFiles = this.template.files.filter(
      (f) => f.name.startsWith("ppt/slides/slide") && f.name.endsWith(".xml")
    );
    console.log("#of slides: ", slideFiles.length);
    slideFiles.forEach((slideFile) => {
      console.log("Processing slide", slideFile.name);
      const slideRelsFile = this.template.files.find(
        (f) =>
          f.name ===
          slideFile.name.replace("slides/", "slides/_rels/") + ".rels"
      );

      if (slideRelsFile) {
        this.processSlide(slideFile, slideRelsFile);
      }
    });
  }

  processSlide(slideFile, slideRelsFile) {
    slideFile.data = slideFile.data.replace(/<a:blip[^>]*>/g, (blipElement) => {
      const dataUriMatch = blipElement.match(
        /embed="data:([^;]+);base64,([^"]+)"/
      );
      const relationshipIdMatch = blipElement.match(/r:embed="(.*?)"/);

      if (!dataUriMatch || !relationshipIdMatch) {
        return blipElement;
      }

      const mime = dataUriMatch[1];
      const content = dataUriMatch[2];
      const relationshipId = relationshipIdMatch[1];

      const [basename, newfile] = this.filestore.fileBaseName(content);
      const extension = mime.split("/")[1];
      const imgFile = `media/${basename}.${extension}`;

      if (newfile) {
        this.template.files.push({
          name: `ppt/${imgFile}`,
          isMarked: false,
          data: Buffer.from(content, "base64"),
          parent: "",
        });

        slideRelsFile.data = slideRelsFile.data.replace(
          new RegExp(`Id="${relationshipId}".*?/>`),
          `Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../${imgFile}"/>`
        );
      }

      return blipElement.replace(
        /embed="[^"]+"/,
        `r:embed="${relationshipId}"`
      );
    });
  }
}

class OdpPostProcessor {
  constructor(template, data, options, filestore) {
    console.log("Processing ODP file");
    this.template = template;
    this.filestore = filestore;
    this.processContent();
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
      // Look for alt text in <svg:title> or <svg:desc>
      const altTextMatch =
        drawFrame.match(/<svg:title>(.*?)<\/svg:title>/s) ||
        drawFrame.match(/<svg:desc>(.*?)<\/svg:desc>/s);

      if (!altTextMatch) {
        // No alt text found, return the original drawFrame
        return drawFrame;
      }

      const altText = altTextMatch[1].trim();

      // Check if altText is a data URI (base64 image)
      const dataUriMatch = altText.match(/data:([^;]+);base64,(.+)/);

      // Alternatively, check if altText is a URL
      const urlMatch = altText.match(/https?:\/\/[^\s"]+/);

      let mime, content;

      if (dataUriMatch) {
        console.log("dataUriMatch", dataUriMatch.substring(0, 20));
        mime = dataUriMatch[1];
        content = dataUriMatch[2];
      } else if (urlMatch) {
        // Fetch the image from the URL (synchronously if possible)
        const imageUrl = urlMatch[0];
        console.log(`Fetching image from URL: ${imageUrl}`);

        // Note: Synchronous HTTP requests are deprecated in Node.js.
        // You might need to adjust your approach to handle this asynchronously,
        // or pre-fetch images before processing.

        // For demonstration purposes, we will skip URL fetching.
        console.log(
          "Skipping URL-based image replacement (needs implementation)"
        );
        return drawFrame;
      } else {
        // Alt text doesn't contain a data URI or URL
        return drawFrame;
      }

      const [basename, newfile] = this.filestore.fileBaseName(content);
      const extension = mime.split("/")[1];
      const imgFile = `Pictures/${basename}.${extension}`;

      if (newfile) {
        // Add the new image file to the template
        this.template.files.push({
          name: imgFile,
          isMarked: false,
          data: Buffer.from(content, "base64"),
          parent: "",
        });

        // Update the manifest.xml
        manifestXml.data = manifestXml.data.replace(
          /(<manifest:manifest)([^>]*>)/,
          `$1$2\n<manifest:file-entry manifest:full-path="${imgFile}" manifest:media-type="${mime}"/>`
        );
      }

      // Update the draw:frame element to reference the new image file
      drawFrame = drawFrame
        // Remove the base64 data from the alt text
        .replace(/<svg:title>.*?<\/svg:title>/s, "<svg:title></svg:title>")
        .replace(/<svg:desc>.*?<\/svg:desc>/s, "<svg:desc></svg:desc>")
        // Update the xlink:href to point to the new image file
        .replace(/xlink:href="[^"]+"/, `xlink:href="${imgFile}"`)
        // Update the draw:mime-type
        .replace(/draw:mime-type="[^"]+"/, `draw:mime-type="${mime}"`);

      return drawFrame;
    });

    console.log("Processed content.xml and updated manifest.xml");
  }
}

module.exports = postprocessor;
