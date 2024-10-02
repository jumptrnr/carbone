var crypto = require("crypto");

const postprocessor = {
  process: function (template, data, options) {
    var fileType = template.extension;
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
  pattern = /(<p:pic>.*?<a:blip[^>]*r:embed=")([^"]+)(".*?<\/p:pic>)/g;

  constructor(template, data, options, filestore) {
    console.log("Starting PptxPostProcessor");
    this.template = template;
    this.filestore = filestore;
    this.processSlides();
  }

  processSlides() {
    const slideFiles = this.template.files.filter(
      (f) => f.name.startsWith("ppt/slides/slide") && f.name.endsWith(".xml")
    );
    console.log(`Found ${slideFiles.length} slide(s) to process`);

    slideFiles.forEach((slideFile) => {
      console.log(`Processing slide: ${slideFile.name}`);
      const slideRelsFile = this.template.files.find(
        (f) =>
          f.name ===
          slideFile.name.replace("slides/", "slides/_rels/") + ".rels"
      );

      if (slideRelsFile) {
        this.processSlide(slideFile, slideRelsFile);
      } else {
        console.log(`No relationship file found for ${slideFile.name}`);
      }
    });
  }

  processSlide(slideFile, slideRelsFile) {
    slideFile.data = slideFile.data.replaceAll(alldrawings, (drawing) => {
      console.log("Processing a drawing in the slide");
      const [, , mime, content] =
        /(descr|title)="(data:([^;]+);base64,(.*?))"/.exec(drawing) || [];
      const [, relationshipId] = /r:embed="(.*?)"/.exec(drawing) || [];

      if (!content || !mime || !relationshipId) {
        console.log("Drawing doesn't contain necessary information, skipping");
        return drawing;
      }

      console.log(`Processing image with relationshipId: ${relationshipId}`);
      const [, extension] = mime.split("/");

      // Save image to media folder
      const [basename, newfile] = this.filestore.fileBaseName(content);
      const imgFile = `media/${basename}.${extension}`;

      if (newfile) {
        console.log(`Adding new file: ppt/${imgFile}`);
        this.template.files.push({
          name: `ppt/${imgFile}`,
          isMarked: false,
          data: Buffer.from(content, "base64"),
          parent: "",
        });

        // Update corresponding entry in slide rels file
        slideRelsFile.data = slideRelsFile.data.replaceAll(
          allrels,
          (relationship) => {
            const [, id] = /Id="(.*?)"/.exec(relationship) || [];
            if (id != relationshipId) return relationship;
            console.log(`Updating relationship for id: ${id}`);
            return relationship.replace(
              /Target=".*?"/,
              `Target="../${imgFile}"`
            );
          }
        );
      }

      console.log("Finished processing drawing");
      return drawing.replace(/(descr|title)="data:[^;]+;base64,.*?"/, '$1=""');
    });
  }
}

module.exports = postprocessor;
