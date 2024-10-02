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

const alldrawings = /<w:drawing>(.*?)<\/w:drawing>/g;
const allrels = /<Relationship (.*?)\/>/g;

class DocxPostProcessor {
  pattern =
    /(<w:drawing>.*<wp:docPr.*title=)("data:image\/(.*);base64,(.+?)")(.*:embed=")(.*?)(".*<\/w:drawing>)/g;

  constructor(template, data, options, filestore) {
    const documentXmlFile = template.files.find(
      (f) => f.name === "word/document.xml"
    );
    if (!documentXmlFile) return;
    const documentXmlRelsFile = template.files.find(
      (f) => f.name === "word/_rels/document.xml.rels"
    );

    documentXmlFile.data = documentXmlFile.data.replaceAll(
      alldrawings,
      function (drawing) {
        const [, , , mime, content] =
          /(title|descr)="(data:([^;]+);base64,(.*?))"/.exec(drawing) || [];
        const [, relationshipId] = /embed="(.*?)"/.exec(drawing) || [];
        if (!content || !mime || !relationshipId) return drawing;
        const [, extension] = mime.split("/");
        // Save image to media folder
        const [basename, newfile] = filestore.fileBaseName(content);
        const imgFile = `media/${basename}.${extension}`;
        if (newfile) {
          template.files.push({
            name: imgFile,
            isMarked: false,
            data: Buffer.from(content, "base64"),
            parent: "",
          });
          // Update corresponding entry in word/_rels/document.xml.rels file
          documentXmlRelsFile.data = documentXmlRelsFile.data.replaceAll(
            allrels,
            function (relationship) {
              const [, id] = /Id="(.*?)"/.exec(relationship) || [];
              if (id != relationshipId) return relationship;
              return relationship.replace(
                /Target=".*?"/g,
                `Target="/${imgFile}"`
              );
            }
          );
        }
        return drawing.replace(
          /(title|descr)="data:[^;]+;base64,.*?"/g,
          '$1=""'
        );
      }
    );
  }
}

class PptxPostProcessor {
  constructor(template, data, options, filestore) {
    console.log("Starting PptxPostProcessor");
    this.template = template;
    this.data = data;
    this.options = options;
    this.filestore = filestore;
    this.process();
  }

  process() {
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
    const pattern =
      /(<p:pic>.*?<a:blip[^>]*?r:embed=")([^"]+)("[^>]*?>.*?<\/p:pic>)/gs;

    slideFile.data = slideFile.data.replace(
      pattern,
      (match, prefix, rId, suffix) => {
        console.log(`Processing image with rId: ${rId}`);

        const relationship = slideRelsFile.data.match(
          new RegExp(`<Relationship Id="${rId}"[^>]*?Target="([^"]+)"[^>]*?>`)
        );
        if (!relationship) {
          console.log(`No relationship found for rId: ${rId}`);
          return match;
        }

        const imagePath = relationship[1];
        console.log(`Image path: ${imagePath}`);

        if (imagePath.startsWith("data:")) {
          const [, mime, content] =
            /data:([^;]+);base64,(.*)/.exec(imagePath) || [];
          if (!content || !mime) {
            console.log(`Invalid image data for rId: ${rId}`);
            return match;
          }

          const [, extension] = mime.split("/");
          const [basename, newFile] = this.filestore.fileBaseName(content);
          const imgFile = `media/${basename}.${extension}`;
          console.log(`New image file: ${imgFile}, Is new file: ${newFile}`);

          if (newFile) {
            console.log(`Adding new file: ppt/${imgFile}`);
            this.template.files.push({
              name: `ppt/${imgFile}`,
              isMarked: false,
              data: Buffer.from(content, "base64"),
              parent: "",
            });

            console.log(`Updating relationship for rId: ${rId}`);
            slideRelsFile.data = slideRelsFile.data.replace(
              new RegExp(`(<Relationship Id="${rId}"[^>]*?Target=")[^"]+(")`),
              `$1${imgFile}$2`
            );
          }
        } else {
          console.log(
            `Image ${imagePath} is not a base64 encoded data URI, skipping`
          );
        }

        console.log(`Finished processing image with rId: ${rId}`);
        return `${prefix}${rId}${suffix}`;
      }
    );
  }
}

module.exports = postprocessor;
