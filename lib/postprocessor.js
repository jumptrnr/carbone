const crypto = require("crypto");
const axios = require("axios");
const path = require("path");
const fs = require("fs");
const html2xml = require("./html2xml");

const docxAssets = path.resolve(__dirname, "..", "assets", "docx");

const postprocessor = {
  process: function (template, _data, _options, callback) {
    var processor;
    switch (template.extension) {
      case "odt":
        processor = new OdtPostProcessor(template);

        break;
      case "docx":
        processor = new DocxPostProcessor(template);

        break;
      case "html":
        processor = new HtmlPostProcessor(template);

        break;
      case "xlsx":
        console.log("xlsx postprocessor created but not implemented");
        break;
      case "ods":
        return callback();

      case "odp":
        processor = new OdpPostProcessor(template);

        break;
      case "pptx":
        processor = new PptxPostProcessor(template);
        break;
    }
    processor.processImages((err) => {
      if (err) return callback(err);
      processor.processHTML((err) => callback(err));
    });
  },
};

class FileStore {
  cache = [];
  images = {};
  fetchErrorMessage = undefined;
  constructor() {}

  fileBaseName(base64) {
    const hash = crypto.createHash("sha256").update(base64).digest("hex");
    const newfile = !this.cache.includes(hash);
    if (newfile) this.cache.push(hash);
    return [hash, newfile];
  }

  imagesFetched(fetchedImages) {
    this.images = fetchedImages
      .filter((i) => !!i)
      .reduce((m, i) => {
        m[i.url] = i;
        return m;
      }, {});
    this.fetchErrorMessage = fetchedImages.find((i) => i && i.error)?.error;
  }

  getImage(data) {
    const [, , mime, content, url] =
      /(data:([^;]+);base64,(.*)|(https:\/\/.*))/.exec(data) || [];
    if (mime && content) return { mime, content };
    if (!url) return {};
    return this.images[url] || {};
  }
}

function fetchImages(urls) {
  if (!urls || !urls.length) return Promise.resolve([]);
  //console.log(`Found ${urls.length} images:`);

  const fetches = [...new Set(urls || [])]
    .map((url) => ({
      url,
      fixed_url: url
        .replace(/&amp;/g, "&")
        .replace(/&#x[A-F0-9]+;/gi, "") // Remove hex HTML entities
        .replace(/&#\d+;/g, "") // Remove decimal HTML entities
        .trim(),
    }))
    .map(({ url, fixed_url }) =>
      axios({
        method: "get",
        url: fixed_url,
        responseType: "arraybuffer",
        timeout: 60000,
        maxContentLength:
          Number(process.env.CARBONE_MAX_IMAGE_URL || 10485760) || 10485760, // default 10MiB
      })
        .then(({ data, headers }) => {
          const contentType = headers["content-type"] || "";
          if (contentType.startsWith("image/")) {
            //console.log(`Successfully fetched image: ${url} (${contentType})`);
            return {
              url,
              mime: contentType,
              // FIXME need to scan for viruses
              content: Buffer.from(data, "base64"),
            };
          } else {
            console.log(`Error: Unrecognizable image content type for ${url}`);
            return {
              url,
              error: "Unrecognisable image content type",
            };
          }
        })
        .catch((cause) => {
          console.log(`Error fetching image ${url}: ${cause.message}`);
          return {
            url,
            error: cause.message || "Image fetching error",
            cause,
          };
        })
    );
  return Promise.allSettled(fetches);
}

const allframes = /<draw:frame (.*?)<\/draw:frame>/g;

class Processor {
  filestore = new FileStore();
  constructor(template) {
    this.template = template;
  }
}

class OdtPostProcessor extends Processor {
  constructor(template) {
    super(template);
    console.log("OdtPostProcessor initialized");
  }

  processImages(callback) {
    console.log("Starting ODT image processing");
    const contentXml = this.template.files.find(
      (f) => f.name === "content.xml"
    );
    if (!contentXml) {
      //console.log("content.xml not found, skipping image processing");
      return callback();
    }
    const manifestXml = this.template.files.find(
      (f) => f.name === "META-INF/manifest.xml"
    );

    const frameUrls = (contentXml.data.match(allframes) || [])
      .map((drawFrame) => {
        const [, url] =
          /<svg:title>(https:\/\/(.*?))<\/svg:title>/.exec(drawFrame) || [];
        return url;
      })
      .filter((url) => !!url);

    console.log(`Found ${frameUrls.length} image URLs to process`);

    fetchImages(frameUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));
        //console.log(`Fetched ${resolved.length} images`);

        // Use base64 data to create new file and update references
        let processedImages = 0;
        contentXml.data = contentXml.data.replaceAll(allframes, (drawFrame) => {
          const { mime, content } = this.filestore.getImage(
            (/<svg:title>(.*?)<\/svg:title>/.exec(drawFrame) || [])[1]
          );
          if (!content || !mime) {
            //console.log("Skipping frame due to missing content or MIME type");
            return drawFrame;
          }
          const [, extension] = mime.split("/");
          // Add new image to Pictures folder
          const [basename, newfile] = this.filestore.fileBaseName(content);
          const imgFile = `Pictures/${basename}.${extension}`;
          if (newfile) {
            this.template.files.push({
              name: imgFile,
              isMarked: false,
              data: Buffer.from(content, "base64"),
              parent: "",
            });
            //console.log(`Added new image: ${imgFile}`);
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
            //console.log(`Updated manifest.xml for ${imgFile}`);
          }
          processedImages++;
          return drawFrame
            .replace(/<svg:title>.*?<\/svg:title>/, "<svg:title></svg:title>")
            .replace(/draw:mime-type="[^"]+"/, `draw:mime-type="${mime}"`)
            .replace(/xlink:href="[^"]+"/, `xlink:href="${imgFile}"`);
        });
        console.log(`Processed ${processedImages} images in content.xml`);
        return callback(
          this.filestore.fetchErrorMessage
            ? new Error(this.filestore.fetchErrorMessage)
            : undefined
        );
      })
      .catch((e) => {
        console.error("Error during image processing:", e);
        callback(new Error("Image processing error"));
      });
  }

  processHTML(callback) {
    console.log("ODT HTML processing not implemented yet");
    // TODO implement support for ODT
    callback();
  }
}

const alldrawings = /<w:drawing>(.*?)<\/w:drawing>/g;
const allparagraphs = /<w:p[^>]*\/>|<w:p.*?>.*?<\/w:p>/g;
const numberingXmlFilePath = "word/numbering.xml";
const allParagraphStyles = /<w:style w:type="paragraph" w:styleId="[^"]+"/g;
const headingCount = 6;

class DocxPostProcessor extends Processor {
  constructor(template) {
    super(template);
    this.documentXmlFile = this.template.files.find(
      (f) => f.name === "word/document.xml"
    );
    this.numberingXmlFile = this.template.files.find(
      (f) => f.name === numberingXmlFilePath
    );
    this.documentXmlRelsFile = this.template.files.find(
      (f) => f.name === "word/_rels/document.xml.rels"
    );
    this.stylesXmlFile = this.template.files.find(
      (f) => f.name === "word/styles.xml"
    );
    this.contentTypesXmlFile = this.template.files.find(
      (f) => f.name === "[Content_Types].xml"
    );

    this.paragraphStyleIds = new Array(headingCount)
      .fill("")
      .map((_, i) => `Heading${i + 1}`)
      .concat(
        (this.stylesXmlFile.data.match(allParagraphStyles) || []).map(
          (pstyle) => /w:styleId="([^"]+)"/.exec(pstyle)[1]
        )
      );
  }
  processImages(callback) {
    if (!this.documentXmlFile) return callback();

    const drawingUrls = (this.documentXmlFile.data.match(alldrawings) || [])
      .map((drawing) => {
        const [, , url] =
          /(title|descr)="(https:\/\/(.*?))"/.exec(drawing) || [];
        return url;
      })
      .filter((url) => !!url);

    fetchImages(drawingUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));
        this.documentXmlFile.data = this.documentXmlFile.data.replaceAll(
          alldrawings,
          (drawing) => {
            const { mime, content } = this.filestore.getImage(
              (/(title|descr)="(.*?)"/.exec(drawing) || [])[2]
            );
            const [, relationshipId] = /embed="(.*?)"/.exec(drawing) || [];
            if (!content || !mime || !relationshipId) return drawing;

            const [, extension] = mime.split("/");
            const [basename, newfile] = this.filestore.fileBaseName(content);
            const imgFile = `word/media/${basename}.${extension}`;

            if (newfile) {
              this.template.files.push({
                name: imgFile,
                isMarked: false,
                data: Buffer.from(content, "base64"),
                parent: "",
              });

              this.documentXmlRelsFile.data =
                this.documentXmlRelsFile.data.replace(
                  "</Relationships>",
                  `<Relationship Id="${basename}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/${imgFile}"/></Relationships>`
                );
            }

            return drawing
              .replace(
                /(title|descr)="(data:[^;]+;base64,.*?|https:\/\/.*?)"/g,
                '$1=""'
              )
              .replace(/embed="(.*?)"/g, `embed="${basename}"`);
          }
        );
        return callback(
          this.filestore.fetchErrorMessage
            ? new Error(this.filestore.fetchErrorMessage)
            : undefined
        );
      })
      .catch((e) => {
        callback(new Error("Image processing error"));
      });
  }
  processHTML(callback) {
    try {
      this._decodeBase64Html();
      this._processNumberingXml();
      this._processHeadingStyles();
    } catch (e) {
      return callback(e);
    }
    return callback();
  }
  _decodeBase64Html() {
    if (!this.documentXmlFile) return;

    this.documentXmlFile.data = this.documentXmlFile.data.replaceAll(
      allparagraphs,
      (paragraph) => {
        const [, base64] =
          paragraph.match(
            /<w:p.*?>.*<w:t>html:([A-Za-z0-9+\/]+=*):html<\/w:t>.*<\/w:p>/
          ) || [];
        if (!base64) return paragraph;

        const [, pstyle] =
          paragraph.match(
            /<w:p.*?>.*?<w:pPr>.*?<w:pStyle w:val="([A-Za-z0-9_]+)"/
          ) || [];
        const [, pstyles] =
          paragraph.match(/<w:p.*?>.*?<w:pPr>.*?<w:rPr>(.*?)<\/w:rPr>/) || [];
        const [, tstyles] =
          paragraph.match(
            /.*<w:rPr>(.*?)<\/w:rPr>.*?<w:t>html:[A-Za-z0-9+\/]+=*/
          ) || [];

        const processedHtml = Buffer.from(base64, "base64").toString("utf8");
        return new html2xml(processedHtml, {
          pstyle,
          pstyles,
          tstyles,
          validStyles: this.paragraphStyleIds,
        }).getXML();
      }
    );
  }
  _processNumberingXml() {
    const numberingXmlTemplate = fs.readFileSync(
      path.resolve(docxAssets, "numbering.xml"),
      "utf8"
    );
    const [abstractNum1001] =
      numberingXmlTemplate.match(
        /<w:abstractNum w:abstractNumId="1001".*?<\/w:abstractNum>/s
      ) || [];
    if (this.numberingXmlFile) {
      const [abstractNums] =
        numberingXmlTemplate.match(/<w:abstractNum.*<\/w:abstractNum>/s) || [];
      this.numberingXmlFile.data = this.numberingXmlFile.data.replace(
        "<w:abstractNum",
        abstractNums + "<w:abstractNum"
      );
    } else {
      this.numberingXmlFile = {
        name: numberingXmlFilePath,
        isMarked: false,
        data: numberingXmlTemplate,
        parent: "",
      };
      this.template.files.push(this.numberingXmlFile);
      this.contentTypesXmlFile.data = this.contentTypesXmlFile.data.replace(
        "<Override ",
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override '
      );
      this.documentXmlRelsFile.data = this.documentXmlRelsFile.data.replace(
        "</Relationships>",
        '<Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
      );
    }
    let additionalAbstractNums = "";
    let numStyles = "";
    const lastNumId = new html2xml().getLastNumId();
    for (let i = 1002; i <= lastNumId; i++)
      additionalAbstractNums += abstractNum1001.replace("1001", i);
    if (additionalAbstractNums)
      this.numberingXmlFile.data = this.numberingXmlFile.data.replace(
        "<w:abstractNum",
        additionalAbstractNums + "<w:abstractNum"
      );
    for (let i = 1000; i <= lastNumId; i++)
      numStyles += `<w:num w:numId="${i}"><w:abstractNumId w:val="${i}"/></w:num>`;
    this.numberingXmlFile.data = this.numberingXmlFile.data.replace(
      "</w:numbering>",
      numStyles + "</w:numbering>"
    );
  }
  _processHeadingStyles() {
    // FIXME Check if styles to be added.
    if (!this.stylesXmlFile) return;

    const headings = (level) => `
      <w:style w:type="paragraph" w:styleId="Heading${level}">
        <w:name w:val="Heading ${level}"/>
        <w:basedOn w:val="Heading"/>
        <w:next w:val="TextBody"/>
        <w:qFormat/>
        <w:pPr>
          <w:numPr>
            <w:ilvl w:val="0"/>
            <w:numId w:val="1"/>
          </w:numPr>
          <w:spacing w:before="240" w:after="120"/>
          <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
          <w:b/>
          <w:bCs/>
          <w:sz w:val="${36 - 4 * level}"/>
          <w:szCs w:val="${36 - 4 * level}"/>
        </w:rPr>
      </w:style>
    `;
    const xml = new Array(headingCount)
      .fill("")
      .map((_, i) => headings(i + 1).trim())
      .join("");
    this.stylesXmlFile.data = this.stylesXmlFile.data.replace(
      /(.*)(<\/w:styles>)/,
      `$1${xml}$2`
    );
  }
}

class HtmlPostProcessor extends Processor {
  constructor(template) {
    super(template);
  }
  processImages(callback) {
    return callback();
  }
  processHTML(callback) {
    try {
      this._decodeBase64Html();
    } catch (e) {
      return callback(e);
    }
    return callback();
  }
  _decodeBase64Html() {
    const htmlFile = this.template.files[0];
    htmlFile.data = htmlFile.data.replaceAll(
      /html:([A-Za-z0-9+/]+=*):html/g,
      (_, base64html) => Buffer.from(base64html, "base64").toString("utf8")
    );
  }
}

class OdpPostProcessor extends Processor {
  constructor(template) {
    super(template);
  }

  processImages(callback) {
    const contentXml = this.template.files.find(
      (f) => f.name === "content.xml"
    );
    if (!contentXml) {
      return callback();
    }
    const manifestXml = this.template.files.find(
      (f) => f.name === "META-INF/manifest.xml"
    );

    const frameUrls = (contentXml.data.match(allframes) || [])
      .map((drawFrame) => {
        const titleMatch = /<svg:title>(https:\/\/.*?)<\/svg:title>/.exec(
          drawFrame
        );
        const descMatch = /<svg:desc>(https:\/\/.*?)<\/svg:desc>/.exec(
          drawFrame
        );
        return titleMatch ? titleMatch[1] : descMatch ? descMatch[1] : null;
      })
      .filter((url) => !!url);

    fetchImages(frameUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));

        let processedImages = 0;
        contentXml.data = contentXml.data.replaceAll(allframes, (drawFrame) => {
          const titleMatch = /<svg:title>(https:\/\/.*?)<\/svg:title>/.exec(
            drawFrame
          );
          const descMatch = /<svg:desc>(https:\/\/.*?)<\/svg:desc>/.exec(
            drawFrame
          );
          const url = titleMatch
            ? titleMatch[1]
            : descMatch
            ? descMatch[1]
            : null;

          if (!url) {
            return drawFrame;
          }

          const { mime, content } = this.filestore.getImage(url);
          if (!content || !mime) {
            return drawFrame;
          }

          const [, extension] = mime.split("/");
          const [basename, newfile] = this.filestore.fileBaseName(content);
          const imgFile = `Pictures/${basename}.${extension}`;

          if (newfile) {
            this.template.files.push({
              name: imgFile,
              isMarked: false,
              data: Buffer.from(content, "base64"),
              parent: "",
            });
            manifestXml.data = manifestXml.data.replace(
              /((.|\n)*)(<\/manifest:manifest>)/,
              (_match, p1, _p2, p3) =>
                `${p1}<manifest:file-entry manifest:full-path="${imgFile}" manifest:media-type="${mime}"/>${p3}`
            );
          }

          processedImages++;
          let updatedDrawFrame = drawFrame
            .replace(/<svg:title>.*?<\/svg:title>/, "<svg:title></svg:title>")
            .replace(/<svg:desc>.*?<\/svg:desc>/, "<svg:desc></svg:desc>")
            .replace(/draw:mime-type="[^"]+"/, `draw:mime-type="${mime}"`)
            .replace(/xlink:href="[^"]+"/, `xlink:href="${imgFile}"`);

          // Ensure the image is actually displayed
          if (!updatedDrawFrame.includes("<draw:image")) {
            updatedDrawFrame = updatedDrawFrame.replace(
              /<draw:frame/,
              `<draw:frame><draw:image xlink:href="${imgFile}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`
            );
          }

          return updatedDrawFrame;
        });

        //console.log(`Processed ${processedImages} images in content.xml`);
        return callback(
          this.filestore.fetchErrorMessage
            ? new Error(this.filestore.fetchErrorMessage)
            : undefined
        );
      })
      .catch((e) => {
        console.error("Error during image processing:", e);
        callback(new Error("Image processing error"));
      });
  }

  processHTML(callback) {
    // TODO implement support for ODP if needed
    callback();
  }
}

class PptxPostProcessor extends Processor {
  constructor(template) {
    super(template);
    this.presentationXmlFile = this.template.files.find(
      (f) => f.name === "ppt/presentation.xml"
    );
    this.presentationXmlRelsFile = this.template.files.find(
      (f) => f.name === "ppt/_rels/presentation.xml.rels"
    );
    this.contentTypesXmlFile = this.template.files.find(
      (f) => f.name === "[Content_Types].xml"
    );
  }

  processImages(callback) {
    if (!this.presentationXmlFile) {
      //console.log("No presentation.xml file found, skipping image processing");
      return callback();
    }

    const blipTags =
      this.presentationXmlFile.data.match(/<a:blip r:embed="[^"]+"/g) || [];
    console.log(`Found ${blipTags.length} <a:blip> tags in presentation.xml`);

    const drawingUrls = blipTags
      .map((blip) => {
        const [, rId] = /r:embed="(rId\d+)"/.exec(blip) || [];
        if (!rId) {
          console.log(`Skipping blip tag due to missing rId: ${blip}`);
          return null;
        }
        const relationship = this.presentationXmlRelsFile.data.match(
          new RegExp(`<Relationship Id="${rId}"[^>]+>`)
        );
        if (!relationship) {
          console.log(
            `Skipping blip tag due to missing relationship for rId ${rId}: ${blip}`
          );
          return null;
        }
        const [, url] = /Target="([^"]+)"/.exec(relationship[0]) || [];
        if (!url) {
          console.log(
            `Skipping blip tag due to missing Target URL in relationship: ${relationship[0]}`
          );
          return null;
        }
        if (!url.startsWith("http")) {
          console.log(`Skipping non-HTTP URL: ${url}`);
          return null;
        }
        return url;
      })
      .filter((url) => !!url);

    //console.log(`Found ${drawingUrls.length} valid image URLs to process`);

    fetchImages(drawingUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));
        this.presentationXmlFile.data = this.presentationXmlFile.data.replace(
          /<a:blip r:embed="(rId\d+)"/g,
          (match, rId) => {
            const relationship = this.presentationXmlRelsFile.data.match(
              new RegExp(`<Relationship Id="${rId}"[^>]+>`)
            );
            if (!relationship) {
              console.log(
                `Skipping image replacement due to missing relationship for rId ${rId}`
              );
              return match;
            }
            const [, url] = /Target="([^"]+)"/.exec(relationship[0]) || [];
            if (!url || !url.startsWith("http")) return match;

            const { mime, content } = this.filestore.getImage(url);
            if (!content || !mime) return match;

            const [, extension] = mime.split("/");
            const [basename, newfile] = this.filestore.fileBaseName(content);
            const imgFile = `ppt/media/${basename}.${extension}`;

            if (newfile) {
              this.template.files.push({
                name: imgFile,
                isMarked: false,
                data: Buffer.from(content, "base64"),
                parent: "",
              });
              this.presentationXmlRelsFile.data =
                this.presentationXmlRelsFile.data.replace(
                  `<Relationship Id="${rId}" `,
                  `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${basename}.${extension}" `
                );
              this.contentTypesXmlFile.data =
                this.contentTypesXmlFile.data.replace(
                  "</Types>",
                  `<Default Extension="${extension}" ContentType="${mime}"/></Types>`
                );
            }

            return `<a:blip r:embed="${rId}"`;
          }
        );

        return callback(
          this.filestore.fetchErrorMessage
            ? new Error(this.filestore.fetchErrorMessage)
            : undefined
        );
      })
      .catch((e) => {
        console.error("Error during image processing:", e);
        callback(new Error("Image processing error"));
      });
  }

  processHTML(callback) {
    // TODO: Implement HTML processing for PPTX if needed
    callback();
  }
}

module.exports = postprocessor;
