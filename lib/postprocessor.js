const crypto = require("crypto");
const axios = require("axios");
const path = require("path");
const fs = require("fs");
const html2xml = require("./html2xml");

const docxAssets = path.resolve(__dirname, "..", "assets", "docx");

const DEBUG = process.env.DEBUG_DOCX_IMAGES === "true";

// Add a debug logging helper function
const debug = (...args) => {
  if (DEBUG) {
    debug(...args);
  }
};

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
        debug("xlsx postprocessor created but not implemented");
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
    if (!base64 || !Buffer.isBuffer(base64)) {
      debug("Invalid image data provided to fileBaseName");
      return [null, false];
    }

    const hash = crypto
      .createHash("sha256")
      .update(base64)
      .update(Date.now().toString())
      .digest("hex");

    const newfile = !this.cache.includes(hash);
    if (newfile) this.cache.push(hash);
    return [hash, newfile];
  }

  imagesFetched(fetchedImages) {
    this.images = fetchedImages
      .filter((i) => i && i.content && i.mime) // Ensure we have valid image data
      .reduce((m, i) => {
        m[i.url] = i;
        return m;
      }, {});
  }

  getImage(data) {
    if (!data) return {};

    const [, , mime, content, url] =
      /(data:([^;]+);base64,(.*)|(https:\/\/.*))/.exec(data) || [];

    if (mime && content) {
      const buffer = Buffer.from(content, "base64");
      return buffer.length > 0 ? { mime, content: buffer } : {};
    }

    if (!url) return {};

    const image = this.images[url];
    return image && image.content ? image : {};
  }
}

function fetchImages(urls) {
  if (!urls || !urls.length) return Promise.resolve([]);

  debug("Original URLs:", urls);

  const fetches = [...new Set(urls || [])]
    .map((url) => {
      // First decode any HTML entities
      let fixed_url = url
        .replace(/&amp;/g, "&")
        .replace(/&#x[A-F0-9]+;/gi, "")
        .replace(/&#\d+;/g, "")
        .replace(/[\n\r\t]/g, "")
        .trim();

      // For very long URLs (like chart URLs), ensure all components are properly encoded
      try {
        // Parse the URL to separate components
        const urlObj = new URL(fixed_url);

        // Properly encode the query parameters while preserving the structure
        if (urlObj.search) {
          const searchParams = new URLSearchParams(urlObj.search);
          urlObj.search = searchParams.toString();
        }

        // Reconstruct the URL with proper encoding
        fixed_url = urlObj.toString();

        debug(`URL cleaning:`, {
          original: url,
          cleaned: fixed_url,
          length: fixed_url.length,
        });

        return {
          url,
          fixed_url,
        };
      } catch (e) {
        console.error(`Invalid URL format:`, {
          original: url,
          error: e.message,
        });
        return null;
      }
    })
    .filter((item) => item !== null)
    .map(({ url, fixed_url }) =>
      axios({
        method: "get",
        url: fixed_url, // No need for encodeURI since we already properly encoded it
        responseType: "arraybuffer",
        timeout: 60000,
        maxContentLength: Number(process.env.CARBONE_MAX_IMAGE_URL || 10485760),
        validateStatus: (status) => status === 200,
        headers: {
          Accept: "image/*",
          "User-Agent": "Carbone-Document-Generator",
        },
      })
        .then(({ data, headers }) => {
          const contentType = headers["content-type"] || "";

          debug(`Successfully fetched image:`, {
            url: fixed_url,
            contentType,
            dataSize: data.length,
          });

          if (!contentType.startsWith("image/")) {
            throw new Error(`Invalid content type: ${contentType}`);
          }

          return {
            value: {
              url: url, // Keep original URL for reference
              mime: contentType,
              content: data,
            },
          };
        })
        .catch((cause) => {
          console.error(`Error fetching image:`, {
            originalUrl: url,
            cleanedUrl: fixed_url,
            error: cause.message,
            status: cause.response?.status,
            contentType: cause.response?.headers?.["content-type"],
          });
          return null;
        })
    );

  return Promise.all(fetches).then((results) => {
    const validResults = results.filter((result) => result !== null);
    debug(
      `Successfully processed ${validResults.length} out of ${results.length} images`
    );
    return validResults;
  });
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
    debug("OdtPostProcessor initialized");
  }

  processImages(callback) {
    debug("Starting ODT image processing");
    const contentXml = this.template.files.find(
      (f) => f.name === "content.xml"
    );
    if (!contentXml) {
      //debug("content.xml not found, skipping image processing");
      return callback();
    }
    const manifestXml = this.template.files.find(
      (f) => f.name === "META-INF/manifest.xml"
    );

    const frameUrls = (contentXml.data.match(allframes) || [])
      .map((drawFrame) => {
        const [, url] =
          /<svg:title>(https:\/\/(.*?))<\/svg:title>/i.exec(drawFrame) || [];
        return url;
      })
      .filter((url) => !!url);

    debug(`Found ${frameUrls.length} image URLs to process`);

    fetchImages(frameUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));
        //debug(`Fetched ${resolved.length} images`);

        // Use base64 data to create new file and update references
        let processedImages = 0;
        contentXml.data = contentXml.data.replaceAll(allframes, (drawFrame) => {
          const { mime, content } = this.filestore.getImage(
            (/<svg:title>(.*?)<\/svg:title>/.exec(drawFrame) || [])[1]
          );
          if (!content || !mime) {
            //debug("Skipping frame due to missing content or MIME type");
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
            //debug(`Added new image: ${imgFile}`);
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
            //debug(`Updated manifest.xml for ${imgFile}`);
          }
          processedImages++;
          return drawFrame
            .replace(/<svg:title>.*?<\/svg:title>/, "<svg:title></svg:title>")
            .replace(/draw:mime-type="[^"]+"/, `draw:mime-type="${mime}"`)
            .replace(/xlink:href="[^"]+"/, `xlink:href="${imgFile}"`);
        });
        debug(`Processed ${processedImages} images in content.xml`);
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
    debug("ODT HTML processing not implemented yet");
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

    this.relationshipCounter = 1000;
    this.relationshipMap = new Map();
  }

  getNextRelationshipId() {
    return `rId${this.relationshipCounter++}`;
  }

  processImages(callback) {
    if (!this.documentXmlFile) return callback();

    debug("Searching for drawings in document...");
    const allDrawings = this.documentXmlFile.data.match(alldrawings) || [];
    debug(`Found ${allDrawings.length} drawings`);

    const drawingUrls = allDrawings
      .map((drawing) => {
        // Check for URLs in both title and descr attributes
        const urlMatch = drawing.match(/(?:title|descr)="(https?:\/\/[^"]+)"/i);

        if (!urlMatch) {
          debug("No URL found in drawing");
          return null;
        }

        const url = urlMatch[1].replace(/&#x[A-F0-9]+;/gi, "").trim();
        const [, relationshipId] = /embed="([^"]+)"/.exec(drawing) || [];

        debug("Found image reference:", {
          url,
          relationshipId,
        });

        return { url, relationshipId };
      })
      .filter((item) => item !== null);

    debug(`Found ${drawingUrls.length} image URLs in drawings`);

    const uniqueUrls = [...new Set(drawingUrls.map((item) => item.url))];
    debug("Unique URLs to fetch:", uniqueUrls);

    fetchImages(uniqueUrls)
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));

        const processedRelationships = new Set();

        this.documentXmlFile.data = this.documentXmlFile.data.replaceAll(
          alldrawings,
          (drawing) => {
            // Check if this drawing contains a URL we need to process
            const urlMatch = drawing.match(
              /(?:title|descr)="(https?:\/\/[^"]+)"/i
            );
            if (!urlMatch) {
              debug("Keeping original drawing (no URL found)");
              return drawing;
            }

            const imageUrl = urlMatch[1].replace(/&#x[A-F0-9]+;/gi, "").trim();
            const { mime, content } = this.filestore.getImage(imageUrl);
            const [, oldRelationshipId] = /embed="([^"]+)"/.exec(drawing) || [];

            if (!content || !mime || !oldRelationshipId) {
              debug(`Skipping drawing due to missing data:`, {
                hasContent: !!content,
                mime,
                oldRelationshipId,
                imageUrl,
              });
              return drawing;
            }

            // Extract original dimensions and other properties to preserve
            const extentMatch = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(
              drawing
            );
            const originalCx = extentMatch ? extentMatch[1] : null;
            const originalCy = extentMatch ? extentMatch[2] : null;

            const [, extension] = mime.split("/");
            let newRelationshipId = this.relationshipMap.get(imageUrl);
            if (!newRelationshipId) {
              newRelationshipId = this.getNextRelationshipId();
              this.relationshipMap.set(imageUrl, newRelationshipId);
            }

            const imgFile = `media/${newRelationshipId}.${extension}`;

            if (!processedRelationships.has(newRelationshipId)) {
              processedRelationships.add(newRelationshipId);

              debug(`Adding new image:`, {
                file: imgFile,
                oldId: oldRelationshipId,
                newId: newRelationshipId,
                url: imageUrl,
              });

              // Add the image file
              this.template.files.push({
                name: `word/${imgFile}`,
                isMarked: false,
                data: Buffer.from(content),
                parent: "",
              });

              // Add relationship entry
              const relationshipEntry = `<Relationship Id="${newRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${imgFile}"/>`;

              // Insert the new relationship before the closing tag
              this.documentXmlRelsFile.data =
                this.documentXmlRelsFile.data.replace(
                  /<\/Relationships>/,
                  `${relationshipEntry}</Relationships>`
                );

              debug(
                `Added relationship: ${oldRelationshipId} -> ${newRelationshipId}`
              );
            }

            // Create updated drawing with new relationship ID but preserve dimensions
            const updatedDrawing = drawing
              .replace(/(title|descr)="[^"]*"/g, '$1=""')
              .replace(/embed="[^"]+"/g, `embed="${newRelationshipId}"`);

            debug(
              `Updated drawing relationship: ${oldRelationshipId} -> ${newRelationshipId}`
            );

            return updatedDrawing;
          }
        );

        debug(`Processed ${processedRelationships.size} unique relationships`);
        debug(
          "Final relationships file content:",
          this.documentXmlRelsFile.data
        );

        return callback(
          this.filestore.fetchErrorMessage
            ? new Error(this.filestore.fetchErrorMessage)
            : undefined
        );
      })
      .catch((e) => {
        console.error("Image processing error:", e);
        callback(new Error(`Image processing error: ${e.message}`));
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
        const titleMatch = /<svg:title>(https:\/\/.*?)<\/svg:title>/i.exec(
          drawFrame
        );
        const descMatch = /<svg:desc>(https:\/\/.*?)<\/svg:desc>/i.exec(
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
          const titleMatch = /<svg:title>(https:\/\/.*?)<\/svg:title>/i.exec(
            drawFrame
          );
          const descMatch = /<svg:desc>(https:\/\/.*?)<\/svg:desc>/i.exec(
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

          // Extract original dimensions to preserve aspect ratio
          const widthMatch = /svg:width="([^"]+)"/.exec(drawFrame);
          const heightMatch = /svg:height="([^"]+)"/.exec(drawFrame);
          const originalWidth = widthMatch ? widthMatch[1] : null;
          const originalHeight = heightMatch ? heightMatch[1] : null;

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
            // Preserve original dimensions if available
            const imageTag = `<draw:image xlink:href="${imgFile}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`;
            updatedDrawFrame = updatedDrawFrame.replace(
              /<draw:frame/,
              `<draw:frame>${imageTag}`
            );
          }

          return updatedDrawFrame;
        });

        //debug(`Processed ${processedImages} images in content.xml`);
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
      //debug("No presentation.xml file found, skipping image processing");
      return callback();
    }

    const blipTags =
      this.presentationXmlFile.data.match(/<a:blip r:embed="[^"]+"/g) || [];
    debug(`Found ${blipTags.length} <a:blip> tags in presentation.xml`);

    const drawingUrls = blipTags
      .map((blip) => {
        const [, rId] = /r:embed="(rId\d+)"/.exec(blip) || [];
        if (!rId) {
          debug(`Skipping blip tag due to missing rId: ${blip}`);
          return null;
        }
        const relationship = this.presentationXmlRelsFile.data.match(
          new RegExp(`<Relationship Id="${rId}"[^>]+>`)
        );
        if (!relationship) {
          debug(
            `Skipping blip tag due to missing relationship for rId ${rId}: ${blip}`
          );
          return null;
        }
        const [, url] = /Target="([^"]+)"/.exec(relationship[0]) || [];
        if (!url) {
          debug(
            `Skipping blip tag due to missing Target URL in relationship: ${relationship[0]}`
          );
          return null;
        }
        if (!url.startsWith("http")) {
          debug(`Skipping non-HTTP URL: ${url}`);
          return null;
        }
        return { url, rId };
      })
      .filter((item) => !!item);

    //debug(`Found ${drawingUrls.length} valid image URLs to process`);

    fetchImages(drawingUrls.map((item) => item.url))
      .then((resolved) => {
        this.filestore.imagesFetched(resolved.map((r) => r.value));

        // Find all image elements to preserve dimensions
        const imageElements = {};
        const extentRegex =
          /<a:blip r:embed="(rId\d+)"[^>]*>[\s\S]*?<a:xfrm>[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"[^>]*>/g;
        let match;
        while (
          (match = extentRegex.exec(this.presentationXmlFile.data)) !== null
        ) {
          const [, rId, cx, cy] = match;
          imageElements[rId] = { cx, cy };
        }

        this.presentationXmlFile.data = this.presentationXmlFile.data.replace(
          /<a:blip r:embed="(rId\d+)"/g,
          (match, rId) => {
            const relationship = this.presentationXmlRelsFile.data.match(
              new RegExp(`<Relationship Id="${rId}"[^>]+>`)
            );
            if (!relationship) {
              debug(
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
