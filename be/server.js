// server.js
const config = require('./config');
const express = require("express");
const multer = require("multer");
const mammoth = require("mammoth");
const TurndownService = require("turndown");
const gfm = require("turndown-plugin-gfm");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const sanitizeHtml = require("sanitize-html");
const os = require("node:os");
const {execSync} = require("child_process");
// Use jsdom when available to safely manipulate HTML DOM before turndown
let jsdom;
try {
  jsdom = require('jsdom');
} catch (e) {
  jsdom = null;
}
const app = express();
app.use(cors());
app.use(express.json());

// Multer setup - store file in memory (buffer) to avoid requiring disk writes
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB limit
  fileFilter: (req, file, cb) => {
    if (
      file.mimetype ===
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ) {
      cb(null, true);
    } else {
      cb(new Error("Only .docx files are allowed"));
    }
  },
});

// Sanitizer options (backend): allow common content and table tags, but strip scripts and event handlers
const SANITIZE_OPTIONS = {
  allowedTags: [
    "a",
    "b",
    "i",
    "strong",
    "em",
    "p",
    "div",
    "span",
    "br",
    "ul",
    "ol",
    "li",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "code",
    "pre",
    "table",
    "thead",
    "tbody",
    "tr",
    "th",
    "td",
    "img",
  ],
  allowedAttributes: {
    a: ["href", "name", "target", "rel"],
    img: ["src", "alt", "title", "width", "height"],
    "*": ["class", "id", "align"],
  },
  // Allow data URIs for images only if you trust your conversion; otherwise remove 'data'
  allowedSchemes: ["http", "https", "mailto", "data"],
};

function convertOutlineToHeadings(markdown) {
    const lines = markdown.split("\n")
    const result = []

    // Regex to detect outline items like "1. Introduction", "1.2. Audience", "1.2.1. Details"
    const outlinePattern = /^(\d+(?:\.\d+)*)\.\s+(.+)$/

    lines.forEach((line) => {
        const match = line.match(outlinePattern)
        if (match) {
            const outlineNumber = match[1] // "1", "1.2", "1.2.1"
            const title = match[2]

            // Calculate heading level based on number of dots + 1
            const level = (outlineNumber.match(/\./g) || []).length + 1
            const heading = "#".repeat(Math.min(level, 6)) + " " + line

            result.push(heading)
        } else {
            result.push(line)
        }
    })

    return result.join("\n")
}

// Thêm số thứ tự vào các heading
function addNumberingToHeadings(html) {
    const counters = [0, 0, 0, 0, 0]; // Đếm cho h1, h2, h3, h4, h5

    // Xử lý từng heading
    html = html.replace(/<h([1-5])>(.*?)<\/h\1>/g, (match, level, content) => {
        level = parseInt(level);

        // Tăng counter của level hiện tại
        counters[level - 1]++;

        // Reset các counter của level thấp hơn
        for (let i = level; i < counters.length; i++) {
            counters[i] = 0;
        }

        // Tạo số thứ tự (VD: 1.2.3)
        const number = counters.slice(0, level)
            .filter(n => n > 0)
            .join('.');

        // Xóa số cũ nếu có trong content
        content = content.replace(/^[\d.]+\s*/, '').trim();

        return `<h${level}>${number}. ${content}</h${level}>`
    });

    return html;
}

// thêm file ảnh vào folder
function addImageToFolder(images){
    const imagesDir = path.join(process.cwd(), "images");

    // tạo folder images nếu chưa tồn tại
    if (!fs.existsSync(imagesDir)) {
        fs.mkdirSync(imagesDir, { recursive: true });
    }

    images.forEach((img, index) => {
        const buffer = Buffer.from(img.buffer, "base64");

        // đặt tên file
        const fileName = img.name || `image_${Date.now()}_${index}.png`;
        const filePath = path.join(imagesDir, fileName);

        fs.writeFileSync(filePath, buffer);
    });
}

// Detect and format JSON blocks (multiline support)
function detectAndFormatJSON(html) {
    const lines = html.split(/(<\/p>|<\/ul>)/);
    //html.split('</p>');
    let result = [];
    let jsonBuffer = [];
    let inJsonBlock = false;
    let braceCount = 0;
    let bracketCount = 0;

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];
        if (!line.trim()) continue;

        const textContent = line.replace(/<[^>]+>/g, '').trim();

        if (!inJsonBlock && textContent.startsWith('{') || textContent.startsWith('[')) {
            inJsonBlock = true;
            jsonBuffer = [textContent];
            // Đếm số dấu ngoặc trong dòng đầu tiên
            braceCount = (textContent.match(/\{/g) || []).length - (textContent.match(/\}/g) || []).length;
            bracketCount = (textContent.match(/\[/g) || []).length - (textContent.match(/\]/g) || []).length;

            // Nếu JSON chỉ có 1 dòng (ví dụ: {"key": "value"})
            if (braceCount === 0 && bracketCount === 0) {
                try {
                    const cleaned = textContent
                        .replace(/\s+/g, ' ')
                        .replace(/\s*:\s*/g, ': ')
                        .replace(/\s*,\s*/g, ', ')
                        .replace(/[“”]/g, '"');

                    const parsed = JSON.parse(cleaned);
                    const formatted = JSON.stringify(parsed, null, 2);
                    result.push(`<pre><code class="language-json">${formatted}</code></pre>`);

                    jsonBuffer = [];
                    inJsonBlock = false;
                    braceCount = 0;
                    bracketCount = 0;
                } catch (e) {
                    result.push(line + '</p>');
                    jsonBuffer = [];
                    inJsonBlock = false;
                    braceCount = 0;
                    bracketCount = 0;
                }
            }
            continue;
        }

        if (inJsonBlock) {
            jsonBuffer.push(textContent);

            // Cập nhật số lượng dấu ngoặc
            braceCount += (textContent.match(/\{/g) || []).length - (textContent.match(/\}/g) || []).length;
            bracketCount += (textContent.match(/\[/g) || []).length - (textContent.match(/\]/g) || []).length;

            // Kiểm tra xem đã đóng hết dấu ngoặc chưa
            if (braceCount === 0 && bracketCount === 0) {
                const jsonText = jsonBuffer.join('\n');
                try {
                    const cleaned = jsonText
                        .replace(/\s+/g, ' ')
                        .replace(/\s*:\s*/g, ': ')
                        .replace(/\s*,\s*/g, ', ')
                        .replace(/\{\s*/g, '{')
                        .replace(/\s*\}/g, '}')
                        .replace(/\[\s*/g, '[')
                        .replace(/\s*\]/g, ']')
                        .replace(/"\s*:\s*/g, '": ')
                        .replace(/,\s*"/g, ', "')
                        .replace(/[“”]/g, '"');

                    const parsed = JSON.parse(cleaned);
                    const formatted = JSON.stringify(parsed, null, 2);
                    result.push(`<pre><code class="language-json">${formatted}</code></pre>`);

                    jsonBuffer = [];
                    inJsonBlock = false;
                    continue;
                } catch (e) {
                    jsonBuffer.forEach(buf => result.push(`<p>${buf}</p>`));
                    jsonBuffer = [];
                    inJsonBlock = false;
                    continue;
                }
            }
            continue;
        }

        result.push(line + '</p>');
    }

    if (jsonBuffer.length > 0) {
        jsonBuffer.forEach(buf => result.push(`<p>${buf}</p>`));
    }

    return result.join('');
}

// Convert buffer to markdown
async function convertDocxBufferToMarkdown(buffer) {

    // Mảng lưu trữ ảnh
    const images = [];
    let imageCounter = 0;

  // mammoth can read from a Buffer if we pass an ArrayBuffer
  const arrayBuffer = bufferToArrayBuffer(buffer);
  // Convert docx -> HTML
  const result = await mammoth.convertToHtml(
    { buffer: arrayBuffer },
    {
      includeDefaultStyleMap: true,
      styleMap: [

          "p[style-name='Style Heading 1 + Bottom: (Single solid line Gray-50%  1.5 pt Lin...'] => h1:fresh", // đặc thù nếu tự tạo h1 theo kiểu riêng tự thay đổi phần này

          "p[style-name='Heading 1'] => h1:fresh",
          "p[style-name='Heading 2'] => h2:fresh",
          "p[style-name='Heading 3'] => h3:fresh",
          "p[style-name='Heading 4'] => h4:fresh",
          "p[style-name='Heading 5'] => h5:fresh",
          "p[style-name='Heading 6'] => h6:fresh",
          "p[style-name='List Paragraph'] => p:fresh",

          // Map theo tên style của bạn
          "p[style-name='ms1'] => h1:fresh",
          "p[style-name='ms2'] => h2:fresh",
          "p[style-name='ms3'] => h3:fresh",
          "p[style-name='ms4'] => h4:fresh",
          "p[style-name='ms5'] => h5:fresh",

          // Hoặc theo list level
          "p[numbering-level='0'] => h1:fresh",
          "p[numbering-level='1'] => h2:fresh",
          "p[numbering-level='2'] => h3:fresh",
          "p[numbering-level='3'] => h4:fresh",
          "p[numbering-level='4'] => h5:fresh"
      ],
      /*convertImage: mammoth.images.inline(function () {
        // Returning base64 images can be enabled here; for now, mammoth will inline images as data: URIs (or empty).
        return Promise.resolve({ src: "" });
      }),*/
        convertImage: mammoth.images.imgElement(function(image) {
            return image.read("base64").then(function(imageBuffer) {
                imageCounter++;
                const imageName = `${config.imgName}_${imageCounter}.${image.contentType.split('/')[1]}`;

                // Lưu ảnh vào mảng
                images.push({
                    name: imageName,
                    contentType: image.contentType,
                    buffer: imageBuffer
                });

                // Trả về HTML img tag với src tạm
                return {
                    src: `./images/${imageName}`
                };
            });
        })
    }
  );

  addImageToFolder(images);

  // Sanitize HTML from mammoth to remove scripts, on* attributes, etc.
  const sanitizedHtml = sanitizeHtml(result.value || "", SANITIZE_OPTIONS);

  // Defensive fix: remove malformed/empty <table> elements that can cause
  // turndown-plugin-gfm to throw when it accesses rows[0]. Prefer DOM parsing
  // with jsdom; fallback to a conservative regex that removes empty tables.
  let cleanedHtml = sanitizedHtml;
  if (jsdom) {
    const { JSDOM } = jsdom;
    const dom = new JSDOM(sanitizedHtml);
    const document = dom.window.document;
    const tables = Array.from(document.getElementsByTagName('table'));
    tables.forEach((table) => {
      // If table has no rows or first row is undefined, remove it
      const rows = table.rows;
      if (!rows || rows.length === 0 || !rows[0]) {
        table.parentNode.removeChild(table);
      }
    });
    cleanedHtml = document.documentElement.innerHTML;
  } else {
    // Fallback: remove tables with no <tr> inside
    cleanedHtml = sanitizedHtml.replace(/<table[\s\S]*?<tr[\s\S]*?<\/tr>[\s\S]*?<\/table>/gi, function(m){
      // keep tables that have at least one <tr>
      return m;
    });
    // Remove tables that do NOT contain any <tr>
    cleanedHtml = cleanedHtml.replace(/<table[\s\S]*?<\/table>/gi, function(m){
      if (/\<tr[\s\S]*?\<\/tr\>/i.test(m)) return m;
      return '';
    });
  }

  cleanedHtml = addNumberingToHeadings(cleanedHtml);
    cleanedHtml = detectAndFormatJSON(cleanedHtml);

  // Setup Turndown with GFM plugin to support tables, strikethrough, task lists
  const turndownService = new TurndownService({
      codeBlockStyle: "fenced",
      fence: '```'
  });
    // Custom rule for code blocks
    turndownService.addRule('codeBlock', {
        filter: function(node) {
            return node.nodeName === 'PRE' &&
                node.firstChild &&
                node.firstChild.nodeName === 'CODE';
        },
        replacement: function(content, node) {
            const className = node.firstChild.className || '';
            const language = className.replace('language-', '') || '';
            return '\n\n```' + language + '\n' + content + '\n```\n\n';
        }
    });
  turndownService.use([gfm.tables]);

  // Convert cleaned HTML -> Markdown
  let markdown = turndownService.turndown(cleanedHtml);
    // Unescape số thứ tự heading
    markdown = markdown.replace(/^(\d+(?:\.\d+)*)\\\./gm, '$1.');
    markdown = markdown.replace(/^[=-]+$/gm, '')
    markdown = convertOutlineToHeadings(markdown)
  return markdown;
}
function bufferToArrayBuffer(buffer) {
  const ab = new ArrayBuffer(buffer.length);
  const view = new Uint8Array(ab);
  for (let i = 0; i < buffer.length; ++i) {
    view[i] = buffer[i];
  }
  return ab;
}

// POST /api/convert
app.post("/api/convert", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const md = await convertDocxBufferToMarkdown(req.file.buffer);

    // Return markdown as JSON and also a downloadable file name suggestion
    res.json({
      filename: path.parse(req.file.originalname).name + ".md",
      markdown: md,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Conversion failed" });
  }
});
// Simple health
app.get("/", (req, res) => res.send("DOCX → MD converter running"));

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
