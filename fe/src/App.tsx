import React, { useState } from "react";
import axios from "axios";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import rehypeRaw from 'rehype-raw'
import rehypeSanitize from 'rehype-sanitize'

import "./App.css";

function App() {
  const [file, setFile] = useState(null);
  const [filename, setFilename] = useState("");
  const [markdown, setMarkdown] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const onChange = (e: any) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);
    setFilename(f.name);
    setMarkdown("");
  };

  const upload = async () => {
    if (!file) return setError("Choose a .docx file first");
    setLoading(true);
    setError("");
    try {
      const form = new FormData();
      form.append("file", file);
      const res = await axios.post("http://localhost:4000/api/convert", form, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      setMarkdown(res.data.markdown || "");
    } catch (err: any) {
      console.error(err);
      setError(err ? err.message : "failed to convert file");
    } finally {
      setLoading(false);
    }
  };

  // Frontend sanitization schema for rehype-sanitize.
  // This is defensive: backend already sanitizes, but we sanitize again before rendering raw HTML.
  const sanitizeSchema = {
    tagNames: [
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
    attributes: {
      a: ["href", "name", "target", "rel"],
      img: ["src", "alt", "title", "width", "height"],
      "*": ["class", "id", "align"],
    },
    protocolSchemes: ["http", "https", "mailto", "data"],
  };

  const downloadMarkdown = () => {
    const blob = new Blob([markdown], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const suggested = filename
      ? filename.replace(/\.docx?$/i, "") + ".md"
      : "converted.md";
    a.download = suggested;
    a.click();
    URL.revokeObjectURL(url);
  };
  return (
    <div className="w-full max-w-4xl bg-white rounded-2xl shadow p-6">
      <h1 className="text-2xl font-semibold mb-4">DOCX → Markdown Converter</h1>

      <div className="mb-4">
        <label className="block text-sm font-medium text-gray-700">
          Choose .docx file
        </label>
        <div className="mt-1 flex items-center gap-3">
          <input type="file" accept=".docx" onChange={onChange} />
          <div className="text-sm text-gray-500">
            {filename || "No file chosen"}
          </div>
          <button
            className="ml-auto bg-blue-600 text-white px-4 py-2 rounded disabled:opacity-50"
            onClick={upload}
            disabled={!file || loading}
          >
            {loading ? "Converting..." : "Convert"}
          </button>
        </div>
        {error && <div className="text-red-600 mt-2">{error}</div>}
      </div>

      <div className="grid grid-cols-2 gap-6">
        <div>
          <h2 className="text-lg font-medium mb-2">Markdown Preview</h2>
          <div className="h-96 overflow-auto border rounded p-3 bg-gray-50">
            {markdown ? (
              <ReactMarkdown
                remarkPlugins={[remarkGfm]}
                rehypePlugins={[rehypeRaw, [rehypeSanitize, sanitizeSchema]]}
              >
                {markdown}
              </ReactMarkdown>
            ) : (
              <div className="text-sm text-gray-400">
                Converted markdown will appear here.
              </div>
            )}
          </div>
        </div>

        <div>
          <h2 className="text-lg font-medium mb-2">Raw Markdown</h2>
          <div className="h-96 overflow-auto border rounded p-3 bg-gray-900 text-white whitespace-pre-wrap">
            {markdown || (
              <div className="text-sm text-gray-400">No content</div>
            )}
          </div>

          <div className="mt-4 flex gap-2">
            <button
              className="bg-green-600 text-white px-4 py-2 rounded disabled:opacity-50"
              onClick={downloadMarkdown}
              disabled={!markdown}
            >
              Download .md
            </button>
            <button
              className="bg-gray-200 px-4 py-2 rounded"
              onClick={() => {
                setMarkdown("");
                setFile(null);
                setFilename("");
              }}
            >
              Clear
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
