import { useState } from "react";
import { Download, Loader2 } from "lucide-react";
import * as pdfjsLib from "pdfjs-dist";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Tab,
  TabStopType,
  AlignmentType,
  LineRuleType,
  PageOrientation,
  Header,
  Footer,
  PageNumber,
} from "docx";
import DropZone from "./DropZone";

const ACCENT = "#C4432E"; // đỏ cam — cùng tông nhận diện PDF quen thuộc

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url
).toString();

/* ------------------------------------------------------------------ */
/* Chuẩn văn bản hành chính (Nghị định 30/2020/NĐ-CP)                  */
/* Font Times New Roman, chữ đen, khổ A4, lề: trên/dưới 20–25mm,        */
/* trái 30–35mm, phải 15–20mm.                                          */
/* ------------------------------------------------------------------ */
const ADMIN_FONT = "Times New Roman";
const MM = 72 / 25.4; // 1mm tính bằng pt
const MARGIN_RANGE = {
  top: [20 * MM, 25 * MM],
  bottom: [20 * MM, 25 * MM],
  left: [30 * MM, 35 * MM],
  right: [15 * MM, 20 * MM],
};
const clamp = (v, [lo, hi]) => Math.min(hi, Math.max(lo, v));
const median = (arr) => {
  const s = [...arr].sort((a, b) => a - b);
  return s.length ? s[Math.floor(s.length / 2)] : 0;
};
const twips = (pt) => Math.round(pt * 20);

// Đầu dòng bắt đầu một đoạn / mục mới (Điều 1, 1., a), -, I. ...)
const LIST_START =
  /^(\d+[.)]\s|\d+(\.\d+)+\.?\s|[a-zđ]\)\s|[-–•+]\s|Điều\s+\d|Chương\s+[IVXLC\d]|Mục\s+\d|[IVX]+\.\s)/i;
const PAGE_NUM_RE = /^[-–—\s]*\d{1,3}[-–—\s]*$/;

/* ----------------------------- Đọc PDF ----------------------------- */

function parseFont(realName) {
  const s = (realName || "").toLowerCase();
  return {
    bold: /bold|black|heavy|semibold|demi|-bd\b|\bbd\b/.test(s),
    italic: /italic|oblique|-it\b|ital\b/.test(s),
    family: /courier|mono|consolas/.test(s)
      ? "Courier New"
      : /arial|helvet|calibri|verdana|tahoma|sans|segoe/.test(s)
      ? "Arial"
      : ADMIN_FONT,
  };
}

// Nạp font của trang để biết tên thật (đậm / nghiêng / có chân hay không).
async function loadFontInfo(page, content, cache) {
  try {
    await page.getOperatorList();
  } catch {
    /* không nạp được thì dùng mặc định */
  }
  new Set(content.items.map((i) => i.fontName)).forEach((name) => {
    if (!name || cache.has(name)) return;
    let real = "";
    try {
      real = page.commonObjs.get(name)?.name || "";
    } catch {
      /* font chưa sẵn sàng */
    }
    cache.set(name, parseFont(real));
  });
}

function addText(runs, text, fmt) {
  if (!text) return;
  const size = Math.round(fmt.size * 2) / 2;
  const last = runs[runs.length - 1];
  if (
    last &&
    last.bold === fmt.bold &&
    last.italic === fmt.italic &&
    last.size === size &&
    last.family === fmt.family
  ) {
    last.text += text;
  } else {
    runs.push({ text, bold: fmt.bold, italic: fmt.italic, size, family: fmt.family });
  }
}

// Một dòng chữ -> các "cụm" (segment) cách nhau bởi khoảng trắng lớn
// (ví dụ khối cơ quan ban hành bên trái và Quốc hiệu bên phải).
function toLine(row) {
  const items = row.items.sort((a, b) => a.x - b.x);
  const segments = [];
  let cur = null;
  let prevEnd = 0;

  items.forEach((it) => {
    const gap = cur ? it.x - prevEnd : 0;
    if (!cur || gap > Math.max(30, 3 * row.size)) {
      cur = { x0: it.x, x1: it.x + it.w, runs: [] };
      segments.push(cur);
    } else if (gap > 0.15 * row.size) {
      const lastRun = cur.runs[cur.runs.length - 1];
      if (lastRun && !/\s$/.test(lastRun.text) && !/^\s/.test(it.str)) {
        addText(cur.runs, " ", lastRun);
      }
    }
    addText(cur.runs, it.str, it);
    cur.x1 = it.x + it.w;
    prevEnd = cur.x1;
  });

  segments.forEach((seg) => {
    seg.runs.forEach((r) => (r.text = r.text.replace(/\s+/g, " ")));
    if (seg.runs.length) {
      seg.runs[0].text = seg.runs[0].text.trimStart();
      seg.runs[seg.runs.length - 1].text = seg.runs[seg.runs.length - 1].text.trimEnd();
    }
    seg.runs = seg.runs.filter((r) => r.text);
    seg.text = seg.runs.map((r) => r.text).join("");
  });

  const valid = segments.filter((s) => s.text);
  return {
    y: row.y,
    size: row.size,
    segments: valid,
    minX: valid[0]?.x0 ?? 0,
    maxX: valid[valid.length - 1]?.x1 ?? 0,
    text: valid.map((s) => s.text).join(" "),
  };
}

function buildLines(items) {
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const rows = [];
  sorted.forEach((it) => {
    const row = rows.find((r) => Math.abs(r.y - it.y) <= Math.max(2, 0.4 * Math.min(r.size, it.size)));
    if (row) {
      row.items.push(it);
      row.size = Math.max(row.size, it.size);
    } else {
      rows.push({ y: it.y, size: it.size, items: [it] });
    }
  });
  return rows
    .map(toLine)
    .filter((l) => l.segments.length)
    .sort((a, b) => b.y - a.y);
}

async function extractPages(file) {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const fontCache = new Map();
  const pages = [];
  let totalChars = 0;

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const { width, height } = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    await loadFontInfo(page, content, fontCache);

    const items = content.items
      .filter((it) => "str" in it && it.str.trim())
      .map((it) => ({
        str: it.str.normalize("NFC"), // gộp dấu tiếng Việt về dạng chuẩn
        x: it.transform[4],
        y: it.transform[5],
        w: it.width,
        size: Math.hypot(it.transform[0], it.transform[1]) || it.height || 13,
        ...(fontCache.get(it.fontName) || parseFont("")),
      }));
    items.forEach((i) => (totalChars += i.str.length));

    let lines = buildLines(items);

    // Số trang nằm sát mép trên/dưới: bỏ khỏi nội dung, dựng lại bằng header/footer.
    let numPos = null;
    lines = lines.filter((l) => {
      if (l.segments.length === 1 && PAGE_NUM_RE.test(l.text)) {
        if (l.y > height * 0.93) return !(numPos = "top");
        if (l.y < height * 0.07) return !(numPos = "bottom");
      }
      return true;
    });

    pages.push({ width, height, lines, numPos });
  }

  return { pages, pageCount: pdf.numPages, totalChars };
}

/* ------------------------- Phân tích bố cục ------------------------- */

function detectLayout(pages) {
  const lines = pages.flatMap((p) => p.lines);
  const minCount = Math.max(2, Math.round(lines.length * 0.06));

  // Lề trái thật: vị trí bắt đầu dòng xuất hiện nhiều nhất ở phía trái.
  const starts = lines.map((l) => l.minX).sort((a, b) => a - b);
  let bodyLeft = starts[0];
  for (const v of starts) {
    if (starts.filter((x) => Math.abs(x - v) <= 2).length >= minCount) {
      bodyLeft = v;
      break;
    }
  }

  // Lề phải thật: vị trí cuối dòng xa nhất mà nhiều dòng chạm tới (văn bản căn đều hai bên).
  const ends = lines.map((l) => l.maxX).sort((a, b) => b - a);
  let bodyRight = ends[0];
  for (const v of ends) {
    if (ends.filter((x) => Math.abs(x - v) <= 3).length >= minCount) {
      bodyRight = v;
      break;
    }
  }

  const filled = pages.filter((p) => p.lines.length);
  const tops = filled.map((p) => p.height - p.lines[0].y - 0.9 * p.lines[0].size);
  const bottoms = filled.map((p) => {
    const l = p.lines[p.lines.length - 1];
    return l.y - 0.25 * l.size;
  });

  const { width, height } = pages[0];
  return {
    bodyLeft,
    bodyRight,
    margins: {
      top: clamp(Math.min(...tops), MARGIN_RANGE.top),
      bottom: clamp(median(bottoms), MARGIN_RANGE.bottom),
      left: clamp(bodyLeft, MARGIN_RANGE.left),
      right: clamp(width - bodyRight, MARGIN_RANGE.right),
    },
    width,
    height,
  };
}

function classify(line, L) {
  if (line.segments.length > 1) return "multi";
  const bodyW = L.bodyRight - L.bodyLeft;
  const indent = line.minX - L.bodyLeft;
  const center = (line.minX + line.maxX) / 2;
  const docCenter = (L.bodyLeft + L.bodyRight) / 2;
  if (indent > 12 && L.bodyRight - line.maxX > 12 && Math.abs(center - docCenter) <= 0.03 * bodyW) {
    return "center";
  }
  if (line.maxX >= L.bodyRight - 4 && indent > 0.4 * bodyW) return "right";
  if (indent >= 50) return "block";
  return "plain";
}

// Ghép các dòng liên tiếp thành một đoạn văn khi dòng trước đã chạm lề phải.
function buildParagraphs(page, L) {
  const bodyW = L.bodyRight - L.bodyLeft;
  const paras = [];
  page.lines.forEach((line) => {
    const kind = classify(line, L);
    const prev = paras[paras.length - 1];
    const prevLine = prev?.lines[prev.lines.length - 1];
    const canMerge =
      prev &&
      prev.kind === "plain" &&
      kind === "plain" &&
      prevLine.maxX >= L.bodyRight - 0.06 * bodyW &&
      line.minX - L.bodyLeft < 12 &&
      !LIST_START.test(line.text) &&
      prevLine.y - line.y <= 1.8 * line.size &&
      Math.abs(prevLine.size - line.size) < 1;
    if (canMerge) prev.lines.push(line);
    else paras.push({ kind, lines: [line] });
  });
  return paras;
}

/* ------------------------- Dựng file Word ------------------------- */

function toRun(r, normalize) {
  return new TextRun({
    text: r.text,
    bold: r.bold,
    italics: r.italic,
    size: Math.round(r.size * 2), // half-point
    font: normalize ? ADMIN_FONT : r.family,
    color: normalize ? "000000" : undefined,
  });
}

function toDocxParagraph(para, prevLast, isFirstOnPage, pageIdx, page, L, normalize) {
  const first = para.lines[0];
  const rel = (x) => Math.max(0, x - L.margins.left);
  const bodyW = L.bodyRight - L.bodyLeft;

  const pitches = para.lines.slice(1).map((l, i) => para.lines[i].y - l.y);
  const pitch = pitches.length ? median(pitches) : null;
  const lineH = pitch || first.size * 1.15;

  // Khoảng cách phía trên đoạn: dựng lại từ toạ độ dòng trong PDF.
  let before = isFirstOnPage
    ? page.height - first.y - 0.9 * first.size - L.margins.top
    : prevLast.y - first.y - lineH;
  before = Math.min(150, Math.max(0, before));
  if (before < 1.5) before = 0;

  const spacing = { before: twips(before), after: 0 };
  if (pitch) {
    spacing.line = twips(pitch);
    spacing.lineRule = LineRuleType.AT_LEAST;
  }

  const options = {
    spacing,
    pageBreakBefore: isFirstOnPage && pageIdx > 0,
  };
  let children = [];

  if (para.kind === "multi") {
    // Dòng nhiều cụm (vd. cơ quan ban hành | Quốc hiệu): dùng tab để giữ đúng vị trí.
    const segs = first.segments;
    options.alignment = AlignmentType.LEFT;
    options.indent = { left: twips(rel(segs[0].x0)) };
    options.tabStops = segs.slice(1).map((s) => ({
      type: TabStopType.LEFT,
      position: twips(rel(s.x0)),
    }));
    segs.forEach((seg, i) => {
      if (i > 0) children.push(new TextRun({ children: [new Tab()] }));
      seg.runs.forEach((r) => children.push(toRun(r, normalize)));
    });
  } else {
    const runs = [];
    para.lines.forEach((line, i) => {
      if (i > 0 && runs.length) addText(runs, " ", runs[runs.length - 1]);
      line.segments[0].runs.forEach((r) => addText(runs, r.text, r));
    });
    children = runs.map((r) => toRun(r, normalize));

    if (para.kind === "center") {
      options.alignment = AlignmentType.CENTER;
    } else if (para.kind === "right") {
      options.alignment = AlignmentType.RIGHT;
    } else if (para.kind === "block") {
      options.alignment = AlignmentType.LEFT;
      options.indent = { left: twips(rel(first.minX)) };
    } else {
      const nonLast = para.lines.slice(0, -1);
      const hits = nonLast.filter((l) => l.maxX >= L.bodyRight - 4).length;
      const justified = nonLast.length > 0 && hits / nonLast.length >= 0.6;
      options.alignment = justified ? AlignmentType.BOTH : AlignmentType.LEFT;
      const firstIndent = first.minX - L.bodyLeft;
      if (firstIndent >= 12 && firstIndent < 50) options.indent = { firstLine: twips(firstIndent) };
    }
  }

  return new Paragraph({ ...options, children });
}

function pageNumberParagraph() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ children: [PageNumber.CURRENT], font: ADMIN_FONT, size: 26, color: "000000" }),
    ],
  });
}

function buildDocxBlob(pages, normalize) {
  const L = detectLayout(pages);

  const children = [];
  let prevLast = null;
  pages.forEach((page, pageIdx) => {
    buildParagraphs(page, L).forEach((para, i) => {
      children.push(toDocxParagraph(para, prevLast, i === 0, pageIdx, page, L, normalize));
      prevLast = para.lines[para.lines.length - 1];
    });
  });

  // Khổ giấy: A4 chuẩn nếu PDF gần A4, ngược lại giữ đúng kích thước PDF.
  const isA4 = Math.abs(L.width - 595.28) < 6 && Math.abs(L.height - 841.89) < 6;
  const w = isA4 ? 11906 : twips(L.width);
  const h = isA4 ? 16838 : twips(L.height);
  const landscape = w > h;

  // Số trang: dựng lại bằng header/footer thật; trang 1 không đánh số nếu PDF gốc không có.
  const numPos = pages.find((p) => p.numPos)?.numPos;
  const section = {
    properties: {
      titlePage: Boolean(numPos) && !pages[0].numPos,
      page: {
        size: {
          width: landscape ? h : w,
          height: landscape ? w : h,
          orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
        },
        margin: {
          top: twips(L.margins.top),
          bottom: twips(L.margins.bottom),
          left: twips(L.margins.left),
          right: twips(L.margins.right),
          header: 709,
          footer: 709,
        },
      },
    },
    children,
  };
  if (numPos === "top") {
    section.headers = {
      default: new Header({ children: [pageNumberParagraph()] }),
      first: new Header({ children: [new Paragraph({})] }),
    };
  } else if (numPos === "bottom") {
    section.footers = {
      default: new Footer({ children: [pageNumberParagraph()] }),
      first: new Footer({ children: [new Paragraph({})] }),
    };
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: normalize ? { font: ADMIN_FONT, size: 26, color: "000000" } : {},
        },
      },
    },
    sections: [section],
  });
  return Packer.toBlob(doc);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/* ------------------------------ Giao diện ------------------------------ */

export default function PdfToWord() {
  const [files, setFiles] = useState([]);
  const [isConverting, setIsConverting] = useState(false);
  const [status, setStatus] = useState(null);
  const [normalize, setNormalize] = useState(true);

  async function handleConvert() {
    setIsConverting(true);
    setStatus(null);
    try {
      const file = files[0];
      const { pages, pageCount, totalChars } = await extractPages(file);

      // Dưới ~30 ký tự/trang gần như chắc chắn là PDF scan ảnh, không có
      // lớp văn bản để trích — báo rõ thay vì âm thầm xuất ra file rỗng.
      if (totalChars < 30 * pageCount) {
        setStatus({
          type: "error",
          text:
            "Không tìm thấy lớp văn bản trong PDF này (có thể là bản scan ảnh). " +
            "Công cụ hiện chỉ đọc được PDF có chữ đánh máy sẵn, chưa hỗ trợ nhận dạng OCR từ ảnh.",
        });
        return;
      }

      const blob = await buildDocxBlob(pages, normalize);
      const baseName = file.name.replace(/\.pdf$/i, "");
      downloadBlob(blob, `${baseName}.docx`);
      setStatus({
        type: "success",
        text: `Đã chuyển đổi ${pageCount} trang, giữ khổ giấy, căn lề, in đậm/nghiêng. Bảng kẻ khung và hình ảnh không được giữ lại — nên đối chiếu với bản gốc trước khi dùng.`,
      });
    } catch (err) {
      console.error(err);
      setStatus({ type: "error", text: "Có lỗi khi chuyển đổi. Kiểm tra lại file PDF rồi thử lại." });
    } finally {
      setIsConverting(false);
    }
  }

  return (
    <div>
      <h2 className="font-semibold text-[#1C1C1A]" style={{ fontSize: "1.25rem" }}>
        PDF → Word
      </h2>
      <p className="mt-1 max-w-xl text-sm text-[#6B6963]">
        Chuyển văn bản trong file PDF thành file Word (.docx) có thể chỉnh
        sửa được, giữ khổ giấy, lề, căn lề, in đậm/nghiêng và khối đầu văn bản
        (cơ quan ban hành, Quốc hiệu). Bảng kẻ khung và hình ảnh không được giữ
        lại. Chưa hỗ trợ PDF dạng scan ảnh (cần OCR).
      </p>

      <div className="mt-6">
        <DropZone
          accent={ACCENT}
          accept="application/pdf"
          multiple={false}
          files={files}
          onFilesSelected={(fl) => setFiles(fl ? Array.from(fl) : [])}
          hint="Chỉ nhận file .pdf"
        />
      </div>

      <label className="mt-5 flex max-w-xl cursor-pointer items-start gap-2 text-sm text-[#1C1C1A]">
        <input
          type="checkbox"
          checked={normalize}
          onChange={(e) => setNormalize(e.target.checked)}
          className="mt-0.5"
          style={{ accentColor: ACCENT }}
        />
        <span>
          Chuẩn hoá theo văn bản hành chính
          <span className="block text-[#6B6963]">
            Font Times New Roman, chữ màu đen, khổ A4, lề theo Nghị định 30/2020/NĐ-CP
            (trên/dưới 20–25mm, trái 30–35mm, phải 15–20mm).
          </span>
        </span>
      </label>

      <button
        onClick={handleConvert}
        disabled={files.length === 0 || isConverting}
        className="mt-4 inline-flex items-center gap-2 rounded-sm px-5 py-2.5 text-sm font-medium text-white transition-opacity disabled:cursor-not-allowed disabled:opacity-40"
        style={{ backgroundColor: ACCENT }}
      >
        {isConverting ? <Loader2 size={16} className="animate-spin" /> : <Download size={16} />}
        {isConverting ? "Đang chuyển đổi..." : "Chuyển sang Word"}
      </button>

      {status && (
        <p
          className="mt-4 rounded-sm px-3 py-2 text-sm"
          style={{
            backgroundColor: status.type === "error" ? "#FBE1DF" : "#F7F6F3",
            color: status.type === "error" ? "#B3261E" : "#6B6963",
          }}
        >
          {status.text}
        </p>
      )}
    </div>
  );
}