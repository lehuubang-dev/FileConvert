import { useState } from "react";
import { Download, Loader2 } from "lucide-react";
import { renderAsync } from "docx-preview";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import DropZone from "./DropZone";

const ACCENT = "#2B5797"; // xanh dương — nhận diện quen thuộc của Word

// docx-preview dựng mỗi trang Word thành 1 <section class="docx"> đúng khổ giấy,
// lề, header/footer, bảng, ảnh... Ta chỉ cần bỏ nền xám + đổ bóng của wrapper
// để chụp từng trang cho sạch.
const PREVIEW_CSS = `
  .docx-wrapper{background:transparent !important;padding:0 !important;display:block !important}
  .docx-wrapper > section.docx{box-shadow:none !important;margin:0 0 24px 0 !important;background:#fff !important}
`;

// Chuẩn hoá theo văn bản hành chính (Nghị định 30/2020/NĐ-CP):
// font Times New Roman, chữ màu đen, không tô nền / không gạch chân link.
// KHÔNG đụng tới cỡ chữ, lề, khoảng cách để layout giữ nguyên như file gốc.
const ADMIN_CSS = `
  .admin-doc section.docx :is(p,span,td,th,li,h1,h2,h3,h4,h5,h6,a,b,i,u,strong,em):not([style*="Wingdings"]):not([style*="Symbol"]){
    font-family:"Times New Roman","Liberation Serif",Times,serif !important;
    color:#000 !important;
  }
  .admin-doc section.docx span{background-color:transparent !important}
  .admin-doc section.docx a{text-decoration:none !important}
`;

const PX_TO_PT = 0.75; // 96dpi (CSS px) -> 72dpi (pt)

// ---------------------------------------------------------------------------
// Lớp chữ "vô hình" phủ lên ảnh trang. Ảnh giữ nguyên hình thức 100%, còn lớp
// chữ cho phép bôi đen / sao chép / tìm kiếm — và quan trọng là để công cụ
// PDF → Word đọc lại được file này (nếu không có lớp chữ, PDF chỉ là ảnh).
// Dùng font Liberation Serif (tương thích Times New Roman, đủ dấu tiếng Việt)
// đặt trong src/assets/fonts/.
// ---------------------------------------------------------------------------
const FONT_URLS = {
  normal: new URL("../assets/fonts/LiberationSerif-Regular.ttf", import.meta.url).href,
  bold: new URL("../assets/fonts/LiberationSerif-Bold.ttf", import.meta.url).href,
  italic: new URL("../assets/fonts/LiberationSerif-Italic.ttf", import.meta.url).href,
  bolditalic: new URL("../assets/fonts/LiberationSerif-BoldItalic.ttf", import.meta.url).href,
};
const FONT_FILES = {
  normal: "Regular",
  bold: "Bold",
  italic: "Italic",
  bolditalic: "BoldItalic",
};
const PDF_FONT = "LiberationSerif";
const ASCENT = 0.891; // tỉ lệ ascent của Liberation Serif / Times New Roman

let fontDataCache = null;

function toBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

async function loadFontData() {
  if (fontDataCache) return fontDataCache;
  const entries = await Promise.all(
    Object.entries(FONT_URLS).map(async ([style, url]) => {
      const res = await fetch(url);
      if (!res.ok) throw new Error("FONT_LOAD");
      return [style, toBase64(await res.arrayBuffer())];
    })
  );
  fontDataCache = Object.fromEntries(entries);
  return fontDataCache;
}

function registerFonts(pdf, fontData) {
  Object.entries(fontData).forEach(([style, b64]) => {
    const file = `${PDF_FONT}-${FONT_FILES[style]}.ttf`;
    pdf.addFileToVFS(file, b64);
    pdf.addFont(file, PDF_FONT, style);
  });
}

// Đặt từng từ của trang vào đúng toạ độ nó đang hiển thị, ở chế độ vô hình.
function addTextLayer(pdf, pageEl, pageRect) {
  const walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT);
  const range = document.createRange();
  let node;
  while ((node = walker.nextNode())) {
    const el = node.parentElement;
    if (!el || !node.data.trim()) continue;
    const cs = getComputedStyle(el);
    if (cs.visibility === "hidden" || cs.display === "none") continue;

    const sizePx = parseFloat(cs.fontSize) || 16;
    const bold = cs.fontWeight === "bold" || parseInt(cs.fontWeight, 10) >= 600;
    const italic = cs.fontStyle === "italic" || cs.fontStyle === "oblique";
    const style = bold && italic ? "bolditalic" : bold ? "bold" : italic ? "italic" : "normal";
    const upper = cs.textTransform === "uppercase";

    pdf.setFont(PDF_FONT, style);
    pdf.setFontSize(sizePx * PX_TO_PT);

    const re = /\S+/g;
    let m;
    while ((m = re.exec(node.data))) {
      range.setStart(node, m.index);
      range.setEnd(node, m.index + m[0].length);
      const r = range.getClientRects()[0];
      if (!r || !r.width) continue;
      const text = (upper ? m[0].toUpperCase() : m[0]).normalize("NFC");
      pdf.text(
        text,
        (r.left - pageRect.left) * PX_TO_PT,
        (r.top - pageRect.top + ASCENT * sizePx) * PX_TO_PT, // y = đường chân chữ
        { renderingMode: "invisible" }
      );
    }
  }
}

function waitForImages(root) {
  const imgs = Array.from(root.querySelectorAll("img"));
  return Promise.all(
    imgs.map((img) =>
      img.complete
        ? Promise.resolve()
        : new Promise((res) => {
            img.onload = res;
            img.onerror = res;
          })
    )
  );
}

export default function WordToPdf() {
  const [files, setFiles] = useState([]);
  const [isConverting, setIsConverting] = useState(false);
  const [status, setStatus] = useState(null);
  const [normalize, setNormalize] = useState(true);

  async function handleConvert() {
    setIsConverting(true);
    setStatus(null);
    let container = null;
    try {
      const file = files[0];
      const arrayBuffer = await file.arrayBuffer();

      // Container phải nằm trong vùng nhìn thấy (0,0) để html2canvas chụp được.
      // z-index âm + nền trắng của giao diện phủ lên nên người dùng không thấy.
      container = document.createElement("div");
      if (normalize) container.className = "admin-doc";
      container.style.cssText = [
        "position:fixed",
        "left:0",
        "top:0",
        "z-index:-9999",
        "width:max-content",
        "background:#ffffff",
        "pointer-events:none",
      ].join(";");
      const styleHost = document.createElement("div");
      const bodyHost = document.createElement("div");
      container.appendChild(styleHost);
      container.appendChild(bodyHost);
      document.body.appendChild(container);

      const extraStyle = document.createElement("style");
      extraStyle.textContent = PREVIEW_CSS + ADMIN_CSS;
      container.appendChild(extraStyle);

      await renderAsync(arrayBuffer, bodyHost, styleHost, {
        className: "docx",
        inWrapper: true,
        ignoreWidth: false, // giữ đúng khổ giấy trong file Word
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true, // ngắt trang đúng như Word
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true,
        useBase64URL: true, // nhúng ảnh dạng base64, tránh lỗi CORS khi chụp
      });

      const pages = Array.from(bodyHost.querySelectorAll("section.docx"));
      if (pages.length === 0) {
        setStatus({ type: "error", text: "File Word không có nội dung để chuyển." });
        return;
      }

      // Chờ ảnh và font tải xong để bố cục không bị lệch.
      await waitForImages(container);
      if (document.fonts?.ready) await document.fonts.ready;

      let pdf = null;
      let fontData = null;
      try {
        fontData = await loadFontData();
      } catch (e) {
        console.warn("Không nạp được font cho lớp chữ:", e);
      }

      // Mỗi trang Word = đúng 1 trang PDF, cùng kích thước → bố cục không đổi.
      for (let i = 0; i < pages.length; i++) {
        const page = pages[i];
        const rect = page.getBoundingClientRect();
        const wPt = rect.width * PX_TO_PT;
        const hPt = rect.height * PX_TO_PT;
        const orientation = wPt > hPt ? "l" : "p";

        const canvas = await html2canvas(page, {
          scale: 2,
          useCORS: true,
          backgroundColor: "#ffffff",
          scrollX: 0,
          scrollY: 0,
        });

        if (!pdf) {
          pdf = new jsPDF({ unit: "pt", format: [wPt, hPt], orientation, compress: true });
          if (fontData) registerFonts(pdf, fontData);
        } else {
          pdf.addPage([wPt, hPt], orientation);
        }
        pdf.addImage(canvas.toDataURL("image/jpeg", 0.95), "JPEG", 0, 0, wPt, hPt);
        if (fontData) addTextLayer(pdf, page, rect);

        // Giải phóng bộ nhớ canvas ngay sau mỗi trang.
        canvas.width = 0;
        canvas.height = 0;
      }

      const baseName = file.name.replace(/\.docx$/i, "");
      pdf.save(`${baseName}.pdf`);

      setStatus({
        type: fontData ? "success" : "info",
        text: fontData
          ? `Đã chuyển đổi xong (${pages.length} trang, giữ nguyên khổ giấy và bố cục). Chữ trong PDF vẫn bôi đen, sao chép được.`
          : `Đã chuyển đổi (${pages.length} trang) nhưng PDF chỉ chứa ảnh, không bôi đen được và công cụ PDF → Word sẽ không đọc lại được. Kiểm tra thư mục src/assets/fonts/.`,
      });
    } catch (err) {
      console.error(err);
      setStatus({ type: "error", text: "Có lỗi khi chuyển đổi. Kiểm tra lại file .docx rồi thử lại." });
    } finally {
      if (container) document.body.removeChild(container);
      setIsConverting(false);
    }
  }

  return (
    <div>
      <h2 className="font-semibold text-[#1C1C1A]" style={{ fontSize: "1.25rem" }}>
        Word → PDF
      </h2>
      <p className="mt-1 max-w-xl text-sm text-[#6B6963]">
        Chuyển file Word (.docx) thành PDF, giữ nguyên khổ giấy, lề và bố cục
        từng trang. Nếu máy chưa cài font của file gốc, trình duyệt sẽ dùng font
        thay thế nên chữ có thể lệch nhẹ.
      </p>

      <div className="mt-6">
        <DropZone
          accent={ACCENT}
          accept=".docx"
          multiple={false}
          files={files}
          onFilesSelected={(fl) => setFiles(fl ? Array.from(fl) : [])}
          hint="Chỉ nhận file .docx"
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
            Font Times New Roman, chữ màu đen. Cỡ chữ, lề và bố cục giữ nguyên như file gốc.
          </span>
        </span>
      </label>

      <button
        onClick={handleConvert}
        disabled={files.length === 0 || isConverting}
        className="mt-4 inline-flex items-center gap-2 rounded-xl px-5 py-2.5 text-sm font-medium text-white transition-opacity disabled:cursor-not-allowed disabled:opacity-40"
        style={{ backgroundColor: ACCENT }}
      >
        {isConverting ? <Loader2 size={16} className="animate-spin" /> : <Download size={16} />}
        {isConverting ? "Đang chuyển đổi..." : "Chuyển sang PDF"}
      </button>

      {status && (
        <p
          className="mt-4 rounded-xl px-3 py-2 text-sm"
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