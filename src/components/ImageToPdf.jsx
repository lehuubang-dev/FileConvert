import { useState } from "react";
import { Download, Loader2 } from "lucide-react";
import { jsPDF } from "jspdf";
import DropZone from "./DropZone";

const ACCENT = "#2F7D5D"; // xanh lá — nhận diện cho ảnh

function loadImage(file) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => resolve({ img, url });
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error(`Không đọc được ảnh "${file.name}"`));
    };
    img.src = url;
  });
}

// Vẽ lại ảnh qua canvas rồi xuất JPEG — tránh phải tự dò định dạng gốc
// (jpg/png/webp đều xử lý giống nhau), đổi lại ảnh PNG có nền trong suốt
// sẽ chuyển thành nền trắng vì JPEG không hỗ trợ alpha.
function imageToJpegDataUrl(img) {
  const canvas = document.createElement("canvas");
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0);
  return canvas.toDataURL("image/jpeg", 0.92);
}

export default function ImageToPdf() {
  const [files, setFiles] = useState([]);
  const [isConverting, setIsConverting] = useState(false);
  const [status, setStatus] = useState(null);

  async function handleConvert() {
    setIsConverting(true);
    setStatus(null);
    try {
      const pdf = new jsPDF({ unit: "pt", format: "a4" });
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 24;

      for (let i = 0; i < files.length; i++) {
        const { img, url } = await loadImage(files[i]);
        const maxW = pageWidth - margin * 2;
        const maxH = pageHeight - margin * 2;
        const scale = Math.min(maxW / img.naturalWidth, maxH / img.naturalHeight);
        const w = img.naturalWidth * scale;
        const h = img.naturalHeight * scale;
        const x = (pageWidth - w) / 2;
        const y = (pageHeight - h) / 2;

        const dataUrl = imageToJpegDataUrl(img);
        if (i > 0) pdf.addPage();
        pdf.addImage(dataUrl, "JPEG", x, y, w, h);
        URL.revokeObjectURL(url);
      }

      const baseName = files.length === 1 ? files[0].name.replace(/\.[^.]+$/, "") : "anh-gop";
      pdf.save(`${baseName}.pdf`);
      setStatus({ type: "success", text: `Đã tạo file PDF từ ${files.length} ảnh.` });
    } catch (err) {
      console.error(err);
      setStatus({ type: "error", text: err.message || "Có lỗi khi chuyển đổi. Thử lại." });
    } finally {
      setIsConverting(false);
    }
  }

  return (
    <div>
      <h2 className="font-semibold text-[#1C1C1A]" style={{ fontSize: "1.25rem" }}>
        Ảnh → PDF
      </h2>
      <p className="mt-1 max-w-xl text-sm text-[#6B6963]">
        Gộp một hoặc nhiều ảnh (jpg, png, webp) thành một file PDF, mỗi ảnh
        một trang, theo đúng thứ tự đã chọn. Ảnh có nền trong suốt sẽ đổi
        thành nền trắng khi xuất PDF.
      </p>

      <div className="mt-6">
        <DropZone
          accent={ACCENT}
          accept="image/jpeg,image/png,image/webp"
          multiple={true}
          files={files}
          onFilesSelected={(fl) => setFiles(fl ? Array.from(fl) : [])}
          hint="Có thể chọn nhiều ảnh cùng lúc"
        />
      </div>

      <button
        onClick={handleConvert}
        disabled={files.length === 0 || isConverting}
        className="mt-5 inline-flex items-center gap-2 rounded-sm px-5 py-2.5 text-sm font-medium text-white transition-opacity disabled:cursor-not-allowed disabled:opacity-40"
        style={{ backgroundColor: ACCENT }}
      >
        {isConverting ? <Loader2 size={16} className="animate-spin" /> : <Download size={16} />}
        {isConverting ? "Đang chuyển đổi..." : "Gộp thành PDF"}
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