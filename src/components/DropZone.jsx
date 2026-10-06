import { useRef, useState } from "react";
import { Upload, FileCheck2 } from "lucide-react";

/**
 * Vùng chọn file dùng chung cho cả 3 bộ chuyển đổi.
 * - accent: màu chủ đạo theo loại file
 * - accept: chuỗi mime/extension cho input file
 * - multiple: cho phép chọn nhiều file cùng lúc hay không
 * - files: mảng File hiện tại (do component cha quản lý state)
 * - onFilesSelected: callback trả về FileList khi người dùng chọn/thả file
 */
export default function DropZone({
  accent,
  accept,
  multiple = false,
  files = [],
  onFilesSelected,
  hint,
}) {
  const inputRef = useRef(null);
  const [isDragOver, setIsDragOver] = useState(false);

  function handleDrop(e) {
    e.preventDefault();
    setIsDragOver(false);
    if (e.dataTransfer.files?.length) onFilesSelected(e.dataTransfer.files);
  }

  return (
    <div>
      <div
        role="button"
        tabIndex={0}
        onClick={() => inputRef.current?.click()}
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            inputRef.current?.click();
          }
        }}
        onDragOver={(e) => {
          e.preventDefault();
          setIsDragOver(true);
        }}
        onDragLeave={() => setIsDragOver(false)}
        onDrop={handleDrop}
        className="cursor-pointer rounded-2xl border-2 border-dashed px-6 py-10 text-center transition-colors focus-visible:outline-2 focus-visible:outline-offset-2"
        style={{
          borderColor: isDragOver || files.length > 0 ? accent : "#E3E1DA",
          backgroundColor: isDragOver ? `${accent}0D` : "#FCFCFB",
          outlineColor: accent,
        }}
      >
        <div
          className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-2xl"
          style={{ backgroundColor: `${accent}14`, color: accent }}
        >
          {files.length > 0 ? <FileCheck2 size={22} /> : <Upload size={22} />}
        </div>
        {files.length > 0 ? (
          <div>
            <p className="text-sm font-medium text-[#1C1C1A]">Đã chọn {files.length} file</p>
            <ul className="mx-auto mt-2 max-w-sm space-y-1 text-left text-xs text-[#6B6963]">
              {Array.from(files).map((f, i) => (
                <li key={i} className="truncate rounded-lg bg-white px-3 py-1.5 ring-1 ring-[#ECEBE6]">
                  {f.name}
                </li>
              ))}
            </ul>
          </div>
        ) : (
          <>
            <p className="text-sm font-medium text-[#1C1C1A]">Kéo thả file vào đây, hoặc bấm để chọn</p>
            {hint && <p className="mt-1 text-xs text-[#8A887F]">{hint}</p>}
          </>
        )}
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          multiple={multiple}
          className="hidden"
          onChange={(e) => {
            if (e.target.files?.length) onFilesSelected(e.target.files);
            e.target.value = "";
          }}
        />
      </div>
      {files.length > 0 && (
        <button
          onClick={() => onFilesSelected(null)}
          className="mt-3 text-xs text-[#8A887F] underline decoration-[#DEDCD3] underline-offset-2 hover:text-[#1C1C1A]"
        >
          Chọn lại
        </button>
      )}
    </div>
  );
}