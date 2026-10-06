import { useState } from "react";
import { FileText, Image as ImageIcon, FileType2, ShieldCheck } from "lucide-react";
import PdfToWord from "./components/PdfToWord";
import ImageToPdf from "./components/ImageToPdf";
import WordToPdf from "./components/WordToPdf";

const TABS = [
  { key: "pdf2word", label: "PDF → Word", desc: "Lấy văn bản để chỉnh sửa", icon: FileText, accent: "#C4432E", Component: PdfToWord },
  { key: "img2pdf", label: "Ảnh → PDF", desc: "Gộp nhiều ảnh thành một file", icon: ImageIcon, accent: "#2F7D5D", Component: ImageToPdf },
  { key: "word2pdf", label: "Word → PDF", desc: "Xuất file .docx sang PDF", icon: FileType2, accent: "#2B5797", Component: WordToPdf },
];

export default function App() {
  const [active, setActive] = useState(TABS[0].key);
  const activeTab = TABS.find((t) => t.key === active);

  return (
    <div className="min-h-screen bg-white px-4 py-12 sm:px-8">
      <div className="mx-auto max-w-3xl">
        <header className="mb-8">
          <h1 className="text-3xl font-semibold tracking-tight text-[#1C1C1A]">CHUYỂN ĐỔI FILE</h1>
          <p className="mt-2 inline-flex items-center gap-1.5 text-sm text-[#6B6963]">
            <ShieldCheck size={15} className="text-[#2F7D5D]" />
            Xử lý ngay trên máy của bạn, không tải lên máy chủ nào.
          </p>
        </header>

        <div role="tablist" className="grid gap-3 sm:grid-cols-3">
          {TABS.map((tab) => {
            const isActive = tab.key === active;
            const Icon = tab.icon;
            return (
              <button
                key={tab.key}
                role="tab"
                aria-selected={isActive}
                onClick={() => setActive(tab.key)}
                className="rounded-2xl border bg-white p-4 text-left transition-all focus-visible:outline-2 focus-visible:outline-offset-2"
                style={{
                  borderColor: isActive ? tab.accent : "#ECEBE6",
                  boxShadow: isActive ? `0 0 0 3px ${tab.accent}1A` : "none",
                  outlineColor: tab.accent,
                }}
              >
                <span
                  className="flex h-9 w-9 items-center justify-center rounded-xl"
                  style={{ backgroundColor: `${tab.accent}14`, color: tab.accent }}
                >
                  <Icon size={18} />
                </span>
                <span className="mt-3 block text-sm font-semibold text-[#1C1C1A]">{tab.label}</span>
                <span className="mt-0.5 block text-xs text-[#8A887F]">{tab.desc}</span>
              </button>
            );
          })}
        </div>

        <div
          role="tabpanel"
          className="mt-5 rounded-3xl border border-[#ECEBE6] bg-white p-6 sm:p-8"
          style={{ boxShadow: "0 8px 30px -12px rgba(28,28,26,0.12)" }}
        >
          <activeTab.Component />
        </div>
      </div>
    </div>
  );
}