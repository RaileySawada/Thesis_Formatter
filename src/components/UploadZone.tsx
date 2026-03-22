import { useRef, useState, useEffect } from "react";

interface Props {
  file: File | null;
  setFile: (f: File | null) => void;
}

export default function UploadZone({ file, setFile }: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragover, setDragover] = useState(false);
  const [dealt, setDealt] = useState(false);

  const handleFile = (f: File) => {
    if (!f.name.endsWith(".docx")) return;
    setFile(f);
  };

  useEffect(() => {
    if (file) {
      setDealt(false);
      // Trigger animation
      requestAnimationFrame(() => {
        requestAnimationFrame(() => setDealt(true));
      });
      const timer = setTimeout(() => setDealt(false), 700);
      return () => clearTimeout(timer);
    }
  }, [file]);

  return (
    <div
      id="drop-zone"
      className={`drop-zone mt-5 rounded-2xl border-2 border-dashed p-6 sm:p-8 transition-all${dragover ? " dragover" : ""}`}
      style={
        file
          ? { borderColor: "#10b981", background: "rgba(16,185,129,.06)" }
          : {
              borderColor: "var(--accent-muted)",
              background: "var(--accent-subtle)",
            }
      }
      onDragOver={(e) => {
        e.preventDefault();
        setDragover(true);
      }}
      onDragLeave={() => setDragover(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragover(false);
        const f = e.dataTransfer?.files?.[0];
        if (f) handleFile(f);
      }}
    >
      {!file ? (
        <label htmlFor="manuscript" className="block cursor-pointer">
          <div className="flex flex-col items-center justify-center text-center">
            <div
              className="flex h-14 w-14 items-center justify-center rounded-2xl shadow-sm sm:h-16 sm:w-16"
              style={{
                background: "var(--surface)",
                border: "1px solid var(--border)",
              }}
            >
              <i
                className="fa-solid fa-cloud-arrow-up text-2xl sm:text-3xl"
                style={{ color: "var(--accent)" }}
              />
            </div>
            <h3
              className="mt-4 text-base font-semibold sm:text-lg"
              style={{ color: "var(--text-primary)" }}
            >
              Upload manuscript
            </h3>
            <p
              className="mt-1.5 text-sm"
              style={{ color: "rgba(255,255,255,0.65)" }}
            >
              Drag &amp; drop your{" "}
              <span
                className="font-semibold"
                style={{ color: "var(--text-primary)" }}
              >
                .docx
              </span>{" "}
              file, or click to browse
            </p>
            <p
              className="mt-1 text-xs"
              style={{ color: "rgba(255,255,255,0.65)" }}
            >
              Microsoft Word Document (.docx)
            </p>
          </div>
          <input
            ref={inputRef}
            id="manuscript"
            type="file"
            accept=".docx"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) handleFile(f);
            }}
          />
        </label>
      ) : (
        <div className="flex flex-col items-center gap-4 w-full min-w-0">
          <div className={`file-cards-wrap${dealt ? " dealt" : ""}`}>
            <div className="file-card fc-p1">
              <div className="fc-lines">
                <div className="fc-line" style={{ width: "70%" }} />
                <div className="fc-line" style={{ width: "90%" }} />
                <div className="fc-line" style={{ width: "55%" }} />
              </div>
            </div>
            <div className="file-card fc-p2">
              <div className="fc-lines">
                <div className="fc-line" style={{ width: "85%" }} />
                <div className="fc-line" style={{ width: "60%" }} />
                <div className="fc-line" style={{ width: "75%" }} />
                <div className="fc-line" style={{ width: "40%" }} />
              </div>
            </div>
            <div className="file-card fc-p3">
              <i className="fa-solid fa-file-lines fc-doc-icon" />
              <div className="fc-lines" style={{ marginTop: 6 }}>
                <div className="fc-line" style={{ width: "80%" }} />
                <div className="fc-line" style={{ width: "65%" }} />
              </div>
              <span
                className="fc-filename"
                style={{
                  display: "block",
                  maxWidth: "100%",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                }}
                title={file.name}
              >
                {file.name}
              </span>
            </div>
          </div>
          <div
            className="text-center min-w-0"
            style={{ maxWidth: "16rem", width: "100%" }}
          >
            <p
              className="text-sm font-bold"
              style={{ color: "var(--text-primary)" }}
            >
              <i
                className="fa-solid fa-circle-check mr-1.5"
                style={{ color: "#10b981" }}
              />
              File ready
            </p>
            <p
              className="mt-0.5 text-xs truncate"
              style={{ color: "var(--text-muted)" }}
              title={file.name}
            >
              {file.name}
            </p>
          </div>
          <label htmlFor="manuscript-change" className="cursor-pointer">
            <span
              className="inline-flex items-center gap-1.5 rounded-xl border px-3 py-1.5 text-xs font-semibold transition hover:opacity-80"
              style={{
                borderColor: "var(--border)",
                background: "var(--surface)",
                color: "var(--text-secondary)",
              }}
            >
              <i className="fa-solid fa-arrow-up-from-bracket text-xs" /> Change
              file
            </span>
            <input
              id="manuscript-change"
              type="file"
              accept=".docx"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleFile(f);
              }}
            />
          </label>
        </div>
      )}
    </div>
  );
}
