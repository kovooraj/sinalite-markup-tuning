"use client";

import { useCallback } from "react";
import { useDropzone } from "react-dropzone";

interface Props {
  label: string;
  hint: string;
  file: File | null;
  onSelect: (file: File | null) => void;
  preview?: string;
  error?: string;
}

export function FileDrop({ label, hint, file, onSelect, preview, error }: Props) {
  const onDrop = useCallback(
    (accepted: File[]) => {
      if (accepted[0]) onSelect(accepted[0]);
    },
    [onSelect]
  );
  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    multiple: false,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
    },
  });

  return (
    <div
      {...getRootProps()}
      className={[
        "flex flex-col items-center justify-center rounded-lg border-2 border-dashed p-6 text-center transition-colors cursor-pointer",
        isDragActive ? "border-blue-500 bg-blue-50" : "border-zinc-300 bg-zinc-50 hover:border-zinc-400",
        error ? "border-red-400 bg-red-50" : "",
      ].join(" ")}
    >
      <input {...getInputProps()} />
      <div className="text-sm font-semibold text-zinc-800">{label}</div>
      <div className="mt-1 text-xs text-zinc-500">{hint}</div>
      {file ? (
        <div className="mt-3 w-full">
          <div className="truncate text-sm text-zinc-900">{file.name}</div>
          <div className="text-xs text-zinc-500">
            {(file.size / 1024 / 1024).toFixed(2)} MB
          </div>
          {preview && (
            <div className="mt-1 text-xs text-zinc-600">{preview}</div>
          )}
          <button
            type="button"
            className="mt-2 text-xs text-blue-600 hover:underline"
            onClick={(e) => {
              e.stopPropagation();
              onSelect(null);
            }}
          >
            Remove
          </button>
        </div>
      ) : (
        <div className="mt-3 text-xs text-zinc-400">Drop or click to select</div>
      )}
      {error && <div className="mt-2 text-xs text-red-600">{error}</div>}
    </div>
  );
}
