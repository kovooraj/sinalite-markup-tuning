"use client";

import { useEffect, useRef, useState } from "react";

interface Props {
  text: string;
  align?: "left" | "right";
}

export function InfoIcon({ text, align = "left" }: Props) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLSpanElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  return (
    <span
      ref={ref}
      className="relative ml-1.5 inline-block align-middle"
      onMouseEnter={() => setOpen(true)}
      onMouseLeave={() => setOpen(false)}
    >
      <button
        type="button"
        onClick={(e) => {
          e.stopPropagation();
          e.preventDefault();
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
        onBlur={() => setOpen(false)}
        className="inline-flex h-4 w-4 items-center justify-center rounded-full border border-zinc-300 bg-white text-[10px] font-semibold text-zinc-500 leading-none hover:border-blue-500 hover:text-blue-600"
        aria-label="More info"
        title={text}
      >
        i
      </button>
      {open && (
        <span
          className={[
            "absolute top-full z-20 mt-1 block w-72 whitespace-normal rounded border border-zinc-200 bg-white p-2 text-left text-xs font-normal text-zinc-700 shadow-lg",
            align === "right" ? "right-0" : "left-0",
          ].join(" ")}
          onClick={(e) => e.stopPropagation()}
        >
          {text}
        </span>
      )}
    </span>
  );
}
