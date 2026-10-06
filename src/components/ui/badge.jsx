import React from "react";

const cn = (...classes) => classes.filter(Boolean).join(" ");

export function Badge({ className, variant = "default", ...props }) {
  const variants = {
    default: "border-transparent bg-slate-900 text-white",
    outline: "border-slate-200 bg-white text-slate-700"
  };
  return <span className={cn("inline-flex items-center rounded-full border px-2.5 py-0.5 text-xs font-semibold transition-colors", variants[variant] || variants.default, className)} {...props} />;
}
