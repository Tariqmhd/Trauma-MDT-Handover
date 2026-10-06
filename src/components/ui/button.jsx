import React from "react";

const cn = (...classes) => classes.filter(Boolean).join(" ");

export function Button({ className, variant = "default", type = "button", ...props }) {
  const variants = {
    default: "bg-slate-900 text-white hover:bg-slate-800",
    outline: "border border-slate-200 bg-white hover:bg-slate-50 text-slate-900",
    ghost: "hover:bg-slate-100 text-slate-900"
  };
  return (
    <button
      type={type}
      className={cn("inline-flex h-10 items-center justify-center rounded-md px-4 py-2 text-sm font-medium transition-colors disabled:pointer-events-none disabled:opacity-50", variants[variant] || variants.default, className)}
      {...props}
    />
  );
}
