import React from "react";

interface ToastProps {
  msg: string;
  onDismiss: () => void;
  type?: "success" | "error";
  actionLabel?: string;
  onAction?: () => void;
}

export default function Toast({
  msg,
  onDismiss,
  type = "success",
  actionLabel,
  onAction,
}: ToastProps) {
  const isError = type === "error";

  const handleAction = () => {
    if (onAction) onAction();
    onDismiss();
  };

  return (
    <div className={`tf-toast tf-toast--${type}`} role="alert">
      <div className="tf-toast-icon-wrap">
        {isError ? (
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
            <path
              d="M8 5v4M8 11v.5"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
            />
            <circle
              cx="8"
              cy="8"
              r="6.5"
              stroke="currentColor"
              strokeWidth="1.5"
            />
          </svg>
        ) : (
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
            <path
              d="M3 8.5L6.5 12L13 5"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        )}
      </div>
      <div className="tf-toast-body">
        <p className="tf-toast-title">
          {isError ? "Something went wrong" : "System Notification"}
        </p>
        <p className="tf-toast-msg">{msg}</p>
      </div>

      {actionLabel && onAction && (
        <button
          className="ml-4 px-3 py-1.5 rounded-lg bg-white/20 hover:bg-white/30 text-white text-[10px] font-bold uppercase tracking-wider transition-colors active:scale-95 border border-white/10"
          onClick={handleAction}
        >
          {actionLabel}
        </button>
      )}

      <button
        className="tf-toast-close"
        onClick={onDismiss}
        aria-label="Dismiss"
      >
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
          <path
            d="M2 2L12 12M12 2L2 12"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
          />
        </svg>
      </button>
      <div className="tf-toast-bar" />
    </div>
  );
}
