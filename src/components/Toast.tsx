interface ToastProps {
  msg: string;
  onDismiss: () => void;
  type?: "success" | "error";
}

export default function Toast({
  msg,
  onDismiss,
  type = "success",
}: ToastProps) {
  const isError = type === "error";
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
          {isError ? "Something went wrong" : "Formatting complete"}
        </p>
        <p className="tf-toast-msg">{msg}</p>
      </div>
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
