import { useState, useCallback } from "react";
import type React from "react";
import { textInput, errorText } from "../styles.js";

export interface EmailChipsProps {
  emails: string[];
  onChange: (emails: string[]) => void;
}

const chipContainer: React.CSSProperties = {
  display: "flex",
  flexWrap: "wrap",
  gap: "6px",
  marginBottom: "8px",
};

const chip: React.CSSProperties = {
  display: "inline-flex",
  alignItems: "center",
  gap: "4px",
  padding: "3px 10px",
  borderRadius: "16px",
  fontSize: "13px",
  fontWeight: 500,
  backgroundColor: "#e0e7ff",
  color: "#3730a3",
};

const removeButton: React.CSSProperties = {
  border: "none",
  background: "none",
  cursor: "pointer",
  fontSize: "14px",
  fontWeight: 700,
  color: "#6366f1",
  padding: "0 2px",
  lineHeight: "1",
};

const addRow: React.CSSProperties = {
  display: "flex",
  gap: "8px",
  alignItems: "center",
};

const addButton: React.CSSProperties = {
  padding: "6px 14px",
  borderRadius: "6px",
  border: "1px solid #e2e8f0",
  backgroundColor: "#f8fafc",
  fontSize: "13px",
  fontWeight: 500,
  cursor: "pointer",
  whiteSpace: "nowrap",
};

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function EmailChips(props: EmailChipsProps) {
  const { emails, onChange } = props;

  const [inputValue, setInputValue] = useState("");
  const [validationError, setValidationError] = useState("");

  const addEmail = useCallback(() => {
    const trimmed = inputValue.trim();
    if (!trimmed) return;

    if (!EMAIL_REGEX.test(trimmed)) {
      setValidationError("Please enter a valid email address");
      return;
    }

    if (emails.includes(trimmed)) {
      setValidationError("This email has already been added");
      return;
    }

    onChange([...emails, trimmed]);
    setInputValue("");
    setValidationError("");
  }, [inputValue, emails, onChange]);

  const removeEmail = useCallback(
    (email: string) => {
      onChange(emails.filter((e) => e !== email));
    },
    [emails, onChange],
  );

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === "Enter") {
        e.preventDefault();
        addEmail();
      }
    },
    [addEmail],
  );

  return (
    <div>
      {/* Existing chips */}
      {emails.length > 0 && (
        <div style={chipContainer}>
          {emails.map((email) => (
            <span key={email} style={chip}>
              {email}
              <button
                type="button"
                style={removeButton}
                onClick={() => removeEmail(email)}
                title={`Remove ${email}`}
              >
                &#215;
              </button>
            </span>
          ))}
        </div>
      )}

      {/* Add input */}
      <div style={addRow}>
        <input
          type="email"
          style={{ ...textInput, flex: 1 }}
          placeholder="user@example.com"
          value={inputValue}
          onChange={(e) => {
            setInputValue(e.target.value);
            setValidationError("");
          }}
          onKeyDown={handleKeyDown}
        />
        <button type="button" style={addButton} onClick={addEmail}>
          Add
        </button>
      </div>

      {/* Validation error */}
      {validationError && (
        <span style={errorText}>{validationError}</span>
      )}
    </div>
  );
}
