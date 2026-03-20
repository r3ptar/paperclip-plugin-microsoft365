import { useCallback } from "react";
import type React from "react";
import { textInput, fieldRow, secondaryButton } from "../styles.js";

export interface KeyValueEditorProps {
  entries: Record<string, string>;
  onChange: (entries: Record<string, string>) => void;
  keyPlaceholder?: string;
  valuePlaceholder?: string;
}

const rowContainer: React.CSSProperties = {
  display: "flex",
  alignItems: "center",
  gap: "8px",
  marginBottom: "8px",
};

const removeBtn: React.CSSProperties = {
  border: "none",
  background: "none",
  cursor: "pointer",
  fontSize: "16px",
  fontWeight: 700,
  color: "var(--muted-foreground)",
  padding: "0 4px",
  lineHeight: "1",
};

export function KeyValueEditor(props: KeyValueEditorProps) {
  const { entries, onChange, keyPlaceholder, valuePlaceholder } = props;

  const keys = Object.keys(entries);

  const handleKeyChange = useCallback(
    (oldKey: string, newKey: string) => {
      const updated: Record<string, string> = {};
      for (const k of Object.keys(entries)) {
        if (k === oldKey) {
          updated[newKey] = entries[k];
        } else {
          updated[k] = entries[k];
        }
      }
      onChange(updated);
    },
    [entries, onChange],
  );

  const handleValueChange = useCallback(
    (key: string, value: string) => {
      onChange({ ...entries, [key]: value });
    },
    [entries, onChange],
  );

  const handleRemove = useCallback(
    (key: string) => {
      const updated = { ...entries };
      delete updated[key];
      onChange(updated);
    },
    [entries, onChange],
  );

  const handleAdd = useCallback(() => {
    // Find a unique empty key
    let newKey = "";
    let idx = 0;
    while (newKey in entries) {
      idx++;
      newKey = `key-${idx}`;
    }
    onChange({ ...entries, [newKey]: "" });
  }, [entries, onChange]);

  return (
    <div>
      {keys.map((key, index) => (
        <div key={index} style={rowContainer}>
          <input
            type="text"
            style={{ ...textInput, flex: 1 }}
            placeholder={keyPlaceholder ?? "Key"}
            value={key}
            onChange={(e) => handleKeyChange(key, e.target.value)}
          />
          <input
            type="text"
            style={{ ...textInput, flex: 1 }}
            placeholder={valuePlaceholder ?? "Value"}
            value={entries[key]}
            onChange={(e) => handleValueChange(key, e.target.value)}
          />
          <button
            type="button"
            style={removeBtn}
            onClick={() => handleRemove(key)}
            title="Remove entry"
          >
            &#215;
          </button>
        </div>
      ))}
      <button type="button" style={secondaryButton} onClick={handleAdd}>
        Add
      </button>
    </div>
  );
}
