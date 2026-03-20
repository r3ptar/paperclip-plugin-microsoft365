import { useState, useCallback, useEffect, useRef } from "react";
import type React from "react";
import { textInput, secondaryButton } from "../styles.js";

export interface KeyValueEditorProps {
  entries: Record<string, string>;
  onChange: (entries: Record<string, string>) => void;
  keyPlaceholder?: string;
  valuePlaceholder?: string;
}

interface Row {
  id: string;
  key: string;
  value: string;
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

let nextId = 1;
function makeId(): string {
  return `kv-${nextId++}`;
}

function entriesToRows(entries: Record<string, string>): Row[] {
  return Object.entries(entries).map(([key, value]) => ({
    id: makeId(),
    key,
    value,
  }));
}

function rowsToEntries(rows: Row[]): Record<string, string> {
  const result: Record<string, string> = {};
  for (const row of rows) {
    // Skip rows with empty keys — they are placeholders
    if (row.key.trim()) {
      result[row.key] = row.value;
    }
  }
  return result;
}

export function KeyValueEditor(props: KeyValueEditorProps) {
  const { entries, onChange, keyPlaceholder, valuePlaceholder } = props;

  const [rows, setRows] = useState<Row[]>(() => entriesToRows(entries));
  const lastEntriesRef = useRef(entries);

  // Sync from parent only when the entries reference changes externally
  useEffect(() => {
    if (entries !== lastEntriesRef.current) {
      lastEntriesRef.current = entries;
      setRows(entriesToRows(entries));
    }
  }, [entries]);

  const emitChange = useCallback(
    (updated: Row[]) => {
      setRows(updated);
      const newEntries = rowsToEntries(updated);
      lastEntriesRef.current = newEntries;
      onChange(newEntries);
    },
    [onChange],
  );

  const handleKeyChange = useCallback(
    (id: string, newKey: string) => {
      emitChange(rows.map((r) => (r.id === id ? { ...r, key: newKey } : r)));
    },
    [rows, emitChange],
  );

  const handleValueChange = useCallback(
    (id: string, value: string) => {
      emitChange(rows.map((r) => (r.id === id ? { ...r, value } : r)));
    },
    [rows, emitChange],
  );

  const handleRemove = useCallback(
    (id: string) => {
      emitChange(rows.filter((r) => r.id !== id));
    },
    [rows, emitChange],
  );

  const handleAdd = useCallback(() => {
    const updated = [...rows, { id: makeId(), key: "", value: "" }];
    setRows(updated);
    // Don't emit yet — empty keys are filtered out by rowsToEntries
  }, [rows]);

  const hasDuplicateKey = (key: string, id: string): boolean => {
    if (!key.trim()) return false;
    return rows.some((r) => r.id !== id && r.key === key);
  };

  return (
    <div>
      {rows.map((row) => {
        const isDup = hasDuplicateKey(row.key, row.id);
        return (
          <div key={row.id} style={rowContainer}>
            <input
              type="text"
              style={{
                ...textInput,
                flex: 1,
                ...(isDup ? { borderColor: "#dc2626" } : {}),
              }}
              placeholder={keyPlaceholder ?? "Key"}
              value={row.key}
              onChange={(e) => handleKeyChange(row.id, e.target.value)}
            />
            <input
              type="text"
              style={{ ...textInput, flex: 1 }}
              placeholder={valuePlaceholder ?? "Value"}
              value={row.value}
              onChange={(e) => handleValueChange(row.id, e.target.value)}
            />
            <button
              type="button"
              style={removeBtn}
              onClick={() => handleRemove(row.id)}
              title="Remove entry"
            >
              &#215;
            </button>
          </div>
        );
      })}
      <button type="button" style={secondaryButton} onClick={handleAdd}>
        Add
      </button>
    </div>
  );
}
