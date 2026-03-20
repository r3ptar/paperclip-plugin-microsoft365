import { useState, useCallback } from "react";
import type React from "react";
import { usePluginData } from "@paperclipai/plugin-sdk/ui";
import { textInput, selectInput, secondaryButton, errorText } from "../styles.js";

export interface AgentIdentityEditorProps {
  entries: Record<string, string>;
  onChange: (entries: Record<string, string>) => void;
  companyId: string | null;
}

type DropdownData = {
  items?: Array<{ id: string; name: string }>;
  error?: string;
};

interface Row {
  id: string;
  agentId: string;
  m365UserId: string;
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
  return `aid-${nextId++}`;
}

export function AgentIdentityEditor(props: AgentIdentityEditorProps) {
  const { entries, onChange, companyId } = props;

  const [rows, setRows] = useState<Row[]>(() =>
    Object.entries(entries).map(([agentId, m365UserId]) => ({
      id: makeId(),
      agentId,
      m365UserId,
    })),
  );

  // Fetch agents and users for dropdowns
  const { data: agentsData, loading: agentsLoading } = usePluginData<DropdownData>(
    "paperclip-agents",
    companyId ? { companyId } : undefined,
  );
  const { data: usersData, loading: usersLoading } = usePluginData<DropdownData>(
    "m365-users",
    companyId ? { companyId } : undefined,
  );

  const agents = agentsData?.items ?? [];
  const users = usersData?.items ?? [];

  const emitChange = useCallback(
    (updated: Row[]) => {
      setRows(updated);
      const result: Record<string, string> = {};
      for (const row of updated) {
        if (row.agentId && row.m365UserId) {
          result[row.agentId] = row.m365UserId;
        }
      }
      onChange(result);
    },
    [onChange],
  );

  const handleAgentChange = useCallback(
    (rowId: string, agentId: string) => {
      emitChange(rows.map((r) => (r.id === rowId ? { ...r, agentId } : r)));
    },
    [rows, emitChange],
  );

  const handleUserChange = useCallback(
    (rowId: string, m365UserId: string) => {
      emitChange(rows.map((r) => (r.id === rowId ? { ...r, m365UserId } : r)));
    },
    [rows, emitChange],
  );

  const handleRemove = useCallback(
    (rowId: string) => {
      emitChange(rows.filter((r) => r.id !== rowId));
    },
    [rows, emitChange],
  );

  const handleAdd = useCallback(() => {
    setRows((prev) => [...prev, { id: makeId(), agentId: "", m365UserId: "" }]);
  }, []);

  return (
    <div>
      {agentsData?.error && <span style={errorText}>{agentsData.error}</span>}
      {usersData?.error && <span style={errorText}>{usersData.error}</span>}

      {rows.map((row) => (
        <div key={row.id} style={rowContainer}>
          <select
            style={{ ...selectInput, flex: 1 }}
            value={row.agentId}
            onChange={(e) => handleAgentChange(row.id, e.target.value)}
            disabled={agentsLoading}
          >
            <option value="">
              {agentsLoading ? "Loading agents..." : "Select agent..."}
            </option>
            {agents.map((a) => (
              <option key={a.id} value={a.id}>
                {a.name}
              </option>
            ))}
            {/* Keep current value visible even if not in the list */}
            {row.agentId && !agents.find((a) => a.id === row.agentId) && (
              <option value={row.agentId}>{row.agentId}</option>
            )}
          </select>

          <select
            style={{ ...selectInput, flex: 1 }}
            value={row.m365UserId}
            onChange={(e) => handleUserChange(row.id, e.target.value)}
            disabled={usersLoading}
          >
            <option value="">
              {usersLoading ? "Loading users..." : "Select M365 user..."}
            </option>
            {users.map((u) => (
              <option key={u.id} value={u.id}>
                {u.name}
              </option>
            ))}
            {row.m365UserId && !users.find((u) => u.id === row.m365UserId) && (
              <option value={row.m365UserId}>{row.m365UserId}</option>
            )}
          </select>

          <button
            type="button"
            style={removeBtn}
            onClick={() => handleRemove(row.id)}
            title="Remove mapping"
          >
            &#215;
          </button>
        </div>
      ))}

      <button type="button" style={secondaryButton} onClick={handleAdd}>
        Add Agent Mapping
      </button>
    </div>
  );
}
