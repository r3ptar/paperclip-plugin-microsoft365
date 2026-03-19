import type React from "react";
import { card, label } from "../styles.js";

export interface ServiceCardProps {
  name: string;
  description: string;
  permissions: string[];
  enabled: boolean;
  onToggle: (enabled: boolean) => void;
}

const enabledCard: React.CSSProperties = {
  ...card,
  borderColor: "#2563eb",
  backgroundColor: "rgba(37, 99, 235, 0.08)",
};

const serviceHeader: React.CSSProperties = {
  display: "flex",
  justifyContent: "space-between",
  alignItems: "center",
  marginBottom: "8px",
};

const serviceName: React.CSSProperties = {
  fontSize: "15px",
  fontWeight: 600,
  color: "inherit",
  display: "flex",
  alignItems: "center",
  gap: "8px",
};

const descStyle: React.CSSProperties = {
  fontSize: "13px",
  opacity: 0.7,
  color: "inherit",
  marginBottom: "10px",
};

const permissionBadge: React.CSSProperties = {
  display: "inline-block",
  padding: "2px 8px",
  borderRadius: "4px",
  fontSize: "11px",
  fontWeight: 500,
  backgroundColor: "rgba(128, 128, 128, 0.15)",
  color: "inherit",
  marginRight: "6px",
  marginBottom: "4px",
};

const checkmark: React.CSSProperties = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  width: "20px",
  height: "20px",
  borderRadius: "50%",
  backgroundColor: "#16a34a",
  color: "#fff",
  fontSize: "12px",
  fontWeight: 700,
};

export function ServiceCard(props: ServiceCardProps) {
  const { name, description, permissions, enabled, onToggle } = props;

  return (
    <div style={enabled ? enabledCard : card}>
      <div style={serviceHeader}>
        <div style={serviceName}>
          {enabled && <span style={checkmark}>&#10003;</span>}
          {name}
        </div>
        <label style={{ display: "flex", alignItems: "center", gap: "6px", cursor: "pointer" }}>
          <input
            type="checkbox"
            checked={enabled}
            onChange={(e) => onToggle(e.target.checked)}
          />
          <span style={{ fontSize: "13px", fontWeight: 500, color: "inherit", userSelect: "none" }}>
            {enabled ? "Enabled" : "Disabled"}
          </span>
        </label>
      </div>

      <div style={descStyle}>{description}</div>

      <div style={label}>Required Permissions</div>
      <div style={{ marginTop: "6px" }}>
        {permissions.map((perm) => (
          <span key={perm} style={permissionBadge}>
            {perm}
          </span>
        ))}
      </div>
    </div>
  );
}
