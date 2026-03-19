import type { PluginDetailTabProps } from "@paperclipai/plugin-sdk/ui";
import { card, label } from "./styles.js";

export function M365ProjectTab(props: PluginDetailTabProps) {
  const { context } = props;

  return (
    <div style={{ padding: "16px" }}>
      <div style={card}>
        <div style={label}>SharePoint Documents</div>
        <div style={{ marginTop: "8px", color: "#94a3b8", fontSize: "13px" }}>
          Use the <strong>sharepoint-search</strong> agent tool to search documents, or configure a SharePoint library in the plugin settings.
        </div>
        <div style={{ marginTop: "12px", fontSize: "12px", color: "#64748b" }}>
          Project: {context.entityId}
        </div>
      </div>
    </div>
  );
}
