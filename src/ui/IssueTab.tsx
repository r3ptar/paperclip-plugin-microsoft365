import {
  usePluginData,
  type PluginDetailTabProps,
} from "@paperclipai/plugin-sdk/ui";
import { card, label, badge } from "./styles.js";
import type { IssueM365Data } from "./types.js";

export function M365IssueTab(props: PluginDetailTabProps) {
  const { context } = props;
  const { data, loading, error } = usePluginData<IssueM365Data>("issue-m365", {
    companyId: context.companyId,
    issueId: context.entityId,
  });

  if (loading) return <div style={{ padding: "16px" }}>Loading M365 data...</div>;
  if (error) return <div style={{ padding: "16px", color: "#dc2626" }}>Error: {error.message}</div>;

  const { plannerTask, calendarEvent } = data ?? { plannerTask: null, calendarEvent: null };

  return (
    <div style={{ padding: "16px" }}>
      <div style={card}>
        <div style={label}>Planner Task</div>
        {plannerTask ? (
          <div style={{ marginTop: "8px" }}>
            <div style={{ fontWeight: 500 }}>{plannerTask.title ?? "Untitled"}</div>
            <div style={{ fontSize: "13px", opacity: 0.6, color: "inherit", marginTop: "4px" }}>
              Status: <span style={badge("#2563eb")}>{plannerTask.status ?? "unknown"}</span>
            </div>
            <div style={{ fontSize: "12px", opacity: 0.5, color: "inherit", marginTop: "4px" }}>
              Last synced: {plannerTask.data?.lastSyncedAt
                ? new Date(plannerTask.data.lastSyncedAt).toLocaleString()
                : "Unknown"}
            </div>
          </div>
        ) : (
          <div style={{ marginTop: "8px", opacity: 0.5, color: "inherit", fontSize: "13px" }}>
            No linked Planner task
          </div>
        )}
      </div>

      <div style={card}>
        <div style={label}>Calendar Event</div>
        {calendarEvent ? (
          <div style={{ marginTop: "8px" }}>
            <div style={{ fontWeight: 500 }}>{calendarEvent.title ?? "Deadline"}</div>
            <div style={{ fontSize: "13px", opacity: 0.6, color: "inherit", marginTop: "4px" }}>
              Due: {calendarEvent.data?.dueDate ?? "—"}
            </div>
          </div>
        ) : (
          <div style={{ marginTop: "8px", opacity: 0.5, color: "inherit", fontSize: "13px" }}>
            No linked calendar event
          </div>
        )}
      </div>
    </div>
  );
}
