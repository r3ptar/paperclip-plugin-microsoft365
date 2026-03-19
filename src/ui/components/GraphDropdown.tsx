import { usePluginData } from "@paperclipai/plugin-sdk/ui";
import {
  fieldRow,
  fieldLabel,
  selectInput,
  errorText,
} from "../styles.js";

export interface GraphDropdownProps {
  label: string;
  dataHandler: string;
  params?: Record<string, string>;
  value: string;
  onChange: (id: string, name: string) => void;
  disabled?: boolean;
  companyId: string | null;
  placeholder?: string;
  /** Wizard credentials passed through to data handlers before config is saved */
  credentials?: { tenantId: string; clientId: string; clientSecret: string };
}

type GraphDropdownData = {
  items?: Array<{ id: string; name: string }>;
  error?: string;
};

export function GraphDropdown(props: GraphDropdownProps) {
  const {
    label: labelText,
    dataHandler,
    params,
    value,
    onChange,
    disabled,
    companyId,
    placeholder,
    credentials,
  } = props;

  const fetchParams = disabled
    ? undefined
    : { companyId, ...params, ...credentials };
  const { data, loading, error } = usePluginData<GraphDropdownData>(
    dataHandler,
    fetchParams,
  );

  const items = data?.items ?? [];
  const dataError = disabled ? null : (data?.error ?? null);
  const isDisabled = disabled || loading;

  const handleChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const selectedId = e.target.value;
    const item = items.find((i) => i.id === selectedId);
    onChange(selectedId, item?.name ?? "");
  };

  return (
    <div style={fieldRow}>
      <span style={fieldLabel}>{labelText}</span>

      {loading && (
        <span style={{ fontSize: "13px", color: "var(--muted-foreground)" }}>Loading...</span>
      )}

      {error && (
        <span style={errorText}>
          Error: {error.message}
        </span>
      )}

      {dataError && (
        <span style={errorText}>
          {dataError}
        </span>
      )}

      <select
        style={{
          ...selectInput,
          opacity: isDisabled ? 0.6 : 1,
        }}
        value={value}
        onChange={handleChange}
        disabled={isDisabled}
      >
        <option value="">
          {loading
            ? "Loading..."
            : items.length === 0 && !error && !dataError
              ? "No items found"
              : placeholder ?? "Select..."}
        </option>
        {items.map((item) => (
          <option key={item.id} value={item.id}>
            {item.name}
          </option>
        ))}
      </select>
    </div>
  );
}
