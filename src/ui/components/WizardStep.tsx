import type React from "react";
import {
  card,
  label,
  primaryButton,
  primaryButtonDisabled,
  secondaryButton,
  secondaryButtonDisabled,
} from "../styles.js";

export interface WizardStepProps {
  title: string;
  description?: string;
  stepNumber: number;
  totalSteps: number;
  canProceed: boolean;
  onNext: () => void;
  onBack?: () => void;
  children: React.ReactNode;
}

const stepIndicator: React.CSSProperties = {
  fontSize: "12px",
  opacity: 0.6,
  color: "inherit",
  fontWeight: 600,
  textTransform: "uppercase",
  letterSpacing: "0.05em",
  marginBottom: "4px",
};

const titleStyle: React.CSSProperties = {
  fontSize: "18px",
  fontWeight: 600,
  color: "inherit",
  margin: "0 0 4px",
};

const descriptionStyle: React.CSSProperties = {
  fontSize: "14px",
  opacity: 0.6,
  color: "inherit",
  margin: "0 0 16px",
};

const navRow: React.CSSProperties = {
  display: "flex",
  justifyContent: "space-between",
  alignItems: "center",
  marginTop: "20px",
  paddingTop: "16px",
  borderTop: "1px solid rgba(128, 128, 128, 0.2)",
};

const progressBar: React.CSSProperties = {
  height: "4px",
  borderRadius: "2px",
  backgroundColor: "rgba(128, 128, 128, 0.2)",
  marginBottom: "16px",
  overflow: "hidden",
};

export function WizardStep(props: WizardStepProps) {
  const {
    title,
    description,
    stepNumber,
    totalSteps,
    canProceed,
    onNext,
    onBack,
    children,
  } = props;

  const isLastStep = stepNumber === totalSteps;
  const nextLabel = isLastStep ? "Save & Activate" : "Next";
  const progressPct = (stepNumber / totalSteps) * 100;

  return (
    <div style={card}>
      {/* Progress bar */}
      <div style={progressBar}>
        <div
          style={{
            height: "100%",
            width: `${progressPct}%`,
            backgroundColor: "#2563eb",
            borderRadius: "2px",
            transition: "width 0.3s ease",
          }}
        />
      </div>

      {/* Step indicator */}
      <div style={stepIndicator}>
        Step {stepNumber} of {totalSteps}
      </div>

      {/* Title */}
      <h3 style={titleStyle}>{title}</h3>

      {/* Description */}
      {description && <p style={descriptionStyle}>{description}</p>}

      {/* Step content */}
      <div style={{ marginTop: description ? "0" : "12px" }}>{children}</div>

      {/* Navigation */}
      <div style={navRow}>
        <div>
          {onBack && (
            <button style={secondaryButton} onClick={onBack}>
              Back
            </button>
          )}
        </div>
        <button
          disabled={!canProceed}
          onClick={onNext}
          style={canProceed ? primaryButton : primaryButtonDisabled}
        >
          {nextLabel}
        </button>
      </div>
    </div>
  );
}
