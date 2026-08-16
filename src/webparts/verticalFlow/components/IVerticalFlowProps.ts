export interface StoredStep {
  id: string;
  label: string;
  isApproval?: boolean;
}

export interface StoredSection {
  id: string;
  title: string;
  numRows: number;
  steps: StoredStep[];
}

export interface StoredPhase {
  id: string;
  title: string;
  sections: StoredSection[];
}

export interface IVerticalFlowProps {
  urls: Record<string, string>;
  phases: StoredPhase[];
  isEditMode: boolean;
  onPhasesChange: (phases: StoredPhase[]) => void;
}
