export interface StoredStep {
  id: string;
  label: string;
  shape?: 'box' | 'approval' | 'sapStatus' | 'gate' | 'group';
  noPrefix?: boolean; // sapStatus shape only - suppress the fixed "SAP Status" first line
  children?: StoredStep[]; // 'group' shape only — steps rendered inside the dashed container
}

export interface StoredSection {
  id: string;
  title: string;
  steps: StoredStep[];
  variant?: 'blue' | 'grey';
  // Prefix for the auto-generated step code badges, e.g. "2.1" -> 2.1.1, 2.1.2...
  // Omit when the section title already starts with the number (it is derived);
  // required for sections with no visible heading.
  codePrefix?: string;
}

export interface StoredPhase {
  id: string;
  title: string;
  sections: StoredSection[];
}

export interface IVerticalFlowProps {
  phases: StoredPhase[];
  isEditMode: boolean;
  onPhasesChange: (phases: StoredPhase[]) => void;
}
