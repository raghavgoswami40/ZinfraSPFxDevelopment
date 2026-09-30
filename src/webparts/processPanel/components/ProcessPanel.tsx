import * as React from 'react';
import styles from './ProcessPanel.module.scss';
import ProcessDetailsSection from './ProcessDetailsSection';
import RasciSection from './RasciSection';
import { IProcessDetailsSectionProps, IRasciSectionProps } from './IProcessPanelProps';

export interface IProcessPanelProps {
  processDetails: IProcessDetailsSectionProps;
  rasci: IRasciSectionProps;
}

/**
 * The two sections stacked in document order — Process Details, then RASCI —
 * inside one wrapper. StickyPanelController pins this whole wrapper as a single
 * unit; it expects the sole child of the web part's domElement, which is
 * exactly what renders here.
 */
const ProcessPanel: React.FC<IProcessPanelProps> = ({ processDetails, rasci }) => (
  <div className={styles.wrapper}>
    <ProcessDetailsSection {...processDetails} />
    <RasciSection {...rasci} />
  </div>
);

export default ProcessPanel;
