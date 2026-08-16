import { StoredPhase } from './IVerticalFlowProps';

export const DEFAULT_PHASES: StoredPhase[] = [
  {
    id: 'phase_1', title: '1  INITIATE', sections: [
      {
        id: 'sec_1_1', title: '1.1\nCOMMISSION', numRows: 1, steps: [
          { id: 'step_1_1_1', label: '1.1.1 Early Contractor Involvement' },
        ],
      },
    ],
  },
  {
    id: 'phase_2', title: '2  PLAN', sections: [
      {
        id: 'sec_2_1', title: '2.1\nESTABLISH', numRows: 1, steps: [
          { id: 'step_2_1_1', label: '2.1.1 Receive Project' },
          { id: 'step_2_1_2', label: '2.1.2 Assign Project Manager' },
          { id: 'step_2_1_3', label: '2.1.3 Define Project' },
          { id: 'step_2_1_4', label: '2.1.4 Assign Project Team' },
          { id: 'step_2_1_5', label: '2.1.5 Develop Project Plans and Establish Systems' },
        ],
      },
      {
        id: 'sec_2_2', title: '2.2\nASSESS', numRows: 1, steps: [
          { id: 'step_2_2_1', label: '2.2.1 Perform Investigations' },
          { id: 'step_2_2_2', label: '2.2.2 Engage Early Works Vendor(s)' },
          { id: 'step_2_2_3', label: '2.2.3 Execute Early Works' },
        ],
      },
      {
        id: 'sec_2_3', title: '2.3\nDESIGN', numRows: 1, steps: [
          { id: 'step_2_3_1', label: '2.3.1 Develop Design Plans' },
          { id: 'step_2_3_2', label: '2.3.2 Engage Designer(s)' },
          { id: 'step_2_3_3', label: '2.3.3 Execute Design' },
        ],
      },
      {
        id: 'sec_2_4', title: '2.4\nFORMULATE', numRows: 1, steps: [
          { id: 'step_2_4_1', label: '2.4.1 Develop Delivery Methodology' },
          { id: 'step_2_4_2', label: '2.4.2 Secure Requirements' },
        ],
      },
    ],
  },
  {
    id: 'phase_3', title: '3  DELIVER', sections: [
      {
        id: 'sec_3_1', title: '3.1\nPROCURE', numRows: 1, steps: [
          { id: 'step_3_1_1', label: '3.1.1 Engage Supplier(s) and Constructor(s)' },
          { id: 'step_3_1_2', label: '3.1.2 Procure Long Lead Material and Equipment' },
        ],
      },
      {
        id: 'sec_3_2', title: '3.2\nPREPARE', numRows: 1, steps: [
          { id: 'step_3_2_1', label: '3.2.1 Finalise Planning' },
          { id: 'step_3_2_2', label: '3.2.2 Perform Pre-Construction Reviews' },
          { id: 'step_3_2_3', label: '3.2.3 Develop Construction/ Commissioning Plans' },
          { id: 'step_3_2_4', label: '3.2.4 Onboard Site Workgroup' },
        ],
      },
      {
        id: 'sec_3_3', title: '3.3\nMOBILISE', numRows: 1, steps: [
          { id: 'step_3_3_1', label: '3.3.1 Set-Up Site' },
          { id: 'step_3_3_2', label: '3.3.2 Deploy Resources' },
        ],
      },
      {
        id: 'sec_3_4', title: '3.4\nCONSTRUCT', numRows: 1, steps: [
          { id: 'step_3_4_1', label: '3.4.1 Execute Construction' },
          { id: 'step_3_4_2', label: '3.4.2 Perform Completions' },
          { id: 'step_3_4_3', label: '3.4.3 Approval', isApproval: true },
        ],
      },
      {
        id: 'sec_3_5', title: '3.5\nCOMMISSION', numRows: 1, steps: [
          { id: 'step_3_5_1', label: '3.5.1 Execute Commissioning' },
          { id: 'step_3_5_2', label: '3.5.2 Validate Performance' },
        ],
      },
      {
        id: 'sec_3_6', title: '3.6\nHANDOVER', numRows: 1, steps: [
          { id: 'step_3_6_1', label: '3.6.1 Deliver Operations & Maintenance Training' },
          { id: 'step_3_6_2', label: '3.6.2 Finalise Operational Deliverables' },
          { id: 'step_3_6_3', label: '3.6.3 Approval', isApproval: true },
        ],
      },
    ],
  },
  {
    id: 'phase_4', title: '4  CLOSE', sections: [
      {
        id: 'sec_4_1', title: '4.1\nSETTLE', numRows: 1, steps: [
          { id: 'step_4_1_1', label: '4.1.1 Create Settlement Rule' },
        ],
      },
      {
        id: 'sec_4_2', title: '4.2\nFINALISE', numRows: 1, steps: [
          { id: 'step_4_2_1', label: '4.2.1 Perform Post Implementation Reviews' },
          { id: 'step_4_2_2', label: '4.2.2 Resolve Defects' },
          { id: 'step_4_2_3', label: '4.2.3 Close Project' },
          { id: 'step_4_2_4', label: '4.2.4 Approval', isApproval: true },
        ],
      },
    ],
  },
];
