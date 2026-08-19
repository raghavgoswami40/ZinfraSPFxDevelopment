import { StoredPhase } from './IVerticalFlowProps';

// Step code badges are generated from each step's position in its section
// (see buildCodeMap in VerticalFlow.tsx), so no codes are stored here.
// Row wrapping is also automatic — steps fill each row before wrapping.
export const DEFAULT_PHASES: StoredPhase[] = [
  {
    id: 'phase_2', title: '2  PLAN', sections: [
      {
        id: 'sec_2_0', title: '', codePrefix: '2.0', variant: 'grey', steps: [
          { id: 'step_2_0_1', label: 'Enable Early Works' },
          { id: 'step_2_0_ca', label: 'Client Approval', shape: 'approval' },
          { id: 'step_2_0_sowc', label: 'SOWC', shape: 'sapStatus' },
          { id: 'step_2_0_gate2', label: 'Gate 2', shape: 'gate' },
        ],
      },
      {
        id: 'sec_2_1', title: '2.1\nESTABLISH', steps: [
          { id: 'step_2_1_rffa', label: 'RFFA', shape: 'sapStatus' },
          { id: 'step_2_1_1', label: 'Receive Project' },
          { id: 'step_2_1_2', label: 'Assign Project Manager' },
          { id: 'step_2_1_3', label: 'Define Project' },
          { id: 'step_2_1_4', label: 'Assign Project Team' },
          { id: 'step_2_1_5', label: 'Develop Project Plans and Establish Systems' },
        ],
      },
      {
        id: 'sec_2_2', title: '2.2\nASSESS', steps: [
          { id: 'step_2_2_1', label: 'Perform Desktop Investigations' },
          { id: 'step_2_2_2', label: 'Engage Site Investigations Vendor(s)' },
          { id: 'step_2_2_3', label: 'Execute Site Investigations' },
        ],
      },
      {
        id: 'sec_2_3', title: '2.3\nDESIGN', steps: [
          { id: 'step_2_3_1', label: 'Develop Design Plans' },
          { id: 'step_2_3_2', label: 'Engage Designer(s)' },
          { id: 'step_2_3_3', label: 'Execute Design' },
        ],
      },
      {
        id: 'sec_2_4', title: '2.4\nFORMULATE', steps: [
          { id: 'step_2_4_1', label: 'Develop Delivery Methodology' },
          { id: 'step_2_4_2', label: 'Secure Requirements' },
        ],
      },
      {
        id: 'sec_2_5', title: '2.5\nPROPOSE', steps: [
          { id: 'step_2_5_1', label: 'Obtain Pricing' },
          { id: 'step_2_5_2', label: 'Develop Proposal' },
          { id: 'step_2_5_faco', label: 'FACO', shape: 'sapStatus' },
          { id: 'step_2_5_ca', label: 'Client Approval', shape: 'approval' },
          { id: 'step_2_5_appr', label: 'APPR', shape: 'sapStatus' },
          {
            id: 'step_2_5_grp', label: 'Customer Initiated Projects Only', shape: 'group', children: [
              { id: 'step_2_5_grp_wait', label: 'WAIT', shape: 'sapStatus' },
              { id: 'step_2_5_grp_capp', label: 'CAPP', shape: 'sapStatus' },
            ],
          },
          { id: 'step_2_5_gate3', label: 'Gate 3', shape: 'gate' },
        ],
      },
    ],
  },
  {
    id: 'phase_3', title: '3  DELIVER', sections: [
      {
        id: 'sec_3_1', title: '3.1\nPROCURE', steps: [
          { id: 'step_3_1_wipr', label: 'WIPR', shape: 'sapStatus' },
          { id: 'step_3_1_1', label: 'Engage Supplier(s) and Constructor(s)' },
          { id: 'step_3_1_2', label: 'Procure Long Lead Item Material and Equipment' },
        ],
      },
      {
        id: 'sec_3_2', title: '3.2\nPREPARE', steps: [
          { id: 'step_3_2_1', label: 'Finalise Planning' },
          { id: 'step_3_2_2', label: 'Perform Pre Mobilisation Reviews' },
          { id: 'step_3_2_3', label: 'Develop Construction and Commissioning Plans' },
          { id: 'step_3_2_4', label: 'Onboard Workgroup' },
          { id: 'step_3_2_gate4', label: 'Gate 4', shape: 'gate' },
        ],
      },
      {
        id: 'sec_3_3', title: '3.3\nMOBILISE', steps: [
          { id: 'step_3_3_1', label: 'Set-up Site' },
          { id: 'step_3_3_2', label: 'Deploy Resources' },
        ],
      },
      {
        id: 'sec_3_4', title: '3.4\nCONSTRUCT', steps: [
          { id: 'step_3_4_1', label: 'Execute Construction' },
          { id: 'step_3_4_2', label: 'Perform Completions' },
          { id: 'step_3_4_ca', label: 'Client Approval', shape: 'approval' },
          { id: 'step_3_4_gate5', label: 'Gate 5', shape: 'gate' },
        ],
      },
      {
        id: 'sec_3_5', title: '3.5\nCOMMISSION', steps: [
          { id: 'step_3_5_1', label: 'Execute Commissioning' },
          { id: 'step_3_5_2', label: 'Validate Performance' },
          { id: 'step_3_5_ca', label: 'Client Approval', shape: 'approval' },
          { id: 'step_3_5_pc', label: 'Practical Completion', shape: 'sapStatus', noPrefix: true },
          { id: 'step_3_5_wcom', label: 'WCOM', shape: 'sapStatus' },
        ],
      },
      {
        id: 'sec_3_6', title: '3.6\nHANDOVER', steps: [
          { id: 'step_3_6_1', label: 'Enable Operational Readiness' },
          { id: 'step_3_6_2', label: 'Finalise Operational Deliverables' },
          { id: 'step_3_6_hold', label: 'HOLD', shape: 'sapStatus' },
          { id: 'step_3_6_ca', label: 'Client Approval', shape: 'approval' },
          { id: 'step_3_6_gate6', label: 'Gate 6', shape: 'gate' },
        ],
      },
    ],
  },
  {
    id: 'phase_4', title: '4  CLOSE', sections: [
      {
        id: 'sec_4_1', title: '4.1\nSETTLE', steps: [
          { id: 'step_4_1_teco', label: 'TECO', shape: 'sapStatus' },
          { id: 'step_4_1_1', label: 'Resolve Defects' },
          { id: 'step_4_1_2', label: 'Perform Post Implementation Reviews' },
          { id: 'step_4_1_3', label: 'Create Settlement Rule' },
        ],
      },
      {
        id: 'sec_4_2', title: '4.2\nFINALISE', steps: [
          { id: 'step_4_2_1', label: 'Capitalise Cost' },
          { id: 'step_4_2_2', label: 'Close Project' },
          { id: 'step_4_2_ca', label: 'Client Approval', shape: 'approval' },
          { id: 'step_4_2_clsd', label: 'CLSD', shape: 'sapStatus' },
          { id: 'step_4_2_gate7', label: 'Gate 7', shape: 'gate' },
        ],
      },
    ],
  },
  {
    id: 'phase_5', title: '5  CONTROL', sections: [
      {
        id: 'sec_5_0', title: '', codePrefix: '5.0', steps: [
          { id: 'step_5_0_1', label: 'Manage Project Controls' },
          { id: 'step_5_0_2', label: 'Monitor Performance' },
          { id: 'step_5_0_3', label: 'Govern Project' },
          { id: 'step_5_0_4', label: 'Manage Change' },
          { id: 'step_5_0_5', label: 'Manage Stakeholders' },
          { id: 'step_5_0_6', label: 'Manage Contacts' },
        ],
      },
    ],
  },
];
