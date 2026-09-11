import { IProcessStep } from './IProcessDetailsProps';

/**
 * Design-time sample only. Lets the panel be reviewed in its populated states
 * without a SharePoint connection; switched on by the "Show sample data" toggle
 * in the property pane.
 *
 * The logos below are stand-ins so the icon states are visible offline. Real
 * logos always come from the System Master list — nothing here is used on the
 * live path, so an unpopulated System Master still shows as unpopulated.
 */

const logo = (markup: string): string =>
  `data:image/svg+xml;charset=utf-8,${encodeURIComponent(markup)}`;

const WORD = logo(
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
     <rect width="32" height="32" rx="4" fill="#2B579A"/>
     <path d="M6 10h3.1l2.1 8.6L13.6 10h2.9l2.4 8.6L21 10h3.1l-3.7 12.4h-3.1L15 14.3l-2.3 8.1H9.6z" fill="#fff"/>
   </svg>`
);

const EXCEL = logo(
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
     <rect width="32" height="32" rx="4" fill="#217346"/>
     <path d="M9 9.5h3.7l3.3 4.9 3.3-4.9H23l-5 6.5 5 6.5h-3.7L16 17.6l-3.3 4.9H9l5-6.5z" fill="#fff"/>
   </svg>`
);

const SAP = logo(
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 32">
     <defs>
       <linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
         <stop offset="0" stop-color="#00AEEF"/>
         <stop offset="1" stop-color="#0057A8"/>
       </linearGradient>
     </defs>
     <path d="M0 4h64L52 28H0z" fill="url(#g)"/>
     <text x="26" y="21.5" font-family="Segoe UI, Arial, sans-serif" font-size="13"
           font-weight="700" fill="#fff" text-anchor="middle">SAP</text>
   </svg>`
);

export const SAMPLE_STEP_WITH_SUBSTEPS: IProcessStep = {
  code: '2.1.3',
  title: 'Define Project',
  breadcrumb: 'PLAN › Establish',
  purpose:
    'Establish the agreed scope, objectives and delivery approach for the project so that ' +
    'planning, estimating and contracting activities all work from a single baseline.',
  inputs: [
    'Approved business case',
    'Client brief and scope of works',
    'Feasibility study outcomes',
  ],
  outputs: [],
  subSteps: [
    {
      code: '2.1.3.1',
      title: 'Confirm scope and objectives with the client',
      outputs: [
        { text: 'Project scope statement', system: 'Word', logoUrl: WORD, href: 'https://example.com/scope.docx' },
        { text: 'Objectives register', system: 'Excel', logoUrl: EXCEL, href: 'https://example.com/objectives.xlsx' },
      ],
    },
    {
      code: '2.1.3.2',
      title: 'Define deliverables and acceptance criteria',
      outputs: [
        { text: 'Deliverables breakdown', system: 'Excel', logoUrl: EXCEL, href: 'https://example.com/wbs.xlsx' },
        { text: 'Project master record created', system: 'SAP', logoUrl: SAP, href: 'https://example.com/sap/CJ20N' },
        // Tagged with a system but no URL recorded — renders dimmed.
        { text: 'Acceptance criteria signed off', system: 'Word', logoUrl: WORD },
        // No system tag at all — renders with no icon.
        { text: 'Verbal agreement noted in minutes', system: '' },
      ],
    },
    {
      code: '2.1.3.3',
      title: 'Establish the delivery approach',
      outputs: [
        { text: 'Delivery strategy note', system: 'Word', logoUrl: WORD, href: 'https://example.com/strategy.docx' },
      ],
    },
  ],
};

export const SAMPLE_STEP_FLAT: IProcessStep = {
  code: '2.1.4',
  title: 'Mobilise Delivery Team',
  breadcrumb: 'PLAN › Establish',
  purpose:
    'Stand up the delivery team with clear accountabilities so work can start against the ' +
    'agreed baseline without further hand-offs.',
  inputs: ['Approved project definition', 'Resource availability forecast'],
  outputs: [
    { text: 'Team mobilisation plan', system: 'Word', logoUrl: WORD, href: 'https://example.com/mobilisation.docx' },
    { text: 'Resource assignments posted', system: 'SAP', logoUrl: SAP, href: 'https://example.com/sap/CAT2' },
    { text: 'Kick-off held', system: '' },
  ],
  subSteps: [],
};
