/* eslint-disable @rushstack/no-new-null --
 * The Process Panel props use null for "nothing selected" by contract. */
import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import ProcessPanel from '../../processPanel/components/ProcessPanel';
import {
  IProcessStep, IRasciRecord, EmptyReason,
} from '../../processPanel/components/IProcessPanelProps';
import { resolveColors } from '../../processPanel/components/roleColors';
import { ProcessDetailsService } from '../../processPanel/services/ProcessDetailsService';
import { RasciService } from '../../processPanel/services/RasciService';

export interface IPanelConfig {
  processListTitle: string;
  systemListTitle: string;
  rasciListTitle: string;
  pdTitle: string;
  showBreadcrumb: boolean;
  pdEmptyIcon: string;
  pdEmptyHeading: string;
  pdEmptyBody: string;
  rasciTitle: string;
  rasciEmptyIcon: string;
  rasciEmptyIconL4: string;
  rasciEmptyHeadingL3: string;
  rasciEmptyBodyL3: string;
  rasciEmptyHeadingL4: string;
  rasciEmptyBodyL4: string;
  colorResponsible: string;
  colorAccountable: string;
  colorSupports: string;
  colorConsulted: string;
  colorInformed: string;
}

export interface IProcessPanelContainerProps {
  spHttpClient: SPHttpClient;
  webUrl: string;
  /** Generated code of the selected L3 step, e.g. "2.1.3". */
  code: string | undefined;
  config: IPanelConfig;
}

interface IPdState {
  step: IProcessStep | null;
  loading: boolean;
  error?: string;
  notFound?: string;
}

interface IRasciState {
  record: IRasciRecord | null;
  loading: boolean;
  error?: string;
  notFound?: string;
}

interface IRasciTarget {
  code?: string;
  title?: string;
  reason: EmptyReason;
  waiting: boolean;
}

const msg = (err: unknown): string => (err instanceof Error && err.message ? err.message : String(err));

/**
 * The Process Panel's fetch cycle as a component, so it can live inside the
 * hierarchy web part. Behaviour is deliberately identical to ProcessPanelWebPart:
 * Process Details loads for the selected L3; RASCI loads for the L3 itself
 * unless it has L4 sub-steps, in which case it waits for one to be opened.
 */
const ProcessPanelContainer: React.FC<IProcessPanelContainerProps> = ({ spHttpClient, webUrl, code, config }) => {
  const pdService = React.useMemo(
    () => new ProcessDetailsService(spHttpClient, webUrl, config.processListTitle, config.systemListTitle),
    [spHttpClient, webUrl, config.processListTitle, config.systemListTitle]
  );
  const rasciService = React.useMemo(
    () => new RasciService(spHttpClient, webUrl, config.rasciListTitle),
    [spHttpClient, webUrl, config.rasciListTitle]
  );

  const [pd, setPd] = React.useState<IPdState>({ step: null, loading: false });
  const [openSub, setOpenSub] = React.useState<string | null>(null);
  const [pdTry, setPdTry] = React.useState(0);

  React.useEffect(() => {
    setOpenSub(null);
    if (!code) { setPd({ step: null, loading: false }); return undefined; }
    let live = true;
    setPd({ step: null, loading: true });
    pdService.loadProcess(code)
      .then(step => { if (live) setPd({ step, loading: false, notFound: step ? undefined : code }); })
      .catch(err => { if (live) setPd({ step: null, loading: false, error: msg(err) }); });
    return () => { live = false; };
  }, [code, pdService, pdTry]);

  // The L3/L4 rule, unchanged from ProcessPanelWebPart._resolveRasciTarget.
  const target = React.useMemo((): IRasciTarget => {
    if (!code) return { reason: 'noSelection', waiting: false };
    if (pd.loading) return { reason: 'noSelection', waiting: true };
    const step = pd.step;
    if (step && step.subSteps.length > 0) {
      const open = openSub ? step.subSteps.filter(s => s.code === openSub)[0] : undefined;
      if (open) return { code: open.code, title: open.title, reason: 'notFound', waiting: false };
      return { reason: 'awaitingL4', waiting: false };
    }
    if (step) return { code: step.code, title: step.title, reason: 'notFound', waiting: false };
    return { code, reason: 'notFound', waiting: false };
  }, [code, pd, openSub]);

  const [rs, setRs] = React.useState<IRasciState>({ record: null, loading: false });
  const [rsTry, setRsTry] = React.useState(0);

  React.useEffect(() => {
    const c = target.code;
    if (!c) { setRs({ record: null, loading: false }); return undefined; }
    let live = true;
    setRs({ record: null, loading: true });
    rasciService.loadRasci(c)
      .then(record => { if (live) setRs({ record, loading: false, notFound: record ? undefined : c }); })
      .catch(err => { if (live) setRs({ record: null, loading: false, error: msg(err) }); });
    return () => { live = false; };
  }, [target.code, rasciService, rsTry]);

  const record: IRasciRecord | null = rs.record
    ? { ...rs.record, title: target.title || rs.record.title }
    : null;

  return (
    <ProcessPanel
      compact
      processDetails={{
        selectedStep: pd.step,
        isLoading: pd.loading,
        error: pd.error,
        notFoundCode: pd.notFound,
        onRetry: () => setPdTry(n => n + 1),
        openSubStepCode: openSub,
        onSubStepToggle: setOpenSub,
        title: config.pdTitle,
        showBreadcrumb: config.showBreadcrumb,
        singleOpenSubStep: true,
        emptyIcon: config.pdEmptyIcon,
        emptyHeading: config.pdEmptyHeading,
        emptyBody: config.pdEmptyBody,
      }}
      rasci={{
        record,
        isLoading: rs.loading || target.waiting,
        error: rs.error,
        emptyReason: rs.notFound ? 'notFound' : target.reason,
        notFoundCode: rs.notFound,
        onRetry: () => setRsTry(n => n + 1),
        title: config.rasciTitle,
        colors: resolveColors({
          responsible: config.colorResponsible,
          accountable: config.colorAccountable,
          supports: config.colorSupports,
          consulted: config.colorConsulted,
          informed: config.colorInformed,
        }),
        emptyIcon: config.rasciEmptyIcon,
        emptyIconL4: config.rasciEmptyIconL4,
        emptyHeadingL3: config.rasciEmptyHeadingL3,
        emptyBodyL3: config.rasciEmptyBodyL3,
        emptyHeadingL4: config.rasciEmptyHeadingL4,
        emptyBodyL4: config.rasciEmptyBodyL4,
      }}
    />
  );
};

export default ProcessPanelContainer;
