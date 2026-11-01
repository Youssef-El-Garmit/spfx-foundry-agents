/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
/**
 * Right-hand document panel (like Claude's artifacts): SharePoint viewer of the document the agent is working on
 * (Word, Excel, PowerPoint, PDF), refreshed after each change, with "Open in Word/Excel/PowerPoint" and "Open in browser".
 * Falls back to the PDF conversion, then to the Office Online embed view.
 */
import * as React from 'react';
import {
  Button,
  Spinner,
  Tooltip,
  Tab,
  TabList,
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody
} from '@fluentui/react-components';
import {
  ArrowClockwise20Regular,
  Dismiss20Regular,
  Open20Regular,
  DocumentText20Regular,
  Window20Regular
} from '@fluentui/react-icons';
import type { IDocumentRef } from '../../services/mcpTypes';
import { DocumentService, type IDocumentLinks } from '../../services/DocumentService';
import strings from 'FoundryAgentsWebPartStrings';

const useStyles = makeStyles({
  panel: {
    display: 'flex',
    flexDirection: 'column',
    minWidth: 0,
    height: '100%',
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1
  },
  title: {
    flex: 1,
    minWidth: 0,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap'
  },
  tabs: { padding: `0 ${tokens.spacingHorizontalS}`, backgroundColor: tokens.colorNeutralBackground1, overflowX: 'auto' },
  body: { flex: 1, position: 'relative', minHeight: 0 },
  frame: { width: '100%', height: '100%', border: 'none', display: 'block', backgroundColor: tokens.colorNeutralBackground3 },
  center: {
    position: 'absolute',
    inset: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    alignItems: 'center',
    justifyContent: 'center',
    color: tokens.colorNeutralForeground3
  },
  bar: { margin: tokens.spacingVerticalS }
});

export interface IDocumentPanelProps {
  documents: IDocumentRef[];
  activeItemId?: string;
  service: DocumentService;
  onSelect: (itemId: string) => void;
  onClose: () => void;
}

export const DocumentPanel: React.FC<IDocumentPanelProps> = ({ documents, activeItemId, service, onSelect, onClose }) => {
  const s = useStyles();
  const doc = documents.find((d) => d.itemId === activeItemId) ?? documents[documents.length - 1];
  const [previewUrl, setPreviewUrl] = React.useState<string | undefined>(undefined);
  const [pdfUrl, setPdfUrl] = React.useState<string | undefined>(undefined);
  const [fallbackUrl, setFallbackUrl] = React.useState<string | undefined>(undefined);
  const [links, setLinks] = React.useState<IDocumentLinks>({});
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [reloadKey, setReloadKey] = React.useState(0);

  // Reload the preview when the document changes (agent edit → updatedAt) or on demand.
  React.useEffect(() => {
    if (!doc) return;
    const ctrl = new AbortController();
    let objectUrl: string | undefined;
    setLoading(true);
    setError(undefined);
    setFallbackUrl(undefined);
    service.getLinks(doc).then(setLinks).catch(() => setLinks({ webUrl: doc.webUrl }));
    // Small delay after an edit so SharePoint has the new version when converting.
    const wait = doc.updatedAt && Date.now() - doc.updatedAt < 4000 ? 1200 : 0;
    const t = window.setTimeout(() => {
      service
        .getPreviewUrl(doc, ctrl.signal)
        .then((u) => {
          setPdfUrl(undefined);
          setPreviewUrl(u);
        })
        .catch(() => {
          if (ctrl.signal.aborted) return;
          setPreviewUrl(undefined);
          return service.getPdfObjectUrl(doc, ctrl.signal).then((u) => {
            objectUrl = u;
            setPdfUrl(u);
          });
        })
        .catch((e: Error) => {
          if (ctrl.signal.aborted) return;
          setPdfUrl(undefined);
          if (doc.webUrl) setFallbackUrl(DocumentService.embedUrl(doc.webUrl));
          else setError(e.message);
        })
        .finally(() => !ctrl.signal.aborted && setLoading(false));
    }, wait);
    return () => {
      ctrl.abort();
      window.clearTimeout(t);
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [doc?.itemId, doc?.updatedAt, reloadKey, service]);

  if (!doc) return null;
  const app = DocumentService.desktopApp(doc.name);
  const openDesktop = links.fileUrl && app ? DocumentService.desktopUrl(links.fileUrl, app.scheme) : undefined;
  const openWeb = links.webUrl ?? doc.webUrl;

  return (
    <section className={s.panel} aria-label={strings.DocPanelAriaLabel.replace('{0}', doc.name)}>
      <div className={s.header}>
        <DocumentText20Regular />
        <span className={s.title} title={doc.name}>{doc.name}</span>
        <Tooltip content={strings.DocRefreshPreview} relationship="label">
          <Button appearance="subtle" icon={<ArrowClockwise20Regular />} onClick={() => setReloadKey((k) => k + 1)} aria-label={strings.DocRefreshPreview} />
        </Tooltip>
        {openDesktop && app && (
          <Button appearance="primary" size="small" icon={<Open20Regular />} as="a" href={openDesktop}>
            {strings.DocOpenInApp.replace('{0}', app.label)}
          </Button>
        )}
        {openWeb && (
          <Tooltip content={strings.DocOpenInBrowser} relationship="label">
            <Button appearance="subtle" icon={<Window20Regular />} as="a" href={openWeb} target="_blank" rel="noopener noreferrer" aria-label={strings.DocOpenInBrowser} />
          </Tooltip>
        )}
        <Tooltip content={strings.DocClosePreview} relationship="label">
          <Button appearance="subtle" icon={<Dismiss20Regular />} onClick={onClose} aria-label={strings.DocClosePreview} />
        </Tooltip>
      </div>
      {documents.length > 1 && (
        <TabList className={s.tabs} size="small" selectedValue={doc.itemId} onTabSelect={(_, d) => onSelect(String(d.value))}>
          {documents.map((d) => (
            <Tab key={d.itemId} value={d.itemId} title={d.name}>
              {d.name.length > 28 ? `${d.name.slice(0, 26)}…` : d.name}
            </Tab>
          ))}
        </TabList>
      )}
      <div className={s.body}>
        {error && (
          <MessageBar intent="error" className={s.bar}>
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}
        {previewUrl && <iframe key={previewUrl} className={s.frame} src={previewUrl} title={strings.DocPreviewTitle.replace('{0}', doc.name)} allow="clipboard-write" />}
        {!previewUrl && pdfUrl && <iframe key={pdfUrl} className={s.frame} src={`${pdfUrl}#view=FitH`} title={strings.DocPdfPreviewTitle.replace('{0}', doc.name)} />}
        {!previewUrl && !pdfUrl && fallbackUrl && <iframe className={s.frame} src={fallbackUrl} title={strings.DocPreviewTitle.replace('{0}', doc.name)} />}
        {loading && (
          <div className={s.center}>
            <Spinner size="medium" />
            <span>{strings.DocUpdatingPreview}</span>
          </div>
        )}
      </div>
    </section>
  );
};
