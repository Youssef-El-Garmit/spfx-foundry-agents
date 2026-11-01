/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../FoundryAgentsWebPartStrings.d.ts" />
/**
 * Documents shown next to the chat: SharePoint viewer (Graph driveItem preview: Word, Excel, PowerPoint, PDF…),
 * PDF conversion as fallback, direct file URL for "Open in the desktop app" (ms-word:/ms-excel:/ms-powerpoint:).
 * Uses the user's Graph token (SPFx).
 */
import type { GetTokenFn } from './SpFxTokenCredential';
import type { IDocumentRef } from './mcpTypes';
import strings from 'FoundryAgentsWebPartStrings';

const GRAPH = 'https://graph.microsoft.com';

export interface IDocumentLinks {
  /** Direct file URL (…/Shared Documents/…/file.docx). */
  fileUrl?: string;
  /** Word Online URL (Doc.aspx). */
  webUrl?: string;
}

export class DocumentService {
  constructor(private readonly getToken: GetTokenFn) {}

  private async graph(path: string, init?: RequestInit): Promise<Response> {
    const token = await this.getToken(GRAPH);
    return fetch(`${GRAPH}/v1.0${path}`, { ...init, headers: { ...(init?.headers ?? {}), Authorization: `Bearer ${token}` } });
  }

  /**
   * SharePoint viewer URL for the file (embed.aspx with a short-lived access token: no sign-in inside the frame).
   * Works for every Office format and PDF, straight from SharePoint / OneDrive.
   */
  async getPreviewUrl(doc: IDocumentRef, signal?: AbortSignal): Promise<string> {
    const res = await this.graph(`/drives/${encodeURIComponent(doc.driveId)}/items/${encodeURIComponent(doc.itemId)}/preview`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: '{}',
      signal
    });
    if (!res.ok) throw new Error(strings.DocPreviewUnavailableStatus.replace('{0}', String(res.status)));
    const body = (await res.json()) as { getUrl?: string };
    if (!body.getUrl) throw new Error(strings.DocPreviewUnavailable);
    return body.getUrl;
  }

  /** PDF rendering of the document (Graph converts .docx server-side). Returns an object URL to revoke after use. */
  async getPdfObjectUrl(doc: IDocumentRef, signal?: AbortSignal): Promise<string> {
    const res = await this.graph(`/drives/${encodeURIComponent(doc.driveId)}/items/${encodeURIComponent(doc.itemId)}/content?format=pdf`, { signal });
    if (!res.ok) throw new Error(strings.DocPdfPreviewUnavailable.replace('{0}', String(res.status)));
    const blob = await res.blob();
    return URL.createObjectURL(new Blob([blob], { type: 'application/pdf' }));
  }

  /** File URL and Word Online URL of the document. */
  async getLinks(doc: IDocumentRef): Promise<IDocumentLinks> {
    const res = await this.graph(
      `/drives/${encodeURIComponent(doc.driveId)}/items/${encodeURIComponent(doc.itemId)}?$select=webUrl,name&$expand=listItem($select=webUrl)`
    );
    if (!res.ok) return { webUrl: doc.webUrl };
    const body = (await res.json()) as { webUrl?: string; listItem?: { webUrl?: string } };
    return { fileUrl: body.listItem?.webUrl, webUrl: body.webUrl ?? doc.webUrl };
  }

  /** Desktop app for a file name: Word, Excel or PowerPoint (undefined for other formats). */
  static desktopApp(fileName: string): { scheme: string; label: string } | undefined {
    const ext = fileName.split('.').pop()?.toLowerCase() ?? '';
    if (['docx', 'doc', 'docm', 'dotx', 'rtf'].includes(ext)) return { scheme: 'ms-word', label: 'Word' };
    if (['xlsx', 'xls', 'xlsm', 'csv'].includes(ext)) return { scheme: 'ms-excel', label: 'Excel' };
    if (['pptx', 'ppt', 'pptm'].includes(ext)) return { scheme: 'ms-powerpoint', label: 'PowerPoint' };
    return undefined;
  }

  /** "Open in the desktop app" link (ms-word:, ms-excel:, ms-powerpoint:). */
  static desktopUrl(fileUrl: string, scheme = 'ms-word'): string {
    // Each path segment encoded exactly once (a name with « & » used to become %2526: file not found in Word).
    const decode = (seg: string): string => { try { return decodeURIComponent(seg); } catch { return seg; } };
    try {
      const u = new URL(fileUrl);
      return `${scheme}:ofe|u|${u.origin}${u.pathname.split('/').map((seg) => encodeURIComponent(decode(seg))).join('/')}`;
    } catch {
      return `${scheme}:ofe|u|${fileUrl}`;
    }
  }

  /** Word Online embedded viewer URL (fallback when the PDF conversion is not available). */
  static embedUrl(webUrl: string): string {
    try {
      const u = new URL(webUrl);
      u.searchParams.set('action', 'embedview');
      u.searchParams.delete('mobileredirect');
      return u.toString();
    } catch {
      return webUrl;
    }
  }
}
