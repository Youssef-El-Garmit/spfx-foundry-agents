/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Tooltip, Button } from '@fluentui/react-components';
import { UserMessage as CopilotUserMessage } from '@fluentui-copilot/react-copilot-chat';
import {
  Document20Regular,
  DocumentPdf20Regular,
  DocumentCsv20Regular,
  DocumentText20Regular,
  DocumentJavascript20Regular,
  DocumentJava20Regular,
  DocumentCss20Regular,
  DocumentSass20Regular,
  DocumentYml20Regular,
  Image20Regular,
  Video20Regular,
  FolderZip20Regular,
  Archive20Regular,
  Code20Regular,
  Copy24Regular,
  CheckmarkCircle24Regular,
} from '@fluentui/react-icons';
import copy from 'copy-to-clipboard';
import { Markdown } from './Markdown';
import strings from 'FoundryAgentsWebPartStrings';
import styles from './UserMessage.module.scss';
import { IChatMessage } from '../ChatLayout';
import { stripContextPrefix } from '../../../utils/contextDisplay';

interface IUserAttachment {
  name: string;
  fileId: string;
}

/** Parses `[Attached file: "name", file_id: id]` markers out of history message content. */
function parseAttachmentsFromContent(content: string): { text: string; attachments: IUserAttachment[] } {
  const attachments: IUserAttachment[] = [];
  const pattern = /\[Attached file: "([^"]+)", file_id: ([^\]]+)\]/g;
  let match;
  while ((match = pattern.exec(content)) !== null) {
    attachments.push({ name: match[1], fileId: match[2] });
  }
  const text = content
    .replace(/\n*\[Attached file: "[^"]+", file_id: [^\]]+\]/g, '')
    .trim();
  return { text, attachments };
}

function getFileTypeIconByName(filename: string): React.ReactElement {
  const ext = (filename.split('.').pop() ?? '').toLowerCase();
  if (ext === 'pdf') return <DocumentPdf20Regular />;
  if (ext === 'csv') return <DocumentCsv20Regular />;
  if (['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg', 'ico'].includes(ext))
    return <Image20Regular />;
  if (['mp4', 'mov', 'avi', 'mkv', 'webm', 'flv'].includes(ext))
    return <Video20Regular />;
  if (['js', 'mjs', 'cjs', 'jsx', 'ts', 'tsx'].includes(ext))
    return <DocumentJavascript20Regular />;
  if (ext === 'java') return <DocumentJava20Regular />;
  if (ext === 'css') return <DocumentCss20Regular />;
  if (['scss', 'sass', 'less'].includes(ext)) return <DocumentSass20Regular />;
  if (['yaml', 'yml'].includes(ext)) return <DocumentYml20Regular />;
  if (['zip', 'gz', 'tar', 'bz2', '7z', 'rar', 'tgz'].includes(ext))
    return <FolderZip20Regular />;
  if (
    [
      'json', 'jsonl', 'xml', 'html', 'htm', 'sh', 'py', 'rb',
      'go', 'rs', 'swift', 'cpp', 'c', 'h', 'cs',
    ].includes(ext)
  )
    return <Code20Regular />;
  if (['zip', 'gz', 'tar', 'bz2', '7z', 'rar'].includes(ext))
    return <Archive20Regular />;
  if (['txt', 'md', 'log'].includes(ext)) return <DocumentText20Regular />;
  return <Document20Regular />;
}

export interface IUserMessageProps {
  message: IChatMessage;
}

export const UserMessage: React.FunctionComponent<IUserMessageProps> = ({ message }) => {
  const { text, attachments } = React.useMemo(() => {
    const forDisplay = stripContextPrefix(message.content);
    if (message.uploadedFiles !== undefined) {
      return { text: forDisplay, attachments: message.uploadedFiles };
    }
    return parseAttachmentsFromContent(forDisplay);
  }, [message.content, message.uploadedFiles]);

  const [copied, setCopied] = React.useState(false);
  const copyTimeoutRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  const handleCopyClick = React.useCallback(() => {
    copy(text);
    if (copyTimeoutRef.current) clearTimeout(copyTimeoutRef.current);
    setCopied(true);
    copyTimeoutRef.current = setTimeout(() => setCopied(false), 1500);
  }, [text]);

  React.useEffect(() => () => {
    if (copyTimeoutRef.current) clearTimeout(copyTimeoutRef.current);
  }, []);

  return (
    <div className={styles.userMessageGroup}>
      {attachments.length > 0 && (
        <div className={styles.attachmentsRow}>
          {attachments.map((a) => (
            <Tooltip key={a.fileId} content={a.name} relationship="label" withArrow>
              <div className={styles.attachmentChip}>
                <span className={styles.attachmentIcon}>
                  {getFileTypeIconByName(a.name)}
                </span>
                <span className={styles.attachmentName}>{a.name}</span>
              </div>
            </Tooltip>
          ))}
        </div>
      )}
      {text && (
        <div className={styles.messageWithActions}>
          <CopilotUserMessage>
            <Markdown content={text} />
          </CopilotUserMessage>
          <div className={styles.messageActions}>
            <Tooltip visible={copied} content={copied ? strings.Copied : ''} relationship="label" withArrow>
              <Button
                appearance="subtle"
                shape="circular"
                icon={copied ? <CheckmarkCircle24Regular /> : <Copy24Regular />}
                size="small"
                onClick={handleCopyClick}
                aria-label={copied ? strings.Copied : strings.CopyReply}
                className={copied ? styles.copyButtonCopied : styles.copyButton}
              />
            </Tooltip>
          </div>
        </div>
      )}
    </div>
  );
};
