/**
 * Recording bar shown in place of the chat input (ChatGPT / Claude style): cancel, red dot + timer,
 * live waveform of the microphone level, confirm. While the audio is transcribed, a shimmer replaces the bars.
 */
/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Button, Spinner, Tooltip, makeStyles, shorthands, tokens } from '@fluentui/react-components';
import { Dismiss20Regular, Checkmark20Regular } from '@fluentui/react-icons';
import type { VoiceState } from '../../../hooks/useVoiceRecorder';
import strings from 'FoundryAgentsWebPartStrings';

const pulse = { '0%': { opacity: 1 }, '50%': { opacity: 0.35 }, '100%': { opacity: 1 } };

const useStyles = makeStyles({
  bar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minHeight: '56px',
    ...shorthands.padding('6px', '8px'),
    ...shorthands.borderRadius('8px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4
  },
  dot: {
    width: '10px',
    height: '10px',
    ...shorthands.borderRadius('50%'),
    backgroundColor: tokens.colorPaletteRedBackground3,
    animationName: pulse,
    animationDuration: '1.2s',
    animationIterationCount: 'infinite',
    flexShrink: 0
  },
  time: { fontVariantNumeric: 'tabular-nums', fontSize: tokens.fontSizeBase300, color: tokens.colorNeutralForeground2, minWidth: '40px' },
  wave: { flex: 1, height: '36px', display: 'block', minWidth: 0 },
  status: { flex: 1, display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }
});

function fmt(ms: number): string {
  const s = Math.floor(ms / 1000);
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
}

export interface IVoiceRecorderBarProps {
  state: VoiceState;
  elapsed: number;
  levels: number[];
  onCancel: () => void;
  onConfirm: () => void;
}

export const VoiceRecorderBar: React.FC<IVoiceRecorderBarProps> = ({ state, elapsed, levels, onCancel, onConfirm }) => {
  const s = useStyles();
  const canvasRef = React.useRef<HTMLCanvasElement>(null);

  // Draw the waveform: one rounded bar per sampled level, newest on the right.
  React.useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    const w = canvas.clientWidth, h = canvas.clientHeight;
    if (canvas.width !== w * dpr || canvas.height !== h * dpr) { canvas.width = w * dpr; canvas.height = h * dpr; }
    const g = canvas.getContext('2d');
    if (!g) return;
    g.setTransform(dpr, 0, 0, dpr, 0, 0);
    g.clearRect(0, 0, w, h);
    const color = getComputedStyle(canvas).color || '#0f6cbd';
    g.fillStyle = color;
    const barW = 3, gap = 3, step = barW + gap;
    const count = Math.floor(w / step);
    const data = levels.slice(-count);
    const offset = w - data.length * step;
    data.forEach((lv, i) => {
      const bh = Math.max(3, lv * (h - 4));
      const x = offset + i * step;
      const y = (h - bh) / 2;
      g.beginPath();
      const r = Math.min(1.5, bh / 2);
      g.moveTo(x + r, y);
      g.arcTo(x + barW, y, x + barW, y + bh, r);
      g.arcTo(x + barW, y + bh, x, y + bh, r);
      g.arcTo(x, y + bh, x, y, r);
      g.arcTo(x, y, x + barW, y, r);
      g.fill();
    });
  }, [levels]);

  const recording = state === 'recording';
  return (
    <div className={s.bar} role="group" aria-label={recording ? strings.VoiceRecording : strings.VoiceTranscribing}>
      <Tooltip content={strings.Cancel} relationship="label">
        <Button appearance="subtle" shape="circular" icon={<Dismiss20Regular />} onClick={onCancel} aria-label={strings.VoiceCancelRecording} disabled={!recording} />
      </Tooltip>
      {recording ? (
        <>
          <span className={s.dot} aria-hidden />
          <span className={s.time} aria-live="off">{fmt(elapsed)}</span>
          <canvas ref={canvasRef} className={s.wave} style={{ color: tokens.colorBrandForeground1 }} aria-hidden />
        </>
      ) : (
        <span className={s.status}><Spinner size="extra-tiny" /> {strings.VoiceTranscribingShort}</span>
      )}
      <Tooltip content={strings.VoiceFinishAndTranscribe} relationship="label">
        <Button appearance="primary" shape="circular" icon={<Checkmark20Regular />} onClick={onConfirm} aria-label={strings.VoiceFinishRecording} disabled={!recording} />
      </Tooltip>
    </div>
  );
};
