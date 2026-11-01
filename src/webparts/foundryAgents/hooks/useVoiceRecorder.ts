/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../FoundryAgentsWebPartStrings.d.ts" />
/**
 * Voice input like ChatGPT / Claude: record with the microphone, live audio levels (for the waveform)
 * and elapsed time, then transcribe with the agent's transcription endpoint when one is configured.
 * The browser speech recognition runs silently in parallel and is used when there is no transcription endpoint
 * or when the server transcription fails (quota, network), so the dictation is never lost.
 */
import * as React from 'react';
import strings from 'FoundryAgentsWebPartStrings';

interface ISpeechRecognitionResult { isFinal: boolean; 0: { transcript: string } }
interface ISpeechRecognitionEvent { resultIndex: number; results: ArrayLike<ISpeechRecognitionResult> }
interface ISpeechRecognition {
  lang: string;
  continuous: boolean;
  interimResults: boolean;
  start(): void;
  stop(): void;
  abort(): void;
  onresult: ((e: ISpeechRecognitionEvent) => void) | undefined;
  onerror: ((e: { error: string }) => void) | undefined;
  onend: (() => void) | undefined;
}

function recognitionCtor(): (new () => ISpeechRecognition) | undefined {
  const w = window as unknown as { SpeechRecognition?: new () => ISpeechRecognition; webkitSpeechRecognition?: new () => ISpeechRecognition };
  return w.SpeechRecognition ?? w.webkitSpeechRecognition;
}

export type VoiceState = 'idle' | 'recording' | 'transcribing';

export interface IVoiceRecorder {
  supported: boolean;
  state: VoiceState;
  /** Elapsed recording time in ms. */
  elapsed: number;
  /** Recent audio levels 0..1, oldest first (for the waveform). */
  levels: number[];
  error?: string;
  start: () => void;
  /** Stop and transcribe; resolves with the text ('' when nothing was understood). */
  stop: () => Promise<string>;
  cancel: () => void;
}

const MAX_LEVELS = 64;
const MAX_MS = 5 * 60_000;

export function useVoiceRecorder(
  transcribe: ((audio: Blob, lang: string) => Promise<string>) | undefined,
  lang: string
): IVoiceRecorder {
  const supported = React.useMemo(
    () => typeof window !== 'undefined' && !!navigator.mediaDevices?.getUserMedia && typeof (window as unknown as { MediaRecorder?: unknown }).MediaRecorder !== 'undefined',
    []
  );
  const [state, setState] = React.useState<VoiceState>('idle');
  const [elapsed, setElapsed] = React.useState(0);
  const [levels, setLevels] = React.useState<number[]>([]);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const rt = React.useRef<{
    stream?: MediaStream;
    recorder?: MediaRecorder;
    chunks: Blob[];
    ctx?: AudioContext;
    raf?: number;
    startedAt: number;
    rec?: ISpeechRecognition;
    backup: string;
    levels: number[];
    cancelled: boolean;
    stopped?: Promise<Blob>;
  }>({ chunks: [], startedAt: 0, backup: '', levels: [], cancelled: false });

  const cleanup = React.useCallback(() => {
    const r = rt.current;
    if (r.raf) cancelAnimationFrame(r.raf);
    r.raf = undefined;
    r.stream?.getTracks().forEach((t) => t.stop());
    r.stream = undefined;
    r.ctx?.close().catch(() => undefined);
    r.ctx = undefined;
    try { r.rec?.abort(); } catch { /* already stopped */ }
    r.rec = undefined;
  }, []);

  React.useEffect(() => () => { rt.current.cancelled = true; cleanup(); }, [cleanup]);

  const start = React.useCallback(() => {
    if (!supported || state !== 'idle') return;
    setError(undefined);
    const r = rt.current;
    r.chunks = [];
    r.backup = '';
    r.levels = [];
    r.cancelled = false;
    setLevels([]);
    setElapsed(0);
    navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true } }).then((stream) => {
      r.stream = stream;
      const recorder = new MediaRecorder(stream);
      r.recorder = recorder;
      r.stopped = new Promise<Blob>((resolve) => {
        recorder.onstop = () => resolve(new Blob(r.chunks, { type: recorder.mimeType || 'audio/webm' }));
      });
      recorder.ondataavailable = (ev) => { if (ev.data.size) r.chunks.push(ev.data); };
      recorder.start(500);
      r.startedAt = Date.now();

      // Audio levels for the waveform
      const Ctx = (window as unknown as { AudioContext?: typeof AudioContext; webkitAudioContext?: typeof AudioContext }).AudioContext
        ?? (window as unknown as { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
      if (Ctx) {
        const ctx = new Ctx();
        r.ctx = ctx;
        const analyser = ctx.createAnalyser();
        analyser.fftSize = 1024;
        ctx.createMediaStreamSource(stream).connect(analyser);
        const buf = new Uint8Array(analyser.fftSize);
        let last = 0;
        const tick = (t: number): void => {
          if (t - last > 70) {
            last = t;
            analyser.getByteTimeDomainData(buf);
            let sum = 0;
            for (let i = 0; i < buf.length; i++) { const v = (buf[i] - 128) / 128; sum += v * v; }
            const rms = Math.sqrt(sum / buf.length);
            const level = Math.min(1, rms * 4.5);
            r.levels = [...r.levels, level].slice(-MAX_LEVELS);
            setLevels(r.levels);
            const ms = Date.now() - r.startedAt;
            setElapsed(ms);
            if (ms > MAX_MS && r.recorder?.state === 'recording') r.recorder.stop();
          }
          r.raf = requestAnimationFrame(tick);
        };
        r.raf = requestAnimationFrame(tick);
      }

      // Silent backup transcription by the browser (used only if the server transcription fails)
      const RC = recognitionCtor();
      if (RC) {
        try {
          const rec = new RC();
          rec.lang = lang;
          rec.continuous = true;
          rec.interimResults = false;
          rec.onresult = (e) => {
            for (let i = e.resultIndex; i < e.results.length; i++) {
              if (e.results[i].isFinal) r.backup = `${r.backup} ${e.results[i][0].transcript}`.trim();
            }
          };
          rec.onerror = () => undefined;
          rec.start();
          r.rec = rec;
        } catch { /* no backup */ }
      }
      setState('recording');
    }).catch((e: Error) => {
      setError(e.name === 'NotAllowedError' ? strings.VoiceMicDenied : strings.VoiceMicUnavailable.replace('{0}', e.message));
      cleanup();
      setState('idle');
    });
  }, [supported, state, lang, cleanup]);

  const stop = React.useCallback(async (): Promise<string> => {
    const r = rt.current;
    if (!r.recorder || state !== 'recording') return '';
    setState('transcribing');
    try { r.rec?.stop(); } catch { /* ignore */ }
    if (r.recorder.state === 'recording') r.recorder.stop();
    const blob = await (r.stopped ?? Promise.resolve(new Blob()));
    // let the browser recognition deliver its last result
    await new Promise((resolve) => setTimeout(resolve, 350));
    const backup = r.backup;
    cleanup();
    if (r.cancelled) { setState('idle'); return ''; }
    let text = '';
    if (transcribe && blob.size > 1500) {
      try {
        text = (await transcribe(blob, lang)).trim();
      } catch (e) {
        if (!backup) setError(/429/.test(String((e as Error).message)) ? strings.VoiceQuotaReached : strings.VoiceTranscriptionFailed);
      }
    }
    setState('idle');
    setElapsed(0);
    return text || backup;
  }, [state, transcribe, lang, cleanup]);

  const cancel = React.useCallback(() => {
    const r = rt.current;
    r.cancelled = true;
    try { if (r.recorder?.state === 'recording') r.recorder.stop(); } catch { /* ignore */ }
    cleanup();
    setState('idle');
    setElapsed(0);
    setLevels([]);
  }, [cleanup]);

  return { supported, state, elapsed, levels, error, start, stop, cancel };
}
