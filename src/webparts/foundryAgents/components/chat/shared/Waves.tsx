import * as React from 'react';
import styles from './Waves.module.scss';

interface Wave {
  amplitude: number;
  frequency: number;
  speed: number;
  baseHeight: number;
  offset: number;
}

export interface IWavesProps {
  /** When set (e.g. in Teams), waves use this brand color instead of reading from DOM. Ensures correct theme before CSS vars are applied. */
  brandColor?: string;
}

/** Parse theme color (rgb/rgba/hex) and return rgba string with given alpha */
function withAlpha(cssColor: string, alpha: number): string {
  const s = cssColor.trim();
  if (!s) return `rgba(0, 120, 212, ${alpha})`;
  const rgb = s.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
  if (rgb) return `rgba(${rgb[1]},${rgb[2]},${rgb[3]},${alpha})`;
  const hex = s.replace(/^#/, '');
  if (hex.length === 6) {
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    return `rgba(${r},${g},${b},${alpha})`;
  }
  return `rgba(0, 120, 212, ${alpha})`;
}

function getThemeWaveColors(el: Element | undefined, brandColorOverride?: string): [string, string, string] {
  const fallback: [string, string, string] = ['rgba(0, 120, 212, 0.15)', 'rgba(0, 120, 212, 0.12)', 'rgba(0, 120, 212, 0.1)'];
  if (brandColorOverride) {
    return [withAlpha(brandColorOverride, 0.15), withAlpha(brandColorOverride, 0.12), withAlpha(brandColorOverride, 0.1)];
  }
  if (!el) return fallback;
  const s = getComputedStyle(el);
  const brand = s.getPropertyValue('--colorBrandBackground1').trim()
    || s.getPropertyValue('--colorBrandBackground').trim()
    || s.getPropertyValue('--colorBrandStroke1').trim()
    || '';
  return brand ? [withAlpha(brand, 0.15), withAlpha(brand, 0.12), withAlpha(brand, 0.1)] as [string, string, string] : fallback;
}

export const Waves: React.FunctionComponent<IWavesProps> = ({ brandColor }) => {
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const animationFrameRef = React.useRef<number | undefined>(undefined);

  React.useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const resizeCanvas = (): void => {
      canvas.width = canvas.offsetWidth;
      canvas.height = canvas.offsetHeight;
    };
    resizeCanvas();
    window.addEventListener('resize', resizeCanvas);

    const waves: Wave[] = [
      { amplitude: 20, frequency: 0.004, speed: 0.02, baseHeight: 0.9, offset: 0 },
      { amplitude: 15, frequency: 0.006, speed: 0.015, baseHeight: 0.85, offset: 0 },
      { amplitude: 10, frequency: 0.008, speed: 0.01, baseHeight: 0.8, offset: 0 }
    ];

    const colors = getThemeWaveColors(canvas.parentElement ?? undefined, brandColor);

    let time = 0;

    const drawWave = (wave: Wave, color: string): void => {
      ctx.beginPath();
      ctx.moveTo(0, canvas.height);
      for (let x = 0; x <= canvas.width; x++) {
        const y =
          wave.baseHeight * canvas.height +
          wave.amplitude * Math.sin(x * wave.frequency + wave.offset);
        ctx.lineTo(x, y);
      }
      ctx.lineTo(canvas.width, canvas.height);
      ctx.lineTo(0, canvas.height);
      ctx.closePath();
      ctx.fillStyle = color;
      ctx.fill();
    };

    const animate = (): void => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      waves.forEach((wave, index) => {
        wave.offset = time * wave.speed;
        drawWave(wave, colors[index]);
      });
      time += 1;
      animationFrameRef.current = requestAnimationFrame(animate);
    };

    const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (!prefersReducedMotion) {
      animate();
    } else {
      waves.forEach((wave, index) => drawWave(wave, colors[index]));
    }

    return () => {
      window.removeEventListener('resize', resizeCanvas);
      if (animationFrameRef.current) {
        cancelAnimationFrame(animationFrameRef.current);
      }
    };
  }, [brandColor]);

  return <canvas ref={canvasRef} className={styles.wavesCanvas} aria-hidden="true" />;
};
