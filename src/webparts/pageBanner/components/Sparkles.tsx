import * as React from 'react';

interface ISparklesProps {
  className?: string;
  density?: number;
  color?: string;
}

interface IStar { x: number; y: number; r: number; speed: number; phase: number; rate: number }

/** Twinkling starfield drawn on a canvas that fills its parent. */
const Sparkles: React.FC<ISparklesProps> = ({ className, density = 120, color = '#ffffff' }) => {
  const canvasRef = React.useRef<HTMLCanvasElement>(null);

  React.useEffect(() => {
    const canvas = canvasRef.current;
    const ctx = canvas ? canvas.getContext('2d') : null;
    if (!canvas || !ctx) return undefined;

    const reduceMotion = !!window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    let width = 0;
    let height = 0;
    let stars: IStar[] = [];
    let frame = 0;

    const setup = (): void => {
      const rect = canvas.getBoundingClientRect();
      const dpr = window.devicePixelRatio || 1;
      width = rect.width;
      height = rect.height;
      canvas.width = width * dpr;
      canvas.height = height * dpr;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      const count = Math.max(20, Math.round(density * (width * height) / (400 * 110)));
      stars = [];
      for (let i = 0; i < count; i++) {
        stars.push({
          x: Math.random() * width,
          y: Math.random() * height,
          r: Math.random() * 0.8 + 0.4,
          speed: Math.random() * 0.25 + 0.05,
          phase: Math.random() * Math.PI * 2,
          rate: Math.random() * 2 + 0.8,
        });
      }
    };

    const draw = (time: number): void => {
      ctx.clearRect(0, 0, width, height);
      ctx.fillStyle = color;
      stars.forEach((s) => {
        if (!reduceMotion) {
          s.y -= s.speed * 0.3;
          s.x += s.speed * 0.15;
          if (s.y < 0) s.y = height;
          if (s.x > width) s.x = 0;
        }
        ctx.globalAlpha = 0.15 + 0.85 * Math.abs(Math.sin((time / 1000) * s.rate + s.phase));
        ctx.beginPath();
        ctx.arc(s.x, s.y, s.r, 0, Math.PI * 2);
        ctx.fill();
      });
      if (!reduceMotion) frame = requestAnimationFrame(draw);
    };

    setup();
    frame = requestAnimationFrame(draw);
    window.addEventListener('resize', setup);
    return () => {
      cancelAnimationFrame(frame);
      window.removeEventListener('resize', setup);
    };
  }, [density, color]);

  return <canvas ref={canvasRef} className={className} aria-hidden="true" />;
};

export default Sparkles;
