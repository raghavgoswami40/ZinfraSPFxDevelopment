import * as React from 'react';

interface FlowButtonProps {
  text?: string;
  onClick?: () => void;
  disabled?: boolean;
}

const ArrowSvg: React.FC<{ wrapStyle: React.CSSProperties; strokeColor: string }> = ({ wrapStyle, strokeColor }) => (
  <svg
    width="14"
    height="14"
    viewBox="0 0 24 24"
    fill="none"
    strokeLinecap="round"
    strokeLinejoin="round"
    style={wrapStyle}
  >
    <path d="M5 12h14M12 5l7 7-7 7" strokeWidth="2.5" stroke={strokeColor} />
  </svg>
);

export const FlowButton: React.FC<FlowButtonProps> = ({ text = 'Go to process', onClick, disabled = false }) => {
  const [hovered, setHovered] = React.useState(false);
  const active = hovered && !disabled;

  const btnStyle: React.CSSProperties = {
    position: 'relative',
    display: 'inline-flex',
    alignItems: 'center',
    overflow: 'hidden',
    borderRadius: active ? 12 : 100,
    border: `1.5px solid ${active ? 'transparent' : 'rgba(255,255,255,0.7)'}`,
    background: 'transparent',
    padding: '6px 24px',
    fontSize: 11,
    fontWeight: 600,
    color: active ? '#fff' : '#fff',
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? 0.4 : 1,
    userSelect: 'none',
    whiteSpace: 'nowrap',
    minWidth: 110,
    justifyContent: 'center',
    transition: [
      'border 600ms cubic-bezier(0.23,1,0.32,1)',
      'border-radius 600ms cubic-bezier(0.23,1,0.32,1)',
    ].join(', '),
  };

  const arrLeftStyle: React.CSSProperties = {
    position: 'absolute',
    left: active ? 10 : '-30%',
    top: '50%',
    transform: 'translateY(-50%)',
    zIndex: 9,
    transition: 'left 800ms cubic-bezier(0.34,1.56,0.64,1)',
    pointerEvents: 'none',
  };

  const textStyle: React.CSSProperties = {
    position: 'relative',
    zIndex: 1,
    transform: `translateX(${active ? 8 : -8}px)`,
    transition: 'transform 800ms ease-out',
  };

  const circleStyle: React.CSSProperties = {
    position: 'absolute',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    width: active ? 240 : 14,
    height: active ? 240 : 14,
    background: 'rgba(0,0,0,0.7)',
    borderRadius: '50%',
    opacity: active ? 1 : 0,
    pointerEvents: 'none',
    transition: [
      'width 800ms cubic-bezier(0.19,1,0.22,1)',
      'height 800ms cubic-bezier(0.19,1,0.22,1)',
      'opacity 800ms cubic-bezier(0.19,1,0.22,1)',
    ].join(', '),
  };

  const arrRightStyle: React.CSSProperties = {
    position: 'absolute',
    right: active ? '-30%' : 10,
    top: '50%',
    transform: 'translateY(-50%)',
    zIndex: 9,
    transition: 'right 800ms cubic-bezier(0.34,1.56,0.64,1)',
    pointerEvents: 'none',
  };

  return (
    <button
      style={btnStyle}
      onMouseEnter={() => { if (!disabled) setHovered(true); }}
      onMouseLeave={() => setHovered(false)}
      onClick={disabled ? undefined : onClick}
    >
      <ArrowSvg wrapStyle={arrLeftStyle} strokeColor="#fff" />
      <span style={textStyle}>{text}</span>
      <span style={circleStyle} />
      <ArrowSvg wrapStyle={arrRightStyle} strokeColor="#fff" />
    </button>
  );
};
