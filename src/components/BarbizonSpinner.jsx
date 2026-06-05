import React from 'react';

export default function BarbizonSpinner({ size = 80, speed = '1.5s', message = 'Powering Systems...' }) {
  // Brand aesthetic: Deep stage navy blue backgrounds require a crisp white / luminous cyan-blue beam
  const brandDarkBlue = '#0B1B3D'; 
  const brandLightBeam = '#38BDF8'; // Luminous theater spotlight gel effect

  return (
    <div style={styles.container}>
      <div style={{ ...styles.spinnerWrapper, width: size, height: size }}>
        <svg
          viewBox="0 0 100 100"
          style={{
            animation: `barbizon-spin ${speed} linear infinite`,
            width: '100%',
            height: '100%',
          }}
        >
          {/* Custom Definitions for Spotlight Glow */}
          <defs>
            <linearGradient id="spotlightGlow" x1="0%" y1="0%" x2="100%" y2="100%">
              <stop offset="0%" stopColor={brandLightBeam} stopOpacity="1" />
              <stop offset="100%" stopColor={brandDarkBlue} stopOpacity="0.1" />
            </linearGradient>
          </defs>

          {/* Background Outer Ring (The Fixture Housing) */}
          <circle
            cx="50"
            cy="50"
            r="44"
            fill="none"
            stroke="#1E293B"
            strokeWidth="3"
          />

          {/* Stepped Fresnel Lens Grooves */}
          <circle cx="50" cy="50" r="34" fill="none" stroke="#334155" strokeWidth="1" strokeDasharray="4 4" />
          <circle cx="50" cy="50" r="24" fill="none" stroke="#334155" strokeWidth="1" strokeDasharray="6 3" />

          {/* The Spinning Light Beam Segment */}
          <path
            d="M 50 6 A 44 44 0 0 1 94 50"
            fill="none"
            stroke="url(#spotlightGlow)"
            strokeWidth="7"
            strokeLinecap="round"
          />

          {/* Central Spotlight Core Bulbs */}
          <circle cx="50" cy="50" r="10" fill={brandDarkBlue} stroke={brandLightBeam} strokeWidth="2" />
          <circle cx="50" cy="50" r="4" fill="#FFFFFF" />
        </svg>
      </div>
      
      {/* Optional branded status text */}
      {message && <p style={styles.text}>{message}</p>}

      {/* Embedded CSS Keyframes for smooth animation loop */}
      <style>{`
        @keyframes barbizon-spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  );
}

const styles = {
  container: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    fontFamily: 'system-ui, -apple-system, sans-serif',
    padding: '20px',
  },
  spinnerWrapper: {
    position: 'relative',
  },
  text: {
    marginTop: '16px',
    color: '#64748B',
    fontSize: '14px',
    fontWeight: '500',
    letterSpacing: '0.05em',
    textTransform: 'uppercase',
  }
};
