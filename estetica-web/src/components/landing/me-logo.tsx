type MeLogoProps = {
  gradientId: string;
  className?: string;
};

export function MeLogo({ gradientId, className }: MeLogoProps) {
  return (
    <svg
      className={className}
      viewBox="0 0 200 140"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden
    >
      <text
        x="10"
        y="100"
        fontFamily="Georgia, serif"
        fontSize="110"
        fontWeight="300"
        fill={`url(#${gradientId})`}
        letterSpacing="-5"
      >
        M
      </text>
      <text
        x="105"
        y="100"
        fontFamily="Georgia, serif"
        fontSize="100"
        fontWeight="300"
        fill={`url(#${gradientId})`}
      >
        E
      </text>
      <path
        d="M15 108 Q100 120 185 108"
        stroke={`url(#${gradientId})`}
        strokeWidth="1.2"
        fill="none"
      />
      <text
        x="100"
        y="130"
        fontFamily="system-ui, sans-serif"
        fontSize="13"
        fontWeight="400"
        fill="#C9A84C"
        textAnchor="middle"
        letterSpacing="6"
      >
        ESTETICA
      </text>
      <text x="158" y="118" fontSize="12" fill="#C9A84C">
        ✦
      </text>
      <defs>
        <linearGradient id={gradientId} x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor="#E8C97A" />
          <stop offset="50%" stopColor="#C9A84C" />
          <stop offset="100%" stopColor="#A07830" />
        </linearGradient>
      </defs>
    </svg>
  );
}
