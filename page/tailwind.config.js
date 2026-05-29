/** Psi Planner — configuração Tailwind (tokens da identidade) */
module.exports = {
  content: ["./*.html"],
  theme: {
    extend: {
      colors: {
        wine: {
          DEFAULT: "#5A1E2A", // vinho marca
          dark: "#3A1018",    // vinho escuro
          light: "#8E3B4C",   // vinho claro / acento
        },
        gold: "#C9A86A",       // dourado champanhe (acento)
        goldsoft: "#E0CBA0",   // dourado claro p/ texto sobre vinho (contraste AA)
        offwhite: "#FAFAF8",   // fundo claro
        ink: "#3B3B3B",        // tinta
        charcoal: "#1A1719",   // fundo dark
        ondark: "#E8E4E6",     // texto sobre dark
        ondarkmute: "#B9B0B4", // texto secundário sobre dark (AA)
        status: {
          green: "#D9EAD3",
          red: "#F4CCCC",
          yellow: "#FFF2CC",
        },
      },
      fontFamily: {
        display: ['Prata', 'Georgia', 'serif'],
        body: ['Inter', 'system-ui', 'sans-serif'],
      },
      letterSpacing: {
        widest2: '0.22em',
      },
      maxWidth: {
        content: '1200px',
      },
      boxShadow: {
        soft: '0 10px 40px -12px rgba(58,16,24,0.18)',
        card: '0 4px 24px -8px rgba(58,16,24,0.12)',
      },
    },
  },
  plugins: [],
};
