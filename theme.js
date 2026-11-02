/* ═══════════════════════════════════════════════════════════════
   THEME SWITCHER — Shared JavaScript for all pages (Fixed to Light)
   ═══════════════════════════════════════════════════════════════ */

(function() {
  'use strict';

  // ── Apply theme to document ──
  function applyTheme(themeId) {
    document.documentElement.setAttribute('data-theme', themeId);
  }

  // ── Initialize ──
  // Always enforce the light/white theme as default and do not render switcher UI
  applyTheme('light');

  // Expose helper API
  window.IITTheme = {
    apply: () => applyTheme('light'),
    getThemes: () => [{ id: 'light', name: 'Light', icon: '☀️' }],
    getCurrent: () => 'light'
  };
})();
