/* ═══════════════════════════════════════════════════════════════
   THEME SWITCHER — Shared JavaScript for all pages (Red Theme Default)
   ═══════════════════════════════════════════════════════════════ */

(function() {
  'use strict';

  // ── Apply theme to document ──
  function applyTheme(themeId) {
    document.documentElement.setAttribute('data-theme', themeId || 'red');
  }

  // ── Initialize ──
  // Always enforce the Crimson Red theme as default
  applyTheme('red');

  // Expose helper API
  window.IITTheme = {
    apply: () => applyTheme('red'),
    getThemes: () => [{ id: 'red', name: 'Crimson Red', icon: '🔴' }],
    getCurrent: () => 'red'
  };
})();
