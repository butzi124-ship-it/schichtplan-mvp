/* Presentation only. Reads the rendered navigation, never app authority or data. */
(function () {
  'use strict';
  function install() {
    const nav = document.getElementById('tabs');
    const login = document.getElementById('loginBox');
    const toggle = document.getElementById('app-menu-toggle');
    const view = document.getElementById('view');
    if (!nav || !login || !toggle || !view) return;
    let menuOpen = true;
    function setMenu(open) {
      menuOpen = open;
      document.body.dataset.hMenuOpen = String(open);
      toggle.setAttribute('aria-expanded', String(open));
    }
    function sync() {
      const signedIn = login.classList.contains('hidden') && !nav.classList.contains('hidden');
      document.body.classList.toggle('h-app-signed-in', signedIn);
      if (!signedIn) setMenu(true);
      // On first login the app resolves its fallback tab after drawing the menu.
      // Its first permitted entry is the visible fallback; style that existing entry.
      if (signedIn && !nav.querySelector('.humbel-main-tab-active')) {
        const first = nav.querySelector('button.humbel-main-tab');
        first?.classList.remove('humbel-main-tab');
        first?.classList.add('humbel-main-tab-active');
      }
      nav.querySelectorAll('button').forEach(button => {
        if (button.classList.contains('humbel-main-tab-active')) button.setAttribute('aria-current', 'page');
        else button.removeAttribute('aria-current');
      });
      view.querySelectorAll('.h-table-scroll').forEach(region => {
        region.setAttribute('tabindex', '0');
        region.setAttribute('role', 'region');
        region.setAttribute('aria-label', region.dataset.tableLabel || 'Tabelle');
      });
    }
    toggle.addEventListener('click', () => setMenu(!menuOpen));
    nav.addEventListener('keydown', event => {
      if (event.key === 'Escape' && matchMedia('(max-width: 800px)').matches) {
        setMenu(false);
        toggle.focus();
      }
    });
    // Existing renders replace buttons. Restore keyboard focus to their new equivalent.
    document.addEventListener('click', event => {
      const button = event.target.closest('button');
      if (!button || !button.matches('#tabs button, .humbel-subtab, .humbel-subtab-active')) return;
      const action = button.getAttribute('onclick');
      const scope = nav.contains(button) ? nav : view;
      requestAnimationFrame(() => {
        sync();
        if (!button.isConnected) {
          const next = Array.from(scope.querySelectorAll('button')).find(b => b.getAttribute('onclick') === action);
          next?.focus({ preventScroll: true });
        }
      });
    }, true);
    new MutationObserver(sync).observe(nav, { childList: true, attributes: true, attributeFilter: ['class'] });
    new MutationObserver(sync).observe(login, { attributes: true, attributeFilter: ['class'] });
    new MutationObserver(sync).observe(view, { childList: true, subtree: true });
    setMenu(true);
    sync();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install);
  else install();
})();
