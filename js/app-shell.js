/* Presentation only. Reads the rendered navigation, never app authority or data. */
(function () {
  'use strict';
  function install() {
    const nav = document.getElementById('tabs');
    const login = document.getElementById('loginBox');
    const toggle = document.getElementById('app-menu-toggle');
    const view = document.getElementById('view');
    if (!nav || !login || !toggle || !view) return;
    const sessionInfo = document.getElementById('sessionInfo');
    // A session announcement is not the status result of an article/form action.
    sessionInfo?.removeAttribute('role');
    sessionInfo?.setAttribute('aria-live', 'polite');
    let menuOpen = true;
    let lastNavigationFocus = null;
    let observedBanner = null;
    const bannerObserver = new ResizeObserver(updateBannerOffset);
    function updateBannerOffset() {
      const height = observedBanner ? Math.ceil(observedBanner.getBoundingClientRect().height) : 0;
      const value = height + 'px';
      if (document.documentElement.style.getPropertyValue('--h-preview-offset') !== value) {
        document.documentElement.style.setProperty('--h-preview-offset', value);
      }
    }
    function setMenu(open) {
      menuOpen = open;
      document.body.dataset.hMenuOpen = String(open);
      toggle.setAttribute('aria-expanded', String(open));
    }
    function sync() {
      const banner = document.getElementById('isolated-preview-banner');
      if (banner !== observedBanner) {
        bannerObserver.disconnect();
        observedBanner = banner;
        if (banner) bannerObserver.observe(banner);
        updateBannerOffset();
      }
      const utilityNodes = ['versionFooter', 'context-help-launcher'].map(id => document.getElementById(id)).filter(Boolean);
      if (utilityNodes.length) {
        let dock = document.getElementById('app-utilities');
        if (!dock) {
          dock = document.createElement('footer');
          dock.id = 'app-utilities';
          dock.setAttribute('aria-label', 'Hilfe und Versionsinformationen');
          document.getElementById('app-shell').appendChild(dock);
        }
        for (const node of utilityNodes) if (node.parentElement !== dock) dock.appendChild(node);
        dock.hidden = utilityNodes.every(node => node.hidden);
      }
      const downtime = document.getElementById('downtimeWidget');
      const header = document.querySelector('.h-app-header');
      if (downtime && header && downtime.parentElement !== header) header.appendChild(downtime);
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
        if (button.textContent === 'Sicherheitsunterweisungen' && !button.querySelector('wbr')) {
          button.replaceChildren(document.createTextNode('Sicherheits'), document.createElement('wbr'), document.createTextNode('unterweisungen'));
        }
        if (button.classList.contains('humbel-main-tab-active')) button.setAttribute('aria-current', 'page');
        else button.removeAttribute('aria-current');
      });
      view.querySelectorAll('.h-table-scroll').forEach(region => {
        region.setAttribute('tabindex', '0');
        region.setAttribute('role', 'region');
        region.setAttribute('aria-label', region.dataset.tableLabel || 'Tabelle');
      });
      for (const [id, label] of [['taskTitle', 'Aufgabe'], ['taskAssignee', 'Mitarbeiter'], ['taskDeadline', 'Frist']]) {
        document.getElementById(id)?.setAttribute('aria-label', label);
      }
      if (lastNavigationFocus && !lastNavigationFocus.element.isConnected && document.activeElement === document.body) {
        const next = Array.from(lastNavigationFocus.scope.querySelectorAll('button')).find(b => b.getAttribute('onclick') === lastNavigationFocus.action);
        if (next) next.focus({ preventScroll: true });
      }
    }
    document.addEventListener('focusin', event => {
      const button = event.target.closest('button');
      lastNavigationFocus = button?.matches('#tabs button, .humbel-subtab, .humbel-subtab-active')
        ? { element: button, action: button.getAttribute('onclick'), scope: nav.contains(button) ? nav : view }
        : null;
    });
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
    // Includes utilities created after DOMContentLoaded and the loopback-only banner.
    new MutationObserver(sync).observe(document.body, { childList: true, subtree: true });
    setMenu(true);
    sync();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install);
  else install();
})();
