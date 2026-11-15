// Isolated test-page measurements; the containment probe restores every inline style.
function productionOverviewGeometry() {
  const describe = el => {
    const r = el.getBoundingClientRect(), c = getComputedStyle(el);
    return { tag: el.tagName, id: el.id, class: typeof el.className === 'string' ? el.className : el.className.baseVal,
      text: el.textContent.trim().replace(/\s+/g, ' ').slice(0, 110), left: r.left, right: r.right, width: r.width,
      clientWidth: el.clientWidth, scrollWidth: el.scrollWidth,
      offsetParent: el.offsetParent ? { tag: el.offsetParent.tagName, id: el.offsetParent.id, class: el.offsetParent.className } : null,
      css: Object.fromEntries(['display','position','width','min-width','max-width','box-sizing','overflow-x','white-space','overflow-wrap','flex','flex-wrap','grid-template-columns','font-size','gap','padding-left','padding-right','clip-path'].map(p => [p, c.getPropertyValue(p)])) };
  };
  const outside = [...document.querySelectorAll('body *')].filter(el => {
    const r = el.getBoundingClientRect(), c = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && c.visibility !== 'hidden' && c.position !== 'fixed'
      && (r.right > document.documentElement.clientWidth + 1 || r.left < -1)
      && (!el.closest('.h-table-scroll') || c.position === 'absolute');
  }).map(el => ({ ...describe(el), parents: [el.parentElement, el.parentElement?.parentElement].filter(Boolean).map(describe) }));
  const table = document.querySelector('.h-production-table-card .h-table-scroll');
  const absoluteHeaders = () => [...document.querySelectorAll('.h-production-table-card .h-table-scroll .h-sr-only')].map(describe);
  const report = { innerWidth, visualViewportWidth: visualViewport.width,
    documentClientWidth: document.documentElement.clientWidth, documentScrollWidth: document.documentElement.scrollWidth,
    candidates: outside.sort((a,b) => b.right - a.right).slice(0, 30),
    keyElements: ['.h-app-header','.h-top-search','.h-production-overview','.h-production-heading','.h-production-heading > div','.h-production-filters','.h-production-split','.h-reference-detail'].map(s => document.querySelector(s)).filter(Boolean).map(describe),
    tables: [...document.querySelectorAll('.h-table-scroll')].map(describe), kpi: [...document.querySelectorAll('.h-kpi')].map(describe), absoluteHeaders: absoluteHeaders() };
  if (table && getComputedStyle(table).position === 'relative') {
    const original = table.getAttribute('style');
    try {
      // Recreate only the parent's former containing-block rule, without touching app data or assertions.
      table.style.position = 'static';
      report.containmentProbe = { formerPosition: getComputedStyle(table).position, formerDocumentScrollWidth: document.documentElement.scrollWidth, formerAbsoluteHeaders: absoluteHeaders() };
    } finally {
      if (original === null) table.removeAttribute('style'); else table.setAttribute('style', original);
    }
    report.containmentProbe.restoredPosition = getComputedStyle(table).position;
    report.containmentProbe.restoredDocumentScrollWidth = document.documentElement.scrollWidth;
    report.containmentProbe.originalInlineStyleRestored = table.getAttribute('style') === original;
  }
  return report;
}
module.exports = { productionOverviewGeometry };
