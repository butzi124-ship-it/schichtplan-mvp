// Read-only DOM measurements. Runs only against the isolated visual-test page.
function productionOverviewGeometry() {
  const describe = el => {
    const r = el.getBoundingClientRect(), c = getComputedStyle(el);
    return { tag: el.tagName, id: el.id, class: typeof el.className === 'string' ? el.className : el.className.baseVal,
      text: el.textContent.trim().replace(/\s+/g, ' ').slice(0, 110), left: r.left, right: r.right, width: r.width,
      clientWidth: el.clientWidth, scrollWidth: el.scrollWidth,
      css: Object.fromEntries(['display','position','width','min-width','max-width','box-sizing','overflow-x','white-space','overflow-wrap','flex','flex-wrap','grid-template-columns','font-size','gap','padding-left','padding-right'].map(p => [p, c.getPropertyValue(p)])) };
  };
  const viewport = 390;
  const outside = [...document.querySelectorAll('body *')].filter(el => {
    const r = el.getBoundingClientRect(), c = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && c.visibility !== 'hidden' && c.position !== 'fixed' && c.clipPath === 'none'
      && (r.right > document.documentElement.clientWidth + 1 || r.left < -1) && !el.closest('.h-table-scroll');
  }).map(el => ({ ...describe(el), parents: [el.parentElement, el.parentElement?.parentElement].filter(Boolean).map(describe) }));
  return { configuredMobileWidth: viewport, innerWidth, visualViewportWidth: visualViewport.width,
    documentClientWidth: document.documentElement.clientWidth, documentScrollWidth: document.documentElement.scrollWidth,
    candidates: outside.sort((a,b) => b.right - a.right).slice(0, 30),
    keyElements: ['.h-app-header','.h-top-search','.h-production-overview','.h-production-heading','.h-production-heading > div','.h-production-filters','.h-production-split','.h-reference-detail'].map(s => document.querySelector(s)).filter(Boolean).map(describe),
    tables: [...document.querySelectorAll('.h-table-scroll')].map(describe), kpi: [...document.querySelectorAll('.h-kpi')].map(describe) };
}
module.exports = { productionOverviewGeometry };
