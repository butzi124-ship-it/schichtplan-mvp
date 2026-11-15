(function () {
"use strict";
const esc = value => String(value == null ? "" : value).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
function tile(title, hint, symbol, action, featured) {
  const cls = "h-home-tile" + (featured ? " featured" : "") + (action ? "" : " unavailable");
  const body = '<span class="h-home-symbol" aria-hidden="true">' + symbol + '</span><strong>' + esc(title) + '</strong><small>' + esc(hint) + '</small>';
  return action ? '<button type="button" class="' + cls + '" onclick="' + action + '">' + body + '</button>' : '<div class="' + cls + '">' + body + '</div>';
}
function box(title, icon, contents) {
  return '<section class="h-home-panel"><h2><span aria-hidden="true">' + icon + '</span> ' + esc(title) + '</h2>' + contents + '</section>';
}
const empty = message => '<p class="h-home-empty">' + esc(message) + '</p>';
window.HumbelHome = Object.freeze({ render(model) {
  const tabs = model.tabs || [];
  const can = tab => tabs.includes(tab);
  const tiles = [
    tile("Fertigung", "Abteilungen · Maschinen · Aufträge", "▥", can("produktion") ? "setTab('produktion')" : "", true),
    tile("Schichtplan", "Schichten · Planung", "▦", model.shiftAction, false),
    tile("Mitarbeiter", "Stammdaten · Team", "♙", can("personalverwaltung") ? "setTab('personalverwaltung')" : "", false),
    tile("Werkzeuge", "Bestand · Journal", "⚒", can("werkzeugverwaltung") ? "setTab('werkzeugverwaltung')" : "", false),
    tile("Artikel", "Stammdaten", "▤", model.articleEnabled && can("produktion") ? "setTab('produktion')" : "", false)
  ].join("");
  const machines = (model.machines || []).slice(0, 8);
  const machineBody = machines.length ? '<div class="h-home-machines">' + machines.map(m => '<div class="h-home-machine"><strong>' + esc(m.name || m.machine_code || m.id) + '</strong><small>Status nicht verfügbar</small></div>').join("") + '</div>' : empty("Keine sichtbaren Maschineninformationen vorhanden.");
  const shifts = (model.shifts || []).slice(0, 4);
  const shiftBody = shifts.length ? '<div class="h-home-list">' + shifts.map(s => '<div><strong>' + esc(s.label || "Schicht") + '</strong><span>' + esc(s.date || "") + ' · ' + esc(s.start || "") + '–' + esc(s.end || "") + '</span></div>').join("") + '</div>' : empty("Keine persönliche Schicht-Kurzansicht verfügbar.");
  const tasks = (model.tasks || []).slice(0, 4);
  const taskBody = tasks.length ? '<div class="h-home-list">' + tasks.map(t => '<div><strong>' + esc(t.title || "Aufgabe") + '</strong><span>' + esc(t.status || "Offen") + '</span></div>').join("") + '</div>' : empty("Keine offenen Aufgaben vorhanden.");
  const countBody = (model.counts || []).length ? '<div class="h-home-list"><div><strong>Sichtbare Stückzahleinträge</strong><span>' + model.counts.length + '</span></div></div>' : empty("Keine verfügbaren Fertigungskennzahlen.");
  return '<main class="h-home"><header class="h-home-hero"><h1>Start / Dashboard</h1><p>Fertigung, Planung und Team – alles in einer Anwendung.</p></header><div class="h-home-tiles">' + tiles + '</div><div class="h-home-panels">' +
    box("Heute im Überblick", "▦", countBody) +
    box("Maschinenübersicht", "⚙", machineBody) +
    box("Fertigungsinformationen", "▥", empty(can("produktion") ? "Auftragsdetails sind in der Fertigung verfügbar." : "Keine Fertigungsinformationen für diese Rolle verfügbar.")) +
    box("Schichtplan – Meine Schichten", "◷", shiftBody) +
    box("Offene Aufgaben", "◎", taskBody) +
    box("Tagesfortschritt Fertigung", "◉", empty("Kein verlässlich berechneter Tagesfortschritt verfügbar.")) + '</div></main>';
}});
}());