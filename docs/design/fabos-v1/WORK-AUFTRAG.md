# Kopierbarer Work-Auftrag – Designpaket übernehmen (Phase 1: LESEN)

**Repository:** `butzi124-ship-it/schichtplan-mvp`  
**Design-Branch:** `design/fabos-reference-v1` (ausschließlich Referenzmaterial)  
**Designpaket:** `docs/design/fabos-v1/`  
**Anwendungsbasis:** `design/reference-production-package1` · v0.20.6 · PR #45.

## Auftrag
1. Lies `AGENTS.md`, `docs/DEVELOPMENT.md` auf dem bestehenden Entwicklungsbranch und alle Dateien aus `docs/design/fabos-v1/` auf dem Design-Branch.
2. Verifiziere die **tatsächlich existierenden** UI-Bereiche, DOM-Struktur, CSS-Tokens und die 33 visuellen Test-Szenen.
3. Erstelle eine präzise Zuordnung: welche Gestaltung gilt in welcher bestehenden Ansicht; welche vorhandenen CSS-/JS-Präsentationsdateien wären betroffen.
4. Vergleiche die 33 Review-only-Aufnahmen aus GitHub Actions 43726264797 mit dem Designpaket. Die SVGs sind Konzeptmuster, keine Snapshots.
5. Markiere Widersprüche, Unklarheiten, Brand-/Logo- und Lizenzrisiken sowie mögliche Regressionen.
6. Schlage **maximal drei kleine Implementierungs-Mini-Aufträge** in sinnvoller Reihenfolge vor, pro Auftrag klar begrenzte Dateien und Tests.

## Strikte Grenzen
- **Nur lesen/analysieren.** Keine Implementierung in diesem Auftrag.
- **Kein** Umbenennen von „Humbel-App“ zu FABOS im Code; hierzu neue Freigabe erforderlich.
- Keine neuen Module, Datenbanken, RLS, Auth, Geschäftsvorgänge oder Rechte anlegen.
- Keine Snapshot-Referenzen ersetzen, keine Bildtoleranzen erhöhen und keine 853 funktionalen Tests beeinträchtigen.
- Keine GitHub-Schreibvorgänge, Branches, Commits oder PRs; kein Merge, Deployment oder Datenbankzugriff.

## Bericht (maximal 350 Wörter)
- Gefundene Referenzdateien mit GitHub-Pfaden.
- Was davon bereits umgesetzt ist, was abweicht.
- Drei priorisierte Mini-Aufträge mit Dateiumfang und Testplan.
- Explizite Fragen/Blocker, die vor dem Design-Merge zu entscheiden sind.

**Danach stoppen und meine Freigabe abwarten.**
