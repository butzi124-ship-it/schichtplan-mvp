(function(root){const topics=[
  {
    "id": "home",
    "title": "Startseite",
    "tab": "dashboard",
    "subtab": "overview",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Schichten, Aufgaben und verfügbare Fertigungsübersichten über Schnellzugriffe öffnen.",
    "fields": [
      {
        "name": "Schnellzugriff",
        "description": "Schnellzugriff im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Ansichtsbereich",
        "description": "Ansichtsbereich im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Start / Dashboard öffnen",
      "Gewünschten Schnellzugriff wählen",
      "Im Zielbereich die verfügbaren Daten prüfen"
    ],
    "effect": "Es werden vorhandene Daten angezeigt. Ein leerer Bereich bedeutet keinen erfundenen Maschinenstatus.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Startseite in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "my-shifts",
    "title": "Meine Schichten",
    "tab": "dashboard",
    "subtab": "myShifts",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Eigene zugewiesene Schichten und Aufgaben ansehen.",
    "fields": [
      {
        "name": "Zeitraum",
        "description": "Zeitraum im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Schicht",
        "description": "Schicht im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Meine Schichten öffnen",
      "Datum und Schicht prüfen"
    ],
    "effect": "Lesen verändert keine Planung. Verbindliche Veröffentlichung und Datenschutz müssen nativ abgenommen werden.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Meine Schichten in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "tasks",
    "title": "Aufgaben",
    "tab": "dashboard",
    "subtab": "todo",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Zugewiesene Aufgaben verfolgen.",
    "fields": [
      {
        "name": "Aufgabe",
        "description": "Aufgabe im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Status",
        "description": "Status im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Termin",
        "description": "Termin im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "To-Do öffnen",
      "Zugewiesene Aufgabe prüfen",
      "Erlaubte Statusänderung bestätigen"
    ],
    "effect": "Aufgabenänderungen folgen dem bestehenden Ablauf; keine Erweiterung von Mitarbeiterrechten.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Aufgaben in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "schedule",
    "title": "Schichtplan",
    "tab": "dashboard",
    "subtab": "oldSchedule",
    "role": "employee",
    "status": "visible_placeholder",
    "purpose": "Vorhandener Navigationsplatzhalter bzw. deaktivierte Altansicht; hier ist kein fertiger Bearbeitungsablauf verfügbar.",
    "fields": [
      {
        "name": "Entwicklungsstand",
        "description": "Die angezeigte Einschränkung beachten; keine Buchung über diesen Platzhalter vornehmen."
      }
    ],
    "steps": [
      "Ansicht öffnen und Entwicklungsstand prüfen."
    ],
    "effect": "Diese Ansicht aktiviert keine neue Schichtplanung oder Speicherung.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Schichtplan in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "employees",
    "title": "Mitarbeiterverwaltung",
    "tab": "personalverwaltung",
    "subtab": "employees",
    "role": "admin",
    "status": "implemented_not_deployed",
    "purpose": "Mitarbeiterstammdaten und Aktivität verwalten.",
    "fields": [
      {
        "name": "Vorname",
        "description": "Vorname im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Nachname",
        "description": "Nachname im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Personalnummer",
        "description": "Personalnummer im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Rolle",
        "description": "Sicherheitskritische Zuordnung; keine Selbstermächtigung."
      },
      {
        "name": "Aktiv",
        "description": "Aktiv im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Abteilung",
        "description": "Abteilung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Personalverwaltung öffnen",
      "Mitarbeiter auswählen",
      "Zulässige Felder prüfen und speichern"
    ],
    "effect": "Stammdatenänderungen sind sicherheitsrelevant. Rollen und Auth-Zuordnung benötigen die noch separat freizugebende DB-Härtung.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Mitarbeiterverwaltung in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "shifts",
    "title": "Schichtmodell",
    "tab": "personalverwaltung",
    "subtab": "shiftModel",
    "role": "admin",
    "status": "implemented_not_deployed",
    "purpose": "Vorhandene Schichtdefinitionen mit Zeiten pflegen.",
    "fields": [
      {
        "name": "Name",
        "description": "Name im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Beginn",
        "description": "Beginn im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Ende",
        "description": "Ende im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Aktiv",
        "description": "Aktiv im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Schichtmodell öffnen",
      "Definition auswählen",
      "Zeiten prüfen und speichern"
    ],
    "effect": "Definitionen beeinflussen Planung; vor Änderungen Abhängigkeiten prüfen.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Schichtmodell in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "planning",
    "title": "Schichtplanung",
    "tab": "personalverwaltung",
    "subtab": "shiftPlanning",
    "role": "leader",
    "status": "visible_placeholder",
    "purpose": "Vorhandener Navigationsplatzhalter; dieser Bereich ist fachlich noch nicht implementiert.",
    "fields": [
      {
        "name": "Woche",
        "description": "Woche im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Mitarbeiter",
        "description": "Mitarbeiter im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Schicht",
        "description": "Schicht im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Vertretung",
        "description": "Vertretung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Ansicht öffnen und ausgewiesenen Entwicklungsstand prüfen."
    ],
    "effect": "Kein fertig implementierter Speichervorgang. Keine neue Funktion als verfügbar ausgeben.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Schichtplanung in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "personnel-settings",
    "title": "Personaleinstellungen",
    "tab": "personalverwaltung",
    "subtab": "settings",
    "role": "admin",
    "status": "visible_placeholder",
    "purpose": "Vorhandener Navigationsplatzhalter; dieser Bereich ist fachlich noch nicht implementiert.",
    "fields": [
      {
        "name": "Einstellung",
        "description": "Einstellung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Status",
        "description": "Status im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Ansicht öffnen und ausgewiesenen Entwicklungsstand prüfen."
    ],
    "effect": "Kein fertig implementierter Speichervorgang. Keine neue Funktion als verfügbar ausgeben.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Personaleinstellungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "departments",
    "title": "Abteilungen",
    "tab": "produktion",
    "subtab": "departments",
    "role": "admin",
    "status": "implemented_not_deployed",
    "purpose": "Abteilungen und bestehende Zuständigkeiten verwalten.",
    "fields": [
      {
        "name": "Name",
        "description": "Name im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Kürzel",
        "description": "Kürzel im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Leitung",
        "description": "Leitung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Aktiv",
        "description": "Aktiv im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Produktion öffnen",
      "Abteilungen wählen",
      "Zuordnung prüfen"
    ],
    "effect": "Mehrfachmitgliedschaft und Organisationsrollen aus 8A sind noch nicht aktiviert.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Abteilungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "machines",
    "title": "Maschinenverwaltung",
    "tab": "produktion",
    "subtab": "machines",
    "role": "leader",
    "status": "implemented_not_deployed",
    "purpose": "Maschinen ihren Produktionsabteilungen zuordnen.",
    "fields": [
      {
        "name": "Name",
        "description": "Name im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Maschinencode",
        "description": "Maschinencode im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Abteilung",
        "description": "Abteilung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Aktiv",
        "description": "Aktiv im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Maschinen öffnen",
      "Maschine auswählen",
      "Zuständigkeit und Status prüfen"
    ],
    "effect": "Maschinenänderungen sind von einer bestätigten physischen OP-Übergabe getrennt.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Maschinenverwaltung in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "orders",
    "title": "Betriebsaufträge und Operationen",
    "tab": "produktion",
    "subtab": "orders",
    "role": "leader",
    "status": "implemented_not_deployed",
    "purpose": "BA anlegen und vorhandene OP-Snapshots verwenden.",
    "fields": [
      {
        "name": "BA-Nummer",
        "description": "Auftragsbezogene eindeutige Nummer, unabhängig vom Artikel."
      },
      {
        "name": "Artikelnummer",
        "description": "Artikelnummer im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Stückzahl",
        "description": "Positive Auftrags- oder Buchungsmenge im jeweils angezeigten Kontext."
      },
      {
        "name": "Liefertermin",
        "description": "Liefertermin im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Maschine",
        "description": "Maschine im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Aufträge / BA öffnen",
      "Artikel und Pflichtfelder prüfen",
      "BA anlegen oder vorhandenen Auftrag öffnen"
    ],
    "effect": "BA und Artikelrevision bleiben getrennt. Doppelte Nummern und veraltete Revisionen werden vom jeweiligen RPC geprüft.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Betriebsaufträge und Operationen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "counts",
    "title": "Mengen und Rüstzeiten",
    "tab": "produktion",
    "subtab": "counts",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Eigene Mengen und vorbereitete Rüstzeiterfassung benutzen.",
    "fields": [
      {
        "name": "Maschine",
        "description": "Maschine im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "BA",
        "description": "BA im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "OP",
        "description": "OP im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Gut",
        "description": "Gut im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Ausschuss",
        "description": "Ausschuss im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "In Abklärung",
        "description": "In Abklärung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Rüstmodus",
        "description": "Rüstmodus im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Stückzahl öffnen",
      "Maschine und Auftrag prüfen",
      "Zulässige Buchung erfassen und Ergebnis prüfen"
    ],
    "effect": "Keine Doppelbuchung. Datenbankfehler zuerst klären; neue Request-ID nicht zur blinden Wiederholung benutzen.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Mengen und Rüstzeiten in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "setups",
    "title": "Spannungen",
    "tab": "produktion",
    "subtab": "setups",
    "role": "leader",
    "status": "visible_placeholder",
    "purpose": "Vorhandener Navigationsplatzhalter bzw. deaktivierte Altansicht; hier ist kein fertiger Bearbeitungsablauf verfügbar.",
    "fields": [
      {
        "name": "Entwicklungsstand",
        "description": "Die angezeigte Einschränkung beachten; keine Buchung über diesen Platzhalter vornehmen."
      }
    ],
    "steps": [
      "Ansicht öffnen und Entwicklungsstand prüfen."
    ],
    "effect": "Diese Registerkarte ist ein Platzhalter. Vorhandene OP-Mengen- und Spannungsfunktionen befinden sich innerhalb der Stückzahlerfassung und sind davon getrennt zu prüfen.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Spannungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "clarification",
    "title": "Ausschuss und Abklärung",
    "tab": "produktion",
    "subtab": "scrap",
    "role": "leader",
    "status": "visible_placeholder",
    "purpose": "Vorhandener Navigationsplatzhalter bzw. deaktivierte Altansicht; hier ist kein fertiger Bearbeitungsablauf verfügbar.",
    "fields": [
      {
        "name": "Entwicklungsstand",
        "description": "Die angezeigte Einschränkung beachten; keine Buchung über diesen Platzhalter vornehmen."
      }
    ],
    "steps": [
      "Ansicht öffnen und Entwicklungsstand prüfen."
    ],
    "effect": "Diese Registerkarte ist ein Platzhalter. Vorhandene OP-Mengen- und Spannungsfunktionen befinden sich innerhalb der Stückzahlerfassung und sind davon getrennt zu prüfen.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Ausschuss und Abklärung in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "protocol",
    "title": "Fertigungsprotokoll",
    "tab": "produktion",
    "subtab": "protocol",
    "role": "leader",
    "status": "implemented_not_deployed",
    "purpose": "Vorhandene Produktionshistorie nachvollziehen.",
    "fields": [
      {
        "name": "Auftrag",
        "description": "Auftrag im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Zeitraum",
        "description": "Zeitraum im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Sortierung",
        "description": "Sortierung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Protokoll öffnen",
      "Auftrag wählen",
      "Ereignisse nachvollziehen"
    ],
    "effect": "Historie ist keine neue QS-Messwertmaske. Export nur im berechtigten Testkontext.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Fertigungsprotokoll in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "production-settings",
    "title": "Produktionseinstellungen",
    "tab": "produktion",
    "subtab": "settings",
    "role": "admin",
    "status": "visible_placeholder",
    "purpose": "Vorhandener Navigationsplatzhalter bzw. deaktivierte Altansicht; hier ist kein fertiger Bearbeitungsablauf verfügbar.",
    "fields": [
      {
        "name": "Entwicklungsstand",
        "description": "Die angezeigte Einschränkung beachten; keine Buchung über diesen Platzhalter vornehmen."
      }
    ],
    "steps": [
      "Ansicht öffnen und Entwicklungsstand prüfen."
    ],
    "effect": "Diese Ansicht aktiviert keine neue Schichtplanung oder Speicherung.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Produktionseinstellungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "articles",
    "title": "Artikelstamm",
    "tab": "produktion",
    "subtab": "articles",
    "role": "leader",
    "status": "implemented_not_deployed",
    "purpose": "Artikelversionen und Operationsvorlagen verwalten.",
    "fields": [
      {
        "name": "Artikelnummer",
        "description": "Artikelnummer im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Bezeichnung",
        "description": "Bezeichnung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Kunde",
        "description": "Kunde im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Zeichnung",
        "description": "Zeichnung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Revision",
        "description": "Revision im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Material",
        "description": "Material im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Technische Hinweise",
        "description": "Technische Hinweise im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Artikelstamm öffnen",
      "Artikelnummer suchen",
      "Entwurf bearbeiten und fachlich freigeben"
    ],
    "effect": "Freigegebene Inhalte werden versioniert; BA-Menge gehört nicht in den Stamm. Produktive Schemafreigabe fehlt.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Artikelstamm in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "readiness",
    "title": "OP und Maschinenbereitschaft",
    "tab": "produktion",
    "subtab": "readiness",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Physische Materialfreigabe von Planung unterscheiden.",
    "fields": [
      {
        "name": "Maschine",
        "description": "Maschine im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "OP",
        "description": "OP im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Menge",
        "description": "Menge im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Übergabegrund",
        "description": "Übergabegrund im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Maschinenübersicht öffnen",
      "Geplante und tatsächlich bereite OP unterscheiden",
      "Erlaubte Übernahme bestätigen"
    ],
    "effect": "Prognose ist keine Materialfreigabe. Der synthetische Test ersetzt keine native Parallelitätsabnahme.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "OP und Maschinenbereitschaft in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "tools",
    "title": "Werkzeugliste",
    "tab": "werkzeugverwaltung",
    "subtab": "list",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Werkzeuge suchen und technische Stammdaten ansehen.",
    "fields": [
      {
        "name": "Suche",
        "description": "Suche im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Hersteller",
        "description": "Hersteller im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Typ",
        "description": "Typ im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Halter",
        "description": "Halter im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Durchmesser",
        "description": "Durchmesser im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Werkzeugverwaltung öffnen",
      "Suche oder Filter setzen",
      "Werkzeugdetails prüfen"
    ],
    "effect": "Suche verändert keinen Bestand. Geometrische Werkzeugsuche ist hier nicht neu implementiert.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Werkzeugliste in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "tool-stock",
    "title": "Werkzeugbestand",
    "tab": "werkzeugverwaltung",
    "subtab": "stock",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Vorhandene Werkzeugbestände nachsehen.",
    "fields": [
      {
        "name": "Werkzeug",
        "description": "Werkzeug im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Lagerort",
        "description": "Lagerort im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Bestand",
        "description": "Bestand im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Bestand öffnen",
      "Werkzeug und Lagerort prüfen"
    ],
    "effect": "Bestand und Journalbuchung gemeinsam prüfen; die Vorschau nutzt synthetischen Browserzustand.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Werkzeugbestand in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "tool-movement",
    "title": "Werkzeugbuchungen",
    "tab": "werkzeugverwaltung",
    "subtab": "movement",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Bestehende Entnahme und Einlagerung erfassen.",
    "fields": [
      {
        "name": "Werkzeug",
        "description": "Werkzeug im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Menge",
        "description": "Menge im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Lagerort",
        "description": "Lagerort im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Bemerkung",
        "description": "Bemerkung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Entnahme / Einlagerung öffnen",
      "Werkzeug und Menge prüfen",
      "Buchung bestätigen"
    ],
    "effect": "In der Vorschau nur flüchtige synthetische Daten; kein produktiver Bestandsnachweis.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Werkzeugbuchungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "tool-journal",
    "title": "Werkzeugjournal",
    "tab": "werkzeugverwaltung",
    "subtab": "journal",
    "role": "leader",
    "status": "implemented_not_deployed",
    "purpose": "Werkzeugbewegungen nachvollziehen.",
    "fields": [
      {
        "name": "Zeitraum",
        "description": "Zeitraum im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Werkzeug",
        "description": "Werkzeug im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Aktion",
        "description": "Aktion im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Journal öffnen",
      "Filter setzen",
      "Bewegung kontrollieren"
    ],
    "effect": "Lesen gibt keine Korrektur- oder Löschrechte. Produktionshistorie niemals überschreiben.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Werkzeugjournal in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "tool-qr",
    "title": "Werkzeug QR Etiketten",
    "tab": "werkzeugverwaltung",
    "subtab": "qrLabels",
    "role": "leader",
    "status": "implemented_not_deployed",
    "purpose": "Vorhandene Werkzeugkennungen für Scanner bereitstellen.",
    "fields": [
      {
        "name": "Werkzeug",
        "description": "Werkzeug im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "QR-Kennung",
        "description": "QR-Kennung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "QR-Etiketten öffnen",
      "Werkzeug wählen",
      "Code prüfen und drucken"
    ],
    "effect": "Nur synthetische Codes im Test. Kein Zugriff auf Kamera oder Live-Inventar für Dokumentationsscreenshots.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Werkzeug QR Etiketten in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "storage",
    "title": "Lagerverwaltung",
    "tab": "lagerverwaltung",
    "subtab": "overview",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Lagerorte und bestehende Such-/Umlagerungsabläufe ansehen.",
    "fields": [
      {
        "name": "Lagerort",
        "description": "Lagerort im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Fach",
        "description": "Fach im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Werkzeug",
        "description": "Werkzeug im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Lagerverwaltung öffnen",
      "Fach suchen",
      "Nur erlaubte Aktion ausführen"
    ],
    "effect": "Lagerorte und tatsächlicher Werkzeugbestand bleiben getrennt nachvollziehbar.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Lagerverwaltung in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "scanner",
    "title": "Scanner",
    "tab": "scanner",
    "subtab": "toolScan",
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Bestehende QR-/Barcodeabläufe bedienen.",
    "fields": [
      {
        "name": "Code",
        "description": "Code im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Kamera",
        "description": "Kamera im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Aktion",
        "description": "Aktion im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Scanner öffnen",
      "Code prüfen",
      "Gefundenes Objekt und Buchungswirkung bestätigen"
    ],
    "effect": "Kameraberechtigung ist optional; ein Scan darf nicht unbemerkt doppelt buchen.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Scanner in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "analytics",
    "title": "Auswertungen",
    "tab": "auswertung",
    "subtab": "toolStats",
    "role": "leader",
    "status": "implemented_not_deployed",
    "purpose": "Bestehende Werkzeug- und Produktionsauswertungen ansehen.",
    "fields": [
      {
        "name": "Zeitraum",
        "description": "Zeitraum im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Abteilung",
        "description": "Abteilung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Export",
        "description": "Export im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Auswertung öffnen",
      "Filter prüfen",
      "Datenlage kontrollieren"
    ],
    "effect": "Unvollständige synthetische Daten erlauben keine Aussage über reale Effizienzgewinne.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Auswertungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "settings",
    "title": "System und Berechtigungen",
    "tab": "adminsystem",
    "subtab": "system",
    "role": "admin",
    "status": "visible_placeholder",
    "purpose": "Vorhandener Navigationsplatzhalter; dieser Bereich ist fachlich noch nicht implementiert.",
    "fields": [
      {
        "name": "Version",
        "description": "Version im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Einstellung",
        "description": "Einstellung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Benutzerrolle",
        "description": "Benutzerrolle im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Ansicht öffnen und ausgewiesenen Entwicklungsstand prüfen."
    ],
    "effect": "Kein fertig implementierter Speichervorgang. Keine neue Funktion als verfügbar ausgeben.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "System und Berechtigungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "learning",
    "title": "Gamification",
    "tab": "gamification",
    "subtab": null,
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Theoriefragen beantworten und eigenen Lernfortschritt verfolgen.",
    "fields": [
      {
        "name": "Antwortmöglichkeiten",
        "description": "Antwortmöglichkeiten im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Verschieben",
        "description": "Verschieben im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Kollegensichtbarkeit",
        "description": "Freiwillige Anzeige in der Abteilungsrangliste, jederzeit widerrufbar."
      }
    ],
    "steps": [
      "Gamification öffnen",
      "Alle zutreffenden Antworten wählen",
      "Antwort verbindlich abgeben und Erklärung lesen"
    ],
    "effect": "Kompetenz, Teilnahme und historische Erfahrung sind getrennt. Freiwillige Rangliste ist privat voreingestellt. Kalender-/30-Tage-Dienst fehlt.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js",
      "js/learning.js"
    ],
    "caption": "Gamification in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "safety",
    "title": "Sicherheitsunterweisungen",
    "tab": "safety-learning",
    "subtab": null,
    "role": "employee",
    "status": "implemented_not_deployed",
    "purpose": "Zugewiesene Theorieunterweisungen unabhängig vom Lernspiel bearbeiten.",
    "fields": [
      {
        "name": "Antworten",
        "description": "Antworten im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Nachweis",
        "description": "Nachweis im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Durchführungsbestätigung",
        "description": "Durchführungsbestätigung im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Sicherheitsunterweisungen öffnen",
      "Zugewiesene Version bearbeiten",
      "Erforderliche praktische Durchführung getrennt bestätigen lassen"
    ],
    "effect": "Ein Theoriequiz ersetzt keine praktische Unterweisung und keine Tätigkeitsfreigabe.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js",
      "js/learning.js"
    ],
    "caption": "Sicherheitsunterweisungen in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "quality",
    "title": "Digitale Qualitätsprüfung",
    "tab": null,
    "subtab": null,
    "role": "leader",
    "status": "planned",
    "purpose": "Geplante merkmalsbezogene Prüfintervalle, Toleranzanzeige und Freigaben.",
    "fields": [
      {
        "name": "Messmerkmal",
        "description": "Messmerkmal im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Messwert",
        "description": "Messwert im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Eingriffsgrenze",
        "description": "Eingriffsgrenze im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Teilenummer",
        "description": "Teilenummer im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Fachspezifikation aus PR #38 lesen"
    ],
    "effect": "Nur geplant. Keine fertige Messmaske oder automatische Gutteilfreigabe behaupten.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Digitale Qualitätsprüfung in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  },
  {
    "id": "documents",
    "title": "Dokumentenverwaltung",
    "tab": null,
    "subtab": null,
    "role": "admin",
    "status": "planned",
    "purpose": "Geplante versionierte Artikel-/BA-/OP-Dokumente.",
    "fields": [
      {
        "name": "Dokumenttyp",
        "description": "Dokumenttyp im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Version",
        "description": "Version im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      },
      {
        "name": "Artikelbezug",
        "description": "Artikelbezug im angezeigten fachlichen Kontext prüfen; keine Werte aus anderen Aufträgen übernehmen."
      }
    ],
    "steps": [
      "Konzept aus PR #37 prüfen"
    ],
    "effect": "Nur geplant. Kein öffentlicher Storage und kein Upload produktiver Dokumente in die Testinstanz.",
    "errors": [
      "Bei fehlender Berechtigung nicht über andere Menüs umgehen.",
      "Bei Speicherfehlern Eingaben erhalten und bestätigten Serverstand prüfen."
    ],
    "sources": [
      "schichtplan_mvp.js"
    ],
    "caption": "Dokumentenverwaltung in der tatsächlich laufenden Testanwendung mit synthetischen Daten."
  }
];if(typeof module!=="undefined")module.exports=topics;else root.HUMBEL_HELP_TOPICS=topics;})(typeof window!=="undefined"?window:globalThis);
