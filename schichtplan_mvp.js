const { createClient } = window.supabase;

const supabaseClient = createClient(
  window.SUPABASE_URL,
  window.SUPABASE_PUBLISHABLE_KEY,
);

let supabaseReady = false;
let currentSupabaseUser = null;
let currentEmployeeRecord = null;
let toolMaterials = [];

const USERS = [
  { name: "Lavdrim", slot: "A", type: "core" },
  { name: "Roger", slot: "B", type: "core" },
  { name: "Dashmir", slot: "C", type: "core" },
  { name: "Thomas", slot: "D", type: "springer" },
  { name: "Musa", slot: "E", type: "springer" },
  { name: "Ardian", slot: "F", type: "springer" },
];
const SLOT_CODES = ["A", "B", "C", "D", "E", "F"];
const DEFAULT_SLOT_ASSIGNMENTS = {
  A: "Lavdrim",
  B: "Roger",
  C: "Dashmir",
  D: "Thomas",
  E: "Musa",
  F: "Ardian",
};
const EMPLOYEE_COLOR_OPTIONS = [
  { key: "green", label: "Grün" },
  { key: "blue", label: "Blau" },
  { key: "yellow", label: "Gelb" },
  { key: "purple", label: "Lila" },
  { key: "orange", label: "Orange" },
  { key: "teal", label: "Türkis" },
  { key: "cyan", label: "Cyan" },
  { key: "pink", label: "Pink" },
  { key: "rose", label: "Rose" },
  { key: "lime", label: "Lime" },
  { key: "gray", label: "Grau" },
];
const DEFAULT_TOOL_LABELS = [
  "Schaftfräser",
  "Trochodialfräser",
  "Radiusfräser",
  "Kugelfräser",
  "Bohrer",
  "NC Anbohrer",
  "Gewindebohrer",
  "Gewindefräser",
  "Gewindeformer",
  "Gewindewirbler",
  "Ausdrehkopf",
];
const DEFAULT_TOOL_MANUFACTURERS = ["SixSigma", "SFS", "THAA"];
const DEFAULT_TOOL_HOLDERS = ["HSK 100", "HSK 63"];

const APP_VERSION = "0.5.05";
const INVENTORY_MODE_ENABLED = false;
const HUMBEL_COLORS = Object.freeze({
  primary: "#0d4682",
  primaryDark: "#08345f",
  accent: "#3fb498",
  background: "#f4f7fa",
  card: "#ffffff",
  text: "#102033",
  textMuted: "#5b6b80",
  border: "#d8e2ee",
});
const VERSION_LOG = [
  {
    version: "0.5.05",
    date: "2026-08-22 09:27",
    changes: ["Produktionsprotokoll ergänzt und Abschluss-Checkliste in Fertigmeldung verschoben."],
  },
  {
    version: "0.5.04",
    date: "2026-08-22 09:16",
    changes: ["Initialisierung der Abschluss-Checkliste stabilisiert."],
  },
  {
    version: "0.5.03",
    date: "2026-08-22 08:53",
    changes: ["Abschluss-Checkliste und Fertigmeldung für Produktionsaufträge ergänzt."],
  },
  {
    version: "0.5.02",
    date: "2026-08-22 08:42",
    changes: ["6M-Ursachenerfassung für Ausschuss und Abklärung ergänzt."],
  },
  {
    version: "0.5.01",
    date: "2026-08-22 08:31",
    changes: ["Speichern und Fehlerbehandlung für Produktionsspannungen stabilisiert."],
  },
  {
    version: "0.5.00",
    date: "2026-08-22 07:44",
    changes: ["Ausschuss und Abklärung je Spannung vorbereitet."],
  },
  {
    version: "0.4.99",
    date: "2026-08-22 07:30",
    changes: ["Gutteilezählung je Spannung und Mitarbeiter vorbereitet."],
  },
  {
    version: "0.4.98",
    date: "2026-08-22 07:15",
    changes: ["Mitarbeiterzuordnung je Produktionsauftrag vorbereitet."],
  },
  {
    version: "0.4.97",
    date: "2026-08-09 08:34",
    changes: ["Spannungen je Produktionsauftrag vorbereitet."],
  },
  {
    version: "0.4.96",
    date: "2026-08-08 06:57",
    changes: ["Produktionsvorschau pro Maschine und BA vorbereitet."],
  },
  {
    version: "0.4.95",
    date: "2026-08-08 06:41",
    changes: ["Zähler-Dashboard mit Maschinen- und Auftragsübersicht vorbereitet."],
  },
  {
    version: "0.4.94",
    date: "2026-08-02 08:37",
    changes: ["Aufträge / BA der Produktionszählung an Supabase angebunden."],
  },
  {
    version: "0.4.93",
    date: "2026-08-02 08:15",
    changes: ["Untertabs als eigene übersichtliche Modulansichten umgesetzt."],
  },
  {
    version: "0.4.92",
    date: "2026-08-02 09:06",
    changes: ["Humbel Corporate Design und Farbgrundlage angewendet."],
  },
  {
    version: "0.4.91",
    date: "2026-08-02 08:31",
    changes: ["Automatisierte Smoke-Tests mit Playwright vorbereitet."],
  },
  {
    version: "0.4.90",
    date: "2026-07-27 08:09",
    changes: [
      "Lagerfach-QR druckbar gemacht und Werkzeugbuchungen mit Menge und Journal korrigiert.",
    ],
  },
  {
    version: "0.4.89",
    date: "2026-07-27 07:02",
    changes: ["Lagerverwaltung-Untertabs korrekt angebunden."],
  },
  {
    version: "0.4.88",
    date: "2026-07-20 13:08",
    changes: ["Werkzeugverwaltung-Untertabs wieder korrekt angebunden."],
  },
  {
    version: "0.4.87",
    date: "2026-07-20 12:29",
    changes: ["Modulstruktur und Funktionslandkarte der App vorbereitet."],
  },
  {
    version: "0.4.86",
    date: "2026-07-20 12:13",
    changes: [
      "Stückzahl-Zähler im Produktionsbereich als Grundmodul ergänzt.",
    ],
  },
  {
    version: "0.4.85",
    date: "2026-07-16 11:25",
    changes: [
      "Abteilungsleiter-Ansicht für Produktion auf eigene Abteilungen begrenzt.",
    ],
  },
  {
    version: "0.4.84",
    date: "2026-06-29 21:19",
    changes: ["Maschinenverwaltung im Produktionsbereich ergänzt."],
  },
  {
    version: "0.4.83",
    date: "2026-06-29 12:30",
    changes: [
      "Systemuser-Schutz in der Mitarbeiterliste sichtbar und wirksam gemacht.",
    ],
  },
  {
    version: "0.4.82",
    date: "2026-06-29 12:21",
    changes: [
      "Admin-, Scanner- und Login-Systemuser in der Personalverwaltung geschützt.",
    ],
  },
  {
    version: "0.4.81",
    date: "2026-06-29 08:45",
    changes: [
      "Mitarbeiterverwaltung mit Abteilungs-Dropdown und Abteilungsleiter-Zuordnung erweitert.",
    ],
  },
  {
    version: "0.4.80",
    date: "2026-06-29 08:21",
    changes: [
      "Produktion-Übertab mit Abteilungsverwaltung vorbereitet.",
    ],
  },
  {
    version: "0.4.79",
    date: "2026-06-28 09:35",
    changes: [
      "Doppelte Schichtmodell-Anzeige bereinigt und Mitarbeiter-Löschung ergänzt.",
    ],
  },
  {
    version: "0.4.78",
    date: "2026-06-28 09:24",
    changes: [
      "Personalverwaltung an vorhandene employees-Tabelle angepasst und technische Supabase-Fehler im Bereich angezeigt.",
    ],
  },
  {
    version: "0.4.77",
    date: "2026-06-28 09:06",
    changes: [
      "Personalverwaltung für Mitarbeiter und Schichtmodell neu aufgebaut.",
    ],
  },
  {
    version: "0.4.76",
    date: "2026-06-27 06:54",
    changes: [
      "Inventurmodus vorübergehend deaktiviert, damit Design und Anordnung überarbeitet werden können.",
    ],
  },
  {
    version: "0.4.75",
    date: "2026-06-20 08:24",
    changes: ["Inventurhistorie mit Supabase-Sessions und Inventurergebnissen ergänzt"],
  },
  {
    version: "0.4.74",
    date: "2026-06-13 09:13",
    changes: ["Ersten Inventurmodus als reine Prüfansicht ergänzt"],
  },
  {
    version: "0.4.73",
    date: "2026-06-07 16:17",
    changes: [
      "Ausleihstatus aus Stammdaten entfernt und in Entnahme-/Einlagerungsworkflow verlagert",
    ],
  },
  {
    version: "0.4.72",
    date: "2026-06-07 10:21",
    changes: ["Werkzeug-Ausleihstatus mit Kostenträger/Maschinennummer ergänzt"],
  },
  {
    version: "0.4.71",
    date: "2026-06-07 07:49",
    changes: ["Eingabefeld für Ausspannlänge (AL) ergänzt"],
  },
  {
    version: "0.4.70",
    date: "2026-06-07 07:30",
    changes: ["Ausspannlänge (AL) in Werkzeugtabelle und Detailansicht ergänzt"],
  },
  {
    version: "0.4.68",
    date: "2026-06-07 07:17",
    changes: [
      "QR-Info im Lagerfach-Popup sichtbar platziert",
      "Echte QR-Code-Darstellung für Lagerfächer ergänzt",
    ],
  },
  {
    version: "0.4.67",
    date: "2026-05-17 09:30",
    changes: ["QR-System für Lagerfächer ergänzt"],
  },
  {
    version: "0.4.66",
    date: "2026-05-17 09:25",
    changes: ["Doppelte Fachbelegung beim Umlagern verhindert"],
  },
  {
    version: "0.4.65",
    date: "2026-05-17 09:18",
    changes: ["Werkzeug-Umlagerung aus Lagerfachansicht ergänzt"],
  },
  {
    version: "0.4.64",
    date: "2026-05-17 09:11",
    changes: ["Treffer-Blinken in der Lagerfachansicht deutlicher gemacht"],
  },
  {
    version: "0.4.63",
    date: "2026-05-17 09:03",
    changes: [
      "Blinkende Hervorhebung für Treffer-Fächer in der Lagerfachansicht ergänzt",
    ],
  },
  {
    version: "0.4.62",
    date: "2026-05-17 05:48",
    changes: ["Exakte T-Nummern-Suche in der Lagerfach-/Regalansicht ergänzt"],
  },
  {
    version: "0.4.61",
    date: "2026-05-17 05:28",
    changes: [
      "Werkzeugsuche mit Hervorhebung in der Lagerfach-/Regalansicht ergänzt",
    ],
  },
  {
    version: "0.4.60",
    date: "2026-05-16 08:33",
    changes: ["Erste Lagerfach-/Regalansicht ergänzt"],
  },
  {
    version: "0.4.59",
    date: "2026-05-16 08:24",
    changes: ["Realtime-Subscription für Werkzeugbereich nach Reload stabilisiert"],
  },
  {
    version: "0.4.58",
    date: "2026-05-16 08:16",
    changes: ["Supabase Realtime für Werkzeugbestand und Tool-Journal ergänzt"],
  },
  {
    version: "0.4.57",
    date: "2026-05-16 08:09",
    changes: ["Temporäre Debug-Logs beim Werkzeugladen entfernt"],
  },
  {
    version: "0.4.56",
    date: "2026-05-16 08:03",
    changes: ["Werkzeug-Ladeansicht nach erfolgreichem Tool-Load korrigiert"],
  },
  {
    version: "0.4.55",
    date: "2026-05-16 05:58",
    changes: ["Debug-Logs für Werkzeugladen aus Supabase ergänzt"],
  },
  {
    version: "0.4.54",
    date: "2026-05-15 22:22",
    changes: [
      "Leeres Rendern des Werkzeug-Tabs vor Supabase-Ladevorgang verhindert",
      "Werkzeugdaten werden nicht mehr aus localStorage wiederhergestellt",
      "Leere Supabase-Loads überschreiben vorhandene Werkzeugdaten nicht mehr",
    ],
  },
  {
    version: "0.4.53",
    date: "2026-05-15 20:36",
    changes: [
      "Werkzeug-Auto-Refresh bis nach initialem Supabase-Laden verzögert",
    ],
  },
  {
    version: "0.4.52",
    date: "2026-05-15 08:55",
    changes: ["Loading-/Render-Reihenfolge für Werkzeugdaten stabilisiert"],
  },
  {
    version: "0.4.51",
    date: "2026-05-15 08:46",
    changes: ["Zeitraumfilter für Werkzeugstatistik ergänzt"],
  },
  {
    version: "0.4.50",
    date: "2026-05-15 08:40",
    changes: ["Einfachen Bereich Werkzeugstatistik ergänzt"],
  },
  {
    version: "0.4.49",
    date: "2026-05-15 08:35",
    changes: ["Filter für Schichtjournal – Werkzeugwechsel ergänzt"],
  },
  {
    version: "0.4.48",
    date: "2026-05-15 04:44",
    changes: ["Werkzeugbereich auf volle Browserbreite erweitert"],
  },
  {
    version: "0.4.47",
    date: "2026-05-14",
    changes: [
      "Werkzeugbestand auf breiteres, lesbares Tabellenlayout zurückgestellt",
    ],
  },
  {
    version: "0.4.46",
    date: "2026-05-14",
    changes: ["Werkzeugbestand ohne horizontales Scrollen optimiert"],
  },
  {
    version: "0.4.45",
    date: "2026-05-14",
    changes: ["Werkzeughistorie pro Werkzeug ergänzt"],
  },
  {
    version: "0.4.44",
    date: "2026-05-14",
    changes: [
      "Werkzeug-Stammdaten in ein Popup verschoben",
      "Scrollposition der Werkzeugliste beim Auto-Refresh erhalten",
    ],
  },
  {
    version: "0.4.43",
    date: "2026-05-14",
    changes: [
      "Werkzeugbereich neu strukturiert: Werkzeugbestand oben, Filter integriert, Admin-Bestandsaktionen untereinander",
    ],
  },
  {
    version: "0.4.42",
    date: "2026-05-14",
    changes: ["Versionslog in der Fußzeile ergänzt"],
  },
  {
    version: "0.4.41",
    changes: [
      "Mindestbestand-Warnfarben unterschieden: erreicht = orange, unterschritten = rot",
    ],
  },
  {
    version: "0.4.40",
    changes: [
      "Mindestbestand-Warnung und Bereich Nachbestellen prüfen ergänzt",
    ],
  },
  {
    version: "0.4.39",
    changes: ["Auto-Aktualisierung im Werkzeugbereich alle 15 Sekunden ergänzt"],
  },
  {
    version: "0.4.38",
    changes: [
      "Werkzeug-Aktualisierung gegen leere Supabase-Rückgaben abgesichert",
    ],
  },
  {
    version: "0.4.37",
    changes: ["Button Werkzeuge aktualisieren ergänzt"],
  },
  {
    version: "0.4.35",
    changes: ["Temporäre Mobile-Debug-Konsole entfernt"],
  },
  {
    version: "0.4.34",
    changes: ["Tool-Journal-Anzeige stabil aus state.toolJournal gerendert"],
  },
  {
    version: "0.4.33",
    changes: ["Tool-Journal bleibt stabil aus Supabase geladen"],
  },
  {
    version: "0.4.31",
    changes: ["Tool-Journal-Anzeige auf Supabase-Daten vorbereitet"],
  },
  {
    version: "0.4.24",
    changes: ["Tool-Journal auf Supabase-Anbindung vorbereitet"],
  },
  {
    version: "0.4.21",
    changes: ["QR-Einlagerung mit Mengenbestätigung ergänzt"],
  },
  {
    version: "0.4.16",
    changes: ["QR-Entnahme mit direkter Bestätigung ergänzt"],
  },
  {
    version: "0.4.15",
    changes: ["Optionaler Plattenradius für Wendeplattenwerkzeuge ergänzt"],
  },
  {
    version: "0.4.14",
    changes: [
      "Bestellte Werkzeugkarten mit T-Nummer und dynamischen Details verbessert",
    ],
  },
  {
    version: "0.4.8",
    changes: ["Kamera-QR-Scanner vorbereitet"],
  },
  {
    version: "0.4.5",
    changes: ["Lokale QR-Code-Grafik für Werkzeugetiketten ergänzt"],
  },
  {
    version: "0.4.2",
    changes: ["Werkzeug-Scanner-Kiosk mit Rolle tool_scanner ergänzt"],
  },
  {
    version: "0.4.0",
    changes: ["Werkzeugbilder per Popup ergänzt"],
  },
];
const STORAGE_KEY = "schichtplan_mvp_v_0_2";
const STORAGE_LETTERS = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
];
const state = loadState();
let currentUser = null;
let currentTab = "schichtplan";
let statsViewPeriod = "week";
let qrScannerStream = null;
let qrScannerTimer = null;
let qrScannerMode = null;
let activeInventorySessionId = null;
let activeInventoryLocationKey = "";
let toolAutoRefreshTimer = null;
let toolPageLoadInProgress = false;
let toolRealtimeChannel = null;

const HELP_TEXTS = {
  planningPersonal: {
    title: "Planung / Personal",
    body: [
      "Hier pflegst du Mitarbeiter, Rollen, Farben und die Slots A-F.",
      "Die Slot-Zuordnung steuert, wer in der automatischen Rotation eingesetzt wird.",
      "Nach dem Speichern werden die Mitarbeiterdaten in Supabase aktualisiert und der Plan nutzt die neue Zuordnung.",
    ],
  },
  abstinenz: {
    title: "Abstinenz",
    body: [
      "Hier trägst du Urlaub, Krankheit und manuelle Abwesenheiten ein.",
      "Abwesenheiten öffnen Schichten, wenn der ursprünglich geplante Mitarbeiter fehlt.",
      "Nach dem Speichern werden die Planungsdaten in Supabase gehalten und beim Neuladen wieder geladen.",
    ],
  },
  ersatzplanung: {
    title: "Ersatzplanung",
    body: [
      "Hier planst du Ersatz oder Ausfall für Schichten während Urlaub oder Krankheit.",
      "Es werden nur Schichten des abwesenden Mitarbeiters im betroffenen Zeitraum angezeigt.",
      "Ersatz schreibt eine manuelle Einteilung, Ausfall sperrt die Schicht. Beides wirkt direkt auf den Plan.",
    ],
  },
  wochenende: {
    title: "Wochenendeinsätze",
    body: [
      "Springer melden, ob sie für Wochenend-Schichten können oder nicht können.",
      "Der Admin teilt Springer nur ein, wenn der ursprüngliche A/B/C-Mitarbeiter abwesend ist.",
      "Die Einteilung wird gespeichert und bleibt nach einem Refresh erhalten.",
    ],
  },
  schichttausch: {
    title: "Schichttausch",
    body: [
      "Hier können A/B/C-Mitarbeiter für einen Zeitraum ihre Rotation tauschen.",
      "Der Tausch wirkt ab dem Startdatum bis zum Enddatum oder bis er zurückgesetzt wird.",
      "Gespeicherte Tausche verändern die automatisch berechnete Schichtzuordnung.",
    ],
  },
  werkzeuge: {
    title: "Werkzeuge",
    body: [
      "Hier verwaltest du Werkzeugbestand, Stammdaten, Bestellungen und Werkzeugwechsel.",
      "Bestand und Werkzeugdaten werden in Supabase gespeichert.",
      "Wechsel und Bestellhistorie sind aktuell noch lokal und dienen als MVP-Journal.",
    ],
  },
  statistik: {
    title: "Statistik",
    body: [
      "Die Statistik vergleicht geplante Stunden, Ist-Stunden und Stillstand.",
      "Du kannst zwischen Woche, Monat und Jahr wechseln.",
      "Die Werte helfen, Abweichungen im Plan und Maschinenlaufzeit sichtbar zu machen.",
    ],
  },
};

function makeWeek(early, late, night, satPrimary, satSecondary) {
  return {
    mondayToFriday: [
      { label: "Früh", start: "05:00", end: "11:00", options: [early] },
      { label: "Spät", start: "13:00", end: "19:00", options: [late] },
      { label: "Nacht", start: "21:00", end: "03:00", options: [night] },
    ],
    saturday: [
      {
        label: "Samstag Morgen",
        start: "05:00",
        end: "11:00",
        options: [satPrimary],
      },
      {
        label: "Samstag Abend",
        start: "16:00",
        end: "22:00",
        options: [satSecondary, "D", "E", "F"],
      },
    ],
    sunday: [
      {
        label: "Sonntag Morgen",
        start: "06:00",
        end: "12:00",
        options: [late],
      },
      {
        label: "Sonntag Nacht",
        start: "18:00",
        end: "24:00",
        options: [night],
      },
    ],
  };
}

const WEEK_TEMPLATES = [
  makeWeek("A", "B", "C", "A", "B"),
  makeWeek("B", "C", "A", "B", "C"),
  makeWeek("C", "A", "B", "C", "A"),
  makeWeek("A", "B", "C", "A", "B"),
  makeWeek("B", "C", "A", "B", "C"),
  makeWeek("C", "A", "B", "C", "A"),
];

const ROTATION_ANCHOR_MONDAY = "2026-01-05";

const PLANNING_SUBTABS = [
  { id: "personal", label: "Personal" },
  { id: "abstinenz", label: "Abstinenz" },
  { id: "wochenende", label: "Wochenendeinsätze" },
  { id: "schichttausch", label: "Schichttausch" },
];
const DASHBOARD_SUBTABS = [
  { id: "overview", label: "Übersicht" },
  { id: "myShifts", label: "Meine Schichten" },
  { id: "todo", label: "To-Do" },
  { id: "oldSchedule", label: "Alt-Schichtplan" },
];
const PERSONNEL_MANAGEMENT_SUBTABS = [
  { id: "employees", label: "Mitarbeiter" },
  { id: "shiftModel", label: "Schichtmodell" },
  { id: "shiftPlanning", label: "Schichtplanung" },
  { id: "settings", label: "Einstellungen" },
];
const PRODUCTION_SUBTABS = [
  { id: "departments", label: "Abteilungen" },
  { id: "machines", label: "Maschinen" },
  { id: "orders", label: "Aufträge / BA" },
  { id: "counts", label: "Stückzahl" },
  { id: "setups", label: "Spannungen" },
  { id: "scrap", label: "Ausschuss / Abklärung" },
  { id: "protocol", label: "Protokoll" },
  { id: "settings", label: "Einstellungen" },
];
const TOOL_MANAGEMENT_SUBTABS = [
  { id: "list", label: "Werkzeugliste" },
  { id: "stock", label: "Bestand" },
  { id: "movement", label: "Entnahme / Einlagerung" },
  { id: "reorder", label: "Nachbestellen" },
  { id: "journal", label: "Journal" },
  { id: "qrLabels", label: "QR-Etiketten" },
  { id: "settings", label: "Einstellungen" },
];
const STORAGE_MANAGEMENT_SUBTABS = [
  { id: "overview", label: "Lagerübersicht" },
  { id: "search", label: "Fachsuche" },
  { id: "move", label: "Umlagern" },
  { id: "qr", label: "Lagerfach-QR" },
  { id: "inventory", label: "Inventur" },
  { id: "settings", label: "Einstellungen" },
];
const SCANNER_SUBTABS = [
  { id: "toolScan", label: "Werkzeug scannen" },
  { id: "storageScan", label: "Lagerfach scannen" },
  { id: "withdraw", label: "Entnahme" },
  { id: "restock", label: "Einlagerung" },
  { id: "return", label: "Rückgabe" },
  { id: "countScan", label: "Stückzahl-Scan" },
];
const ANALYTICS_SUBTABS = [
  { id: "toolStats", label: "Werkzeugstatistik" },
  { id: "orderStats", label: "Bestellstatistik" },
  { id: "productionStats", label: "Produktionszahlen" },
  { id: "employeesDepartments", label: "Mitarbeiter / Abteilung" },
  { id: "protocols", label: "Protokolle" },
  { id: "export", label: "Export" },
];
const ADMIN_SYSTEM_SUBTABS = [
  { id: "roles", label: "Rollen" },
  { id: "accounts", label: "Benutzer / Login-Konten" },
  { id: "company", label: "Firmen-Einstellungen" },
  { id: "version", label: "Version / Logs" },
  { id: "backup", label: "Backup" },
  { id: "setup", label: "Setup-Assistent" },
];
const VALID_EMPLOYEE_ROLES = [
  "admin",
  "employee",
  "tool_scanner",
  "department_admin",
];
const SHIFT_DEFINITION_TYPES = [
  {
    shift_key: "early",
    name: "Frühschicht",
    requires_time: true,
  },
  {
    shift_key: "late",
    name: "Spätschicht",
    requires_time: true,
  },
  {
    shift_key: "night",
    name: "Nachtschicht",
    requires_time: true,
  },
  {
    shift_key: "day",
    name: "Tagschicht",
    requires_time: false,
  },
];

function injectHumbelDesignStyles() {
  if (document.getElementById("humbelCorporateDesignStyles")) return;
  const style = document.createElement("style");
  style.id = "humbelCorporateDesignStyles";
  style.textContent = `
    :root {
      --humbel-blue: ${HUMBEL_COLORS.primary};
      --humbel-blue-dark: ${HUMBEL_COLORS.primaryDark};
      --humbel-accent: ${HUMBEL_COLORS.accent};
      --humbel-bg: ${HUMBEL_COLORS.background};
      --humbel-card: ${HUMBEL_COLORS.card};
      --humbel-text: ${HUMBEL_COLORS.text};
      --humbel-muted: ${HUMBEL_COLORS.textMuted};
      --humbel-border: ${HUMBEL_COLORS.border};
    }

    body {
      background: var(--humbel-bg) !important;
      color: var(--humbel-text) !important;
    }

    header.bg-white,
    #loginBox,
    #view > .bg-white,
    .bg-white.rounded-xl.shadow,
    .bg-white.rounded-xl.shadow-xl {
      background: var(--humbel-card) !important;
      border: 1px solid var(--humbel-border) !important;
      box-shadow: 0 12px 32px rgba(13, 70, 130, 0.08) !important;
    }

    h1,
    h2,
    h3,
    .text-slate-900 {
      color: var(--humbel-text) !important;
    }

    .text-slate-500,
    .text-slate-600,
    .text-slate-700 {
      color: var(--humbel-muted) !important;
    }

    input,
    select,
    textarea {
      border-color: var(--humbel-border) !important;
      color: var(--humbel-text) !important;
      background: #ffffff !important;
    }

    input:focus,
    select:focus,
    textarea:focus {
      outline: 2px solid rgba(63, 180, 152, 0.28) !important;
      border-color: var(--humbel-accent) !important;
    }

    .border,
    .border-slate-100,
    .border-slate-200,
    .border-slate-300,
    .border-slate-400 {
      border-color: var(--humbel-border) !important;
    }

    .bg-slate-50,
    .bg-slate-100 {
      background-color: #f7fafc !important;
    }

    thead.bg-slate-100,
    thead.bg-slate-200,
    .sticky.bg-slate-100,
    .sticky.bg-slate-200 {
      background-color: #e8f0f7 !important;
      color: var(--humbel-blue-dark) !important;
    }

    table tbody tr:hover {
      background-color: rgba(13, 70, 130, 0.04);
    }

    .humbel-main-nav {
      background: rgba(255, 255, 255, 0.78);
      border: 1px solid var(--humbel-border);
      border-radius: 12px;
      padding: 8px;
      box-shadow: 0 10px 26px rgba(13, 70, 130, 0.06);
    }

    .humbel-main-tab,
    .humbel-subtab {
      background: #ffffff !important;
      color: var(--humbel-blue-dark) !important;
      border-color: var(--humbel-border) !important;
      font-weight: 600;
      box-shadow: 0 1px 2px rgba(16, 32, 51, 0.04);
    }

    .humbel-main-tab:hover,
    .humbel-subtab:hover {
      background: #eef5fb !important;
      border-color: rgba(13, 70, 130, 0.34) !important;
    }

    .humbel-main-tab-active,
    .humbel-subtab-active {
      background: var(--humbel-blue) !important;
      color: #ffffff !important;
      border-color: var(--humbel-blue) !important;
      box-shadow: 0 8px 20px rgba(13, 70, 130, 0.22);
    }

    .humbel-tab-attention {
      border-color: #fb7185 !important;
    }

    button.bg-slate-900,
    button.bg-blue-600,
    button.bg-blue-700 {
      background-color: var(--humbel-blue) !important;
      color: #ffffff !important;
    }

    button.bg-slate-700,
    button.bg-slate-800 {
      background-color: var(--humbel-blue-dark) !important;
      color: #ffffff !important;
    }

    button.bg-emerald-600,
    button.bg-emerald-700,
    button.bg-green-600,
    button.bg-green-700 {
      background-color: var(--humbel-accent) !important;
      color: #ffffff !important;
    }

    button.bg-slate-200,
    button.bg-slate-100 {
      background-color: #eef3f8 !important;
      color: var(--humbel-blue-dark) !important;
      border: 1px solid var(--humbel-border);
    }

    button.bg-red-100,
    button.bg-red-700,
    button.bg-red-800,
    button.bg-rose-700 {
      background-color: #b42318 !important;
      color: #ffffff;
    }

    .bg-emerald-50 {
      background-color: rgba(63, 180, 152, 0.12) !important;
    }

    .bg-emerald-100 {
      background-color: rgba(63, 180, 152, 0.18) !important;
      color: #0f6f5e !important;
    }

    .bg-blue-50,
    .bg-sky-50 {
      background-color: rgba(13, 70, 130, 0.08) !important;
    }

    .text-blue-800,
    .text-sky-800 {
      color: var(--humbel-blue-dark) !important;
    }

    .text-emerald-700,
    .text-emerald-800 {
      color: #0f6f5e !important;
    }
  `;
  document.head.appendChild(style);
}

function setLoginStatus(message, isError = false) {
  const el = document.getElementById("loginStatus");
  if (!el) return;
  el.textContent = message;
  el.className = isError
    ? "text-sm rounded border border-rose-200 bg-rose-50 p-3 text-rose-700"
    : "text-sm rounded border border-slate-200 bg-slate-50 p-3 text-slate-600";
}

function fillLogin(type) {
  const emailInput = document.getElementById("loginEmail");
  const passwordInput = document.getElementById("loginPassword");
  if (!emailInput || !passwordInput) return;

  if (type === "admin" && window.TEST_ADMIN_EMAIL) {
    emailInput.value = window.TEST_ADMIN_EMAIL;
  }
  if (type === "employee" && window.TEST_EMPLOYEE_EMAIL) {
    emailInput.value = window.TEST_EMPLOYEE_EMAIL;
  }

  passwordInput.focus();
}

async function testSupabaseConnection() {
  try {
    const { error } = await supabaseClient
      .from("employees")
      .select("id")
      .limit(1);

    if (error) {
      console.error("Supabase-Test fehlgeschlagen:", error);
      setLoginStatus(`Supabase-Fehler: ${error.message}`, true);
      return false;
    }

    console.log("Supabase-Verbindung ok.");
    setLoginStatus("Supabase-Verbindung ok. Bitte anmelden.");
    return true;
  } catch (err) {
    console.error("Supabase konnte nicht initialisiert werden:", err);
    setLoginStatus("Supabase konnte nicht initialisiert werden.", true);
    return false;
  }
}

async function loadEmployeesFromSupabase() {
  const { data, error } = await supabaseClient
    .from("employees")
    .select("*")
    .order("created_at", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von employees:", error);
    state.ui = state.ui || {};
    state.ui.personalManagementEmployeesError =
      formatPersonalManagementSupabaseError(
        error,
        "Mitarbeiter konnten nicht geladen werden",
      );
    return null;
  }

  state.ui = state.ui || {};
  state.ui.personalManagementEmployeesError = "";
  return data || [];
}

function normalizeEmployeeFromDb(row) {
  const fallbackNameParts = String(row.display_name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  const fallbackFirstName = fallbackNameParts.shift() || row.display_name || "";
  const fallbackLastName = fallbackNameParts.join(" ");
  const firstName = row.first_name || fallbackFirstName || "";
  const lastName = row.last_name || fallbackLastName || "";
  const displayName =
    row.display_name ||
    [firstName, lastName].filter(Boolean).join(" ").trim() ||
    row.personnel_no ||
    "";
  const isActive = row.is_active !== false;
  const role = String(row.role || "employee").trim().toLowerCase();
  const authUserId = row.auth_user_id || "";

  return {
    id: row.id,
    authUserId,
    auth_user_id: authUserId,
    first_name: firstName,
    last_name: lastName,
    personnel_no: row.personnel_no || "",
    department: row.department || "",
    department_id: row.department_id || "",
    name: displayName,
    display_name: displayName,
    role,
    type: row.employee_type || "springer",
    employee_type: row.employee_type || "springer",
    slot: row.slot_code || "",
    slot_code: row.slot_code || "",
    isActive,
    is_active: isActive,
    color_key: row.color_key || "",
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyEmployeesToState(rows) {
  if (!Array.isArray(rows)) return;

  const employees = rows
    .map(normalizeEmployeeFromDb)
    .sort((a, b) =>
      `${a.last_name} ${a.first_name} ${a.display_name}`.localeCompare(
        `${b.last_name} ${b.first_name} ${b.display_name}`,
        "de",
      ),
    );
  const employeeMap = {};
  const nextSlotAssignments = { ...(state.slotAssignments || {}) };

  employees.forEach((employee) => {
    employeeMap[employee.id] = employee;
    if (
      employee.isActive &&
      employee.slot &&
      SLOT_CODES.includes(employee.slot)
    ) {
      nextSlotAssignments[employee.slot] = employee.name;
    }
  });

  SLOT_CODES.forEach((slot) => {
    const assignedName = nextSlotAssignments[slot];
    const assignedEmployee = employees.find((employee) => {
      return employee.name === assignedName && employee.isActive;
    });
    if (!assignedEmployee) {
      const defaultName = DEFAULT_SLOT_ASSIGNMENTS[slot] || "";
      const defaultEmployee = employees.find((employee) => {
        return employee.name === defaultName && employee.isActive;
      });
      nextSlotAssignments[slot] = defaultEmployee ? defaultName : "";
    }
  });

  state.employees = employeeMap;
  state.employeesList = employees;
  state.slotAssignments = nextSlotAssignments;
}

async function loadShiftDefinitionsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("shift_definitions")
    .select("*")
    .order("shift_key", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von shift_definitions:", error);
    state.ui.personalManagementShiftDefinitionsError =
      formatPersonalManagementSupabaseError(
        error,
        "Schichtdefinitionen konnten nicht geladen werden",
      );
    return null;
  }

  state.ui.personalManagementShiftDefinitionsError = "";
  return data || [];
}

function normalizeShiftDefinitionFromDb(row) {
  const template = SHIFT_DEFINITION_TYPES.find(
    (entry) => entry.shift_key === row.shift_key,
  );
  return {
    id: row.id,
    name: row.name || template?.name || row.shift_key || "",
    shift_key: row.shift_key || "",
    start_time: row.start_time || "",
    end_time: row.end_time || "",
    requires_time:
      row.requires_time !== undefined
        ? row.requires_time !== false
        : template?.requires_time !== false,
    active: row.active !== false,
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyShiftDefinitionsToState(rows) {
  if (!Array.isArray(rows)) return;
  state.shiftDefinitions = rows.map(normalizeShiftDefinitionFromDb);
}

async function loadDepartmentsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("departments")
    .select("*")
    .order("name", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von departments:", error);
    state.ui.productionDepartmentsError = formatProductionSupabaseError(
      error,
      "Abteilungen konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionDepartmentsError = "";
  return data || [];
}

function normalizeDepartmentFromDb(row) {
  return {
    id: row.id,
    name: row.name || "",
    code: row.code || "",
    leader_employee_id: row.leader_employee_id || "",
    active: row.active !== false,
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyDepartmentsToState(rows) {
  if (!Array.isArray(rows)) return;
  state.departments = rows.map(normalizeDepartmentFromDb);
}

async function loadProductionMachinesFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_machines")
    .select("*")
    .order("name", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_machines:", error);
    state.ui.productionMachinesError = formatProductionSupabaseError(
      error,
      "Maschinen konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionMachinesError = "";
  return data || [];
}

function normalizeProductionMachineFromDb(row) {
  return {
    id: row.id,
    name: row.name || "",
    machine_code: row.machine_code || "",
    department_id: row.department_id || "",
    active: row.active !== false,
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionMachinesToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionMachines = rows.map(normalizeProductionMachineFromDb);
}

async function loadProductionCountsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_counts")
    .select("*")
    .order("updated_at", { ascending: false });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_counts:", error);
    state.ui.productionCountsError = formatProductionSupabaseError(
      error,
      "Stückzahlen konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionCountsError = "";
  return data || [];
}

async function loadProductionOrdersFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_orders")
    .select("*")
    .order("updated_at", { ascending: false });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_orders:", error);
    state.ui.productionOrdersError = formatProductionSupabaseError(
      error,
      "Aufträge / BA konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionOrdersError = "";
  return data || [];
}

async function loadProductionOrderStationsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_order_stations")
    .select("*")
    .order("station_no", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_order_stations:", error);
    state.ui.productionStationsError = formatProductionSupabaseError(
      error,
      "Spannungen konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionStationsError = "";
  return data || [];
}

async function loadProductionOrderEmployeesFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_order_employees")
    .select("*")
    .order("sort_order", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_order_employees:", error);
    state.ui.productionOrderEmployeesError = formatProductionSupabaseError(
      error,
      "Auftragsmitarbeiter konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionOrderEmployeesError = "";
  return data || [];
}

async function loadProductionStationCountsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_station_counts")
    .select("*")
    .order("updated_at", { ascending: false });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_station_counts:", error);
    state.ui.productionStationCountsError = formatProductionSupabaseError(
      error,
      "Gutteil-Zähler konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionStationCountsError = "";
  return data || [];
}

async function loadProductionStationEventsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_station_events")
    .select("*")
    .order("created_at", { ascending: false });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_station_events:", error);
    state.ui.productionStationEventsError = formatProductionSupabaseError(
      error,
      "Produktionsereignisse konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionStationEventsError = "";
  return data || [];
}

async function loadProductionOrderHistoryFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_order_history")
    .select("*")
    .order("created_at", { ascending: false });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_order_history:", error);
    state.ui.productionOrderHistoryError = formatProductionSupabaseError(
      error,
      "Auftragshistorie konnte nicht geladen werden",
    );
    return null;
  }

  state.ui.productionOrderHistoryError = "";
  return data || [];
}

async function loadProductionQaCausesFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_qa_causes")
    .select("*")
    .eq("active", true)
    .order("group_label", { ascending: true })
    .order("sort_order", { ascending: true })
    .order("reason_label", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_qa_causes:", error);
    state.ui.productionQaCausesError = formatProductionSupabaseError(
      error,
      "6M-Ursachen konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionQaCausesError = "";
  return data || [];
}

async function loadProductionChecklistTemplatesFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_checklist_templates")
    .select("*")
    .eq("active", true)
    .order("sort_order", { ascending: true })
    .order("item_label", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    console.error("Fehler beim Laden von production_checklist_templates:", error);
    state.ui.productionChecklistTemplatesError = formatProductionSupabaseError(
      error,
      "Abschluss-Checklisten-Vorlagen konnten nicht geladen werden",
    );
    return null;
  }

  state.ui.productionChecklistTemplatesError = "";
  return data || [];
}

async function loadProductionOrderChecklistFromSupabase() {
  const { data, error } = await supabaseClient
    .from("production_order_checklist")
    .select("*")
    .order("sort_order", { ascending: true })
    .order("item_label", { ascending: true });

  state.ui = state.ui || {};

  if (error) {
    logProductionOrderChecklistSupabaseError("select", error);
    state.ui.productionOrderChecklistError = formatProductionSupabaseError(
      error,
      "Abschluss-Checkliste konnte nicht geladen werden",
    );
    return null;
  }

  state.ui.productionOrderChecklistError = "";
  return data || [];
}

function normalizeProductionOrderFromDb(row) {
  return {
    id: row.id,
    machine_id: row.machine_id || "",
    department_id: row.department_id || "",
    ba_number: row.ba_number || "",
    article_number: row.article_number || "",
    ba_quantity: Math.max(0, Number(row.ba_quantity || 0)),
    target_quantity: Math.max(0, Number(row.target_quantity || 0)),
    pallet_count: Math.max(0, Number(row.pallet_count || 0)),
    pieces_per_pallet: Math.max(0, Number(row.pieces_per_pallet || 0)),
    use_chain_logic: row.use_chain_logic === true,
    status: row.status || "running",
    started_at: row.started_at || null,
    completed_at: row.completed_at || null,
    created_by_employee_id: row.created_by_employee_id || "",
    updated_by_employee_id: row.updated_by_employee_id || "",
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionOrdersToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionOrders = rows.map(normalizeProductionOrderFromDb);
}

function normalizeProductionOrderStationFromDb(row) {
  const stationNo = Math.max(1, Math.trunc(Number(row.station_no || 1)));
  const actualMinutes =
    row.actual_time_minutes === null || row.actual_time_minutes === undefined || row.actual_time_minutes === ""
      ? null
      : Math.max(0, Math.trunc(Number(row.actual_time_minutes || 0)));
  return {
    id: row.id,
    order_id: row.order_id || "",
    station_no: stationNo,
    name: row.name || `Spannung ${stationNo}`,
    lock_name: row.lock_name === true,
    op_number: row.op_number || "",
    time_status: row.time_status === "changed" ? "changed" : "ok",
    actual_time_minutes: actualMinutes,
    scrap_total: Math.max(0, Number(row.scrap_total || 0)),
    clarify_total: Math.max(0, Number(row.clarify_total || 0)),
    scrap_lifetime: Math.max(0, Number(row.scrap_lifetime || 0)),
    clarify_lifetime: Math.max(0, Number(row.clarify_lifetime || 0)),
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionOrderStationsToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionOrderStations = rows
    .map(normalizeProductionOrderStationFromDb)
    .sort((a, b) => a.station_no - b.station_no);
}

function normalizeProductionOrderEmployeeFromDb(row) {
  return {
    id: row.id,
    order_id: row.order_id || "",
    employee_id: row.employee_id || "",
    employee_name: row.employee_name || "",
    personnel_no: row.personnel_no || "",
    role: row.role || "worker",
    sort_order: Math.max(0, Math.trunc(Number(row.sort_order || 0))),
    active: row.active !== false,
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionOrderEmployeesToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionOrderEmployees = rows
    .map(normalizeProductionOrderEmployeeFromDb)
    .sort((a, b) => a.sort_order - b.sort_order);
}

function normalizeProductionStationCountFromDb(row) {
  return {
    id: row.id,
    order_id: row.order_id || "",
    station_id: row.station_id || "",
    order_employee_id: row.order_employee_id || "",
    good_qty: Math.max(0, Math.trunc(Number(row.good_qty || 0))),
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionStationCountsToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionStationCounts = rows.map(normalizeProductionStationCountFromDb);
}

function normalizeProductionStationEventFromDb(row) {
  return {
    id: row.id,
    order_id: row.order_id || "",
    station_id: row.station_id || "",
    order_employee_id: row.order_employee_id || "",
    employee_id: row.employee_id || "",
    event_type: row.event_type || "",
    qty: Number(row.qty || 0),
    qa_cause_id: row.qa_cause_id || "",
    note: row.note || "",
    created_at: row.created_at || null,
  };
}

function applyProductionStationEventsToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionStationEvents = rows.map(normalizeProductionStationEventFromDb);
}

function normalizeProductionOrderHistoryFromDb(row) {
  return {
    id: row.id,
    order_id: row.order_id || "",
    station_id: row.station_id || "",
    employee_id: row.employee_id || "",
    history_type: row.history_type || "",
    qty: Number(row.qty || 0),
    payload: row.payload || null,
    note: row.note || "",
    created_at: row.created_at || null,
  };
}

function applyProductionOrderHistoryToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionOrderHistory = rows.map(normalizeProductionOrderHistoryFromDb);
}

function normalizeProductionQaCauseFromDb(row) {
  return {
    id: row.id,
    group_id: row.group_id || "",
    group_label: row.group_label || "Ohne Gruppe",
    reason_id: row.reason_id || "",
    reason_label: row.reason_label || "",
    active: row.active !== false,
    sort_order: Math.max(0, Math.trunc(Number(row.sort_order || 0))),
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionQaCausesToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionQaCauses = rows
    .map(normalizeProductionQaCauseFromDb)
    .filter((cause) => cause.active && cause.reason_label)
    .sort((a, b) =>
      `${a.group_label} ${String(a.sort_order).padStart(6, "0")} ${a.reason_label}`.localeCompare(
        `${b.group_label} ${String(b.sort_order).padStart(6, "0")} ${b.reason_label}`,
        "de",
      ),
    );
}

function normalizeProductionChecklistTemplateFromDb(row) {
  return {
    id: row.id,
    item_key: row.item_key || "",
    item_label: row.item_label || "",
    active: row.active !== false,
    sort_order: Math.max(0, Math.trunc(Number(row.sort_order || 0))),
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionChecklistTemplatesToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionChecklistTemplates = rows
    .map(normalizeProductionChecklistTemplateFromDb)
    .filter((item) => item.active && item.item_key && item.item_label)
    .sort((a, b) =>
      `${String(a.sort_order).padStart(6, "0")} ${a.item_label}`.localeCompare(
        `${String(b.sort_order).padStart(6, "0")} ${b.item_label}`,
        "de",
      ),
    );
}

function normalizeProductionOrderChecklistFromDb(row) {
  return {
    id: row.id,
    order_id: row.order_id || "",
    item_key: row.item_key || "",
    item_label: row.item_label || "",
    checked: row.checked === true,
    checked_at: row.checked_at || null,
    checked_by_employee_id: row.checked_by_employee_id || "",
    sort_order: Math.max(0, Math.trunc(Number(row.sort_order || 0))),
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyProductionOrderChecklistToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionOrderChecklist = rows
    .map(normalizeProductionOrderChecklistFromDb)
    .filter((item) => item.order_id && item.item_key)
    .sort((a, b) =>
      `${String(a.sort_order).padStart(6, "0")} ${a.item_label}`.localeCompare(
        `${String(b.sort_order).padStart(6, "0")} ${b.item_label}`,
        "de",
      ),
    );
}

function normalizeProductionCountFromDb(row) {
  return {
    id: row.id,
    machine_id: row.machine_id || "",
    department_id: row.department_id || "",
    employee_id: row.employee_id || "",
    order_no: row.order_no || "",
    article_no: row.article_no || "",
    good_qty: Math.max(0, Number(row.good_qty || 0)),
    scrap_qty: Math.max(0, Number(row.scrap_qty || 0)),
    note: row.note || "",
    status: row.status || "running",
    started_at: row.started_at || null,
    updated_at: row.updated_at || null,
    completed_at: row.completed_at || null,
  };
}

function applyProductionCountsToState(rows) {
  if (!Array.isArray(rows)) return;
  state.productionCounts = rows.map(normalizeProductionCountFromDb);
}

function normalizeToolFromDb(row) {
  return {
    id: row.id,
    tNumber: row.t_number,
    label: row.label,
    diameter: row.diameter,
    overhangLength: row.overhang_length || "",
    threadPrefix: row.thread_prefix || "",
    threadPitch: row.thread_pitch || "",
    cornerRadius: row.corner_radius || "",
    materialId: row.material_id || null,
    shelf: row.shelf,
    articleNo: row.article_no,
    holder: row.holder,
    stock: Number(row.stock || 0),
    minStock: Number(row.min_stock || 0),
    optimalStock: Number(row.optimal_stock || 0),
    manufacturer: row.manufacturer || "",
    ordered: !!row.ordered,
    orderedQty: Number(row.ordered_qty || 0),
    insertTool: !!row.insert_tool,
    insertEdges: Number(row.insert_edges || 0),
    insertRadius: row.insert_radius || "",
    isBorrowed: row.is_borrowed === true,
    borrowedTo: row.borrowed_to || "",
    borrowedAt: row.borrowed_at || "",
  };
}

function normalizeTaskFromDb(row) {
  return {
    id: row.id,
    title: row.title || "",
    description: row.description || "",
    assignee: row.assigned_to || "",
    assignedTo: row.assigned_to || "",
    dueDate: row.due_date || "",
    status: row.status || "open",
    createdAt: row.created_at || null,
    completedAt: row.completed_at || null,
  };
}

async function loadTasksFromSupabase() {
  const { data, error } = await supabaseClient
    .from("planner_tasks")
    .select("*")
    .order("created_at", { ascending: false });

  if (error) {
    console.error("Fehler beim Laden von planner_tasks:", error);
    return null;
  }

  console.log("planner_tasks raw data:", data);
  return (data || []).map(normalizeTaskFromDb);
}

async function loadToolsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("tools")
    .select("*")
    .order("t_number", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von tools:", error);
    return [];
  }

  return (data || []).map(normalizeToolFromDb);
}

async function reloadSingleToolFromSupabase(toolId) {
  const { data, error } = await supabaseClient
    .from("tools")
    .select("*")
    .eq("id", toolId)
    .maybeSingle();

  if (error || !data) {
    console.warn("Werkzeug konnte nicht einzeln geladen werden", {
      toolId,
      error,
    });
    return null;
  }

  const normalized = normalizeToolFromDb(data);
  state.tools = state.tools.map((t) => (t.id === toolId ? normalized : t));
  return normalized;
}

function normalizeToolJournalFromDb(row) {
  return {
    id: row.id,
    toolId: row.tool_id,
    toolTNumber: row.tool_t_number,
    toolLabel: row.tool_label,
    action: row.action,
    qty: row.qty,
    stockBefore: row.stock_before,
    stockAfter: row.stock_after,
    user: row.user_name,
    createdAt: row.created_at,
  };
}

async function loadToolJournalFromSupabase() {
  const { data, error } = await supabaseClient
    .from("tool_journal")
    .select("*")
    .order("created_at", { ascending: false })
    .limit(100);

  if (error) {
    console.warn("Tool-Journal konnte nicht geladen werden", error);
    return;
  }

  state.toolJournal = (data || []).map(normalizeToolJournalFromDb);
  console.log(
    "Tool-Journal aus Supabase geladen:",
    state.toolJournal.length,
  );
  return state.toolJournal;
}

async function addToolJournalEntry(entry) {
  const localEntry = {
    id: `journal-${Date.now()}`,
    toolId: entry.toolId || null,
    toolTNumber: entry.toolTNumber || entry.tNumber || "",
    toolLabel: entry.toolLabel || entry.label || "",
    action: entry.action,
    qty: Number(entry.qty || 0),
    stockBefore: Number(entry.stockBefore || 0),
    stockAfter: Number(entry.stockAfter || 0),
    user: entry.user || currentUser?.name || "System",
    createdAt: new Date().toISOString(),
  };
  localEntry.at = localEntry.createdAt.slice(0, 16).replace("T", " ");
  localEntry.tNumber = localEntry.toolTNumber;

  const payload = {
    tool_id: entry.toolId || null,
    tool_t_number: entry.toolTNumber || entry.tNumber || null,
    tool_label: entry.toolLabel || entry.label || null,
    action: entry.action,
    qty: Number(entry.qty || 0),
    stock_before: Number(entry.stockBefore || 0),
    stock_after: Number(entry.stockAfter || 0),
    user_name: entry.user || currentUser?.name || "System",
  };

  const { error } = await supabaseClient
    .from("tool_journal")
    .insert(payload);

  if (error) {
    console.warn("Tool-Journal konnte nicht in Supabase gespeichert werden", error);
    state.toolJournal.unshift(localEntry);
    return localEntry;
  }

  state.toolJournal.unshift(localEntry);
  return localEntry;
}

async function refreshToolsAndRender() {
  const tools = await loadToolsFromSupabase();
  if (Array.isArray(tools) && tools.length > 0) {
    state.tools = tools;
  }
  await loadToolJournalFromSupabase();
  persist();
  render();
}

async function loadToolPageData() {
  state.ui = state.ui || {};
  const previousTools = Array.isArray(state.tools) ? [...state.tools] : [];

  state.ui.toolsLoading = true;
  render();

  try {
    const tools = await loadToolsFromSupabase();
    if (Array.isArray(tools) && tools.length > 0) {
      state.tools = tools;
    } else if (previousTools.length > 0) {
      state.tools = previousTools;
      console.warn(
        "Werkzeugdaten wurden nicht geleert, vorherige Daten bleiben erhalten.",
      );
    } else {
      console.warn(
        "Werkzeugdaten wurden noch nicht geladen, leerer Supabase-Load wird ignoriert.",
      );
    }

    await loadToolJournalFromSupabase();
    state.ui.toolsInitialLoaded = true;
  } finally {
    state.ui.toolsLoading = false;
    render();
  }
}

async function forceLoadToolPageData() {
  if (toolPageLoadInProgress) return;
  toolPageLoadInProgress = true;

  try {
    const tools = await loadToolsFromSupabase();
    if (Array.isArray(tools) && tools.length > 0) {
      state.tools = tools;
    } else {
      console.warn(
        "Force-Load hat keine Tools geliefert, state.tools bleibt unverändert.",
      );
    }
    await loadToolJournalFromSupabase();
    state.ui = state.ui || {};
    state.ui.toolsInitialLoaded = true;
  } catch (err) {
    console.error("Werkzeugdaten konnten nicht geladen werden", err);
  } finally {
    toolPageLoadInProgress = false;
    render();
  }
}

async function refreshToolPageData() {
  if (!state.ui?.supabaseReady) {
    console.warn("Werkzeug-Refresh übersprungen: Supabase noch nicht bereit.");
    return;
  }

  if (!state.ui?.toolsInitialLoaded) {
    console.warn(
      "Werkzeug-Refresh übersprungen: Werkzeugdaten noch nicht initial geladen.",
    );
    return;
  }

  const stockScrollEl = document.getElementById("toolStockTableScroll");
  const stockScrollTop = stockScrollEl ? stockScrollEl.scrollTop : 0;
  const previousTools = Array.isArray(state.tools) ? [...state.tools] : [];
  const previousJournal = Array.isArray(state.toolJournal)
    ? [...state.toolJournal]
    : [];

  const tools = await loadToolsFromSupabase();
  if (Array.isArray(tools) && tools.length > 0) {
    state.tools = tools;
  } else if (previousTools.length > 0) {
    state.tools = previousTools;
    console.warn(
      "Werkzeug-Refresh hat keine Tools geladen, vorherige Daten bleiben erhalten.",
    );
  }

  await loadToolJournalFromSupabase();
  if (
    (!Array.isArray(state.toolJournal) || state.toolJournal.length === 0) &&
    previousJournal.length > 0
  ) {
    state.toolJournal = previousJournal;
    console.warn(
      "Journal-Refresh hat keine Einträge geladen, vorherige Daten bleiben erhalten.",
    );
  }

  persist();
  render();

  const newStockScrollEl = document.getElementById("toolStockTableScroll");
  if (newStockScrollEl) {
    newStockScrollEl.scrollTop = stockScrollTop;
  }
}

function startToolAutoRefresh() {
  stopToolAutoRefresh();

  if (currentTab !== "werkzeuge") return;
  if (toolRealtimeChannel) return;
  if (!state.ui?.supabaseReady || !state.ui?.toolsInitialLoaded) return;
  if (!Array.isArray(state.tools) || state.tools.length === 0) return;

  toolAutoRefreshTimer = setInterval(async () => {
    if (currentTab !== "werkzeuge") {
      stopToolAutoRefresh();
      return;
    }

    await refreshToolPageData();
  }, 15000);
}

function stopToolAutoRefresh() {
  if (toolAutoRefreshTimer) {
    clearInterval(toolAutoRefreshTimer);
    toolAutoRefreshTimer = null;
  }
}

async function startToolRealtimeSubscription() {
  if (!supabaseClient) return;
  if (toolRealtimeChannel) return;

  const { data } = await supabaseClient.auth.getSession();
  if (!data?.session) {
    console.warn("Tool Realtime nicht gestartet: keine Supabase-Session");
    return;
  }

  toolRealtimeChannel = supabaseClient
    .channel("tool-realtime-updates")
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "tools" },
      async () => {
        if (currentTab === "werkzeuge") {
          await refreshToolPageData();
        }
      },
    )
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "tool_journal" },
      async () => {
        if (currentTab === "werkzeuge") {
          await refreshToolPageData();
        }
      },
    )
    .subscribe((status) => {
      console.log("Tool Realtime Status:", status);
      if (status === "CLOSED" || status === "CHANNEL_ERROR") {
        toolRealtimeChannel = null;
        startToolAutoRefresh();
      }
    });
}

function stopToolRealtimeSubscription() {
  if (toolRealtimeChannel && supabaseClient) {
    supabaseClient.removeChannel(toolRealtimeChannel);
    toolRealtimeChannel = null;
  }
}

async function getCurrentEmployeeRecord() {
  const {
    data: { user },
    error: userError,
  } = await supabaseClient.auth.getUser();

  if (userError) {
    console.error("Fehler bei auth.getUser():", userError);
    return null;
  }

  if (!user) return null;

  const { data, error } = await supabaseClient
    .from("employees")
    .select("*")
    .eq("auth_user_id", user.id)
    .eq("is_active", true)
    .maybeSingle();

  if (error) {
    console.error("Fehler beim Laden des employees-Datensatzes:", error);
    return null;
  }

  if (data && !VALID_EMPLOYEE_ROLES.includes(data.role)) {
    console.error("Ungültige employees-Rolle beim Login:", data.role);
    return null;
  }

  return data || null;
}

async function loadToolMaterialsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("tool_materials")
    .select("*")
    .eq("is_active", true)
    .order("sort_order", { ascending: true })
    .order("name", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von tool_materials:", error);
    return [];
  }

  return data || [];
}

function getToolMaterialNameById(materialId) {
  if (!materialId) return "-";
  const found = toolMaterials.find((m) => m.id === materialId);
  return found?.name || "-";
}

async function loadPlanningDataFromSupabase() {
  const [
    assignmentsRes,
    absencesRes,
    vacationsRes,
    sickLeavesRes,
    availabilityRes,
    swapsRes,
    cancellationsRes,
    saturdayRequestsRes,
    replacementsRes,
  ] = await Promise.all([
    supabaseClient
      .from("planner_assignments")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_absences")
      .select("*")
      .order("absence_date", { ascending: true }),
    supabaseClient
      .from("planner_vacations")
      .select("*")
      .order("from_date", { ascending: true }),
    supabaseClient
      .from("planner_sick_leaves")
      .select("*")
      .order("from_date", { ascending: true }),
    supabaseClient
      .from("planner_availability")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_swaps")
      .select("*")
      .order("start_date", { ascending: true }),
    supabaseClient
      .from("planner_shift_cancellations")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_saturday_requests")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_absence_replacements")
      .select("*")
      .order("shift_date", { ascending: true }),
  ]);

  const responses = [
    ["planner_assignments", assignmentsRes],
    ["planner_absences", absencesRes],
    ["planner_vacations", vacationsRes],
    ["planner_sick_leaves", sickLeavesRes],
    ["planner_availability", availabilityRes],
    ["planner_swaps", swapsRes],
    ["planner_shift_cancellations", cancellationsRes],
    ["planner_saturday_requests", saturdayRequestsRes],
    ["planner_absence_replacements", replacementsRes],
  ];

  const firstError = responses.find(([, res]) => res.error)?.[1]?.error;
  if (firstError) {
    console.error("Fehler beim Laden der Planungsdaten:", firstError);
    return null;
  }

  const tasks = await loadTasksFromSupabase();

  const assignments = {};
  (assignmentsRes.data || []).forEach((row) => {
    assignments[row.shift_id] = row.assigned_user;
  });

  const absences = {};
  (absencesRes.data || []).forEach((row) => {
    absences[row.absence_key] = true;
  });

  const vacations = (vacationsRes.data || []).map((row) => ({
    id: row.id,
    user: row.user_name,
    from: row.from_date,
    to: row.to_date,
  }));

  const sickLeaves = (sickLeavesRes.data || []).map((row) => ({
    id: row.id,
    user: row.user_name,
    from: row.from_date,
    to: row.to_date,
  }));

  const availability = {};
  (availabilityRes.data || []).forEach((row) => {
    availability[row.availability_key] = row.status;
  });

  const swaps = (swapsRes.data || []).map((row) => ({
    id: row.id,
    userA: row.user_a,
    userB: row.user_b,
    startDate: row.start_date,
    endDate: row.end_date || null,
  }));

  const shiftCancellations = {};
  (cancellationsRes.data || []).forEach((row) => {
    shiftCancellations[row.shift_id] = true;
  });

  const saturdayEveningRequests = {};
  (saturdayRequestsRes.data || []).forEach((row) => {
    saturdayEveningRequests[row.request_key] = true;
  });

  const { data: specialDaysData } = await supabaseClient
    .from("planner_special_days")
    .select("*");

  state.specialDays = {};
  (specialDaysData || []).forEach((d) => {
    state.specialDays[d.day_date] = d;
  });

  const absenceReplacements = {};
  (replacementsRes.data || []).forEach((row) => {
    absenceReplacements[row.shift_id] = {
      sourceType: row.source_type,
      sourceId: row.source_id,
      absentUser: row.absent_user,
      from: row.from_date,
      to: row.to_date,
      mode: row.mode,
      replacementUser: row.replacement_user,
      weekFrom: row.week_from,
      weekTo: row.week_to,
    };
  });

  return {
    assignments,
    absences,
    vacations,
    sickLeaves,
    availability,
    swaps,
    shiftCancellations,
    saturdayEveningRequests,
    absenceReplacements,
    tasks,
  };
}

async function loginWithSupabase() {
  const email = document.getElementById("loginEmail")?.value?.trim() || "";
  const password = document.getElementById("loginPassword")?.value || "";

  if (!email || !password) {
    setLoginStatus("Bitte E-Mail und Passwort eingeben.", true);
    return;
  }

  setLoginStatus("Anmeldung läuft ...");

  const { error } = await supabaseClient.auth.signInWithPassword({
    email,
    password,
  });

  if (error) {
    console.error("Supabase-Login fehlgeschlagen:", error);
    setLoginStatus(`Login fehlgeschlagen: ${error.message}`, true);
    return;
  }

  await syncSupabaseSessionToApp();
}

async function logoutSupabase() {
  await supabaseClient.auth.signOut();
  stopToolAutoRefresh();
  stopToolRealtimeSubscription();
  state.ui = state.ui || {};
  state.ui.supabaseReady = false;
  state.ui.toolsInitialLoaded = false;
  currentSupabaseUser = null;
  currentEmployeeRecord = null;
  currentUser = null;
  currentTab = "schichtplan";

  document.getElementById("loginBox")?.classList.remove("hidden");
  document.getElementById("tabs")?.classList.add("hidden");
  document.getElementById("sessionInfo").textContent = "";
  document.getElementById("view").innerHTML = "";
  setLoginStatus("Supabase-Session gelöscht.");
}

async function syncSupabaseSessionToApp() {
  stopToolAutoRefresh();
  state.ui = state.ui || {};
  state.ui.supabaseReady = false;
  state.ui.toolsInitialLoaded = false;

  const {
    data: { user },
    error,
  } = await supabaseClient.auth.getUser();

  if (error) {
    console.error("Fehler beim Abrufen des angemeldeten Users:", error);
    setLoginStatus(`Fehler beim Session-Lesen: ${error.message}`, true);
    return;
  }

  if (!user) {
    setLoginStatus("Keine aktive Supabase-Session vorhanden.");
    return;
  }

  currentSupabaseUser = user;
  currentEmployeeRecord = await getCurrentEmployeeRecord();

  if (!currentEmployeeRecord) {
    setLoginStatus(
      "Login ok, aber kein aktiver employees-Eintrag gefunden.",
      true,
    );
    return;
  }

  currentUser = {
    name: currentEmployeeRecord.display_name,
    role: currentEmployeeRecord.role,
  };

  const employees = await loadEmployeesFromSupabase();
  const shiftDefinitions = await loadShiftDefinitionsFromSupabase();
  const departments = await loadDepartmentsFromSupabase();
  const productionMachines = await loadProductionMachinesFromSupabase();
  const productionCounts = await loadProductionCountsFromSupabase();
  const productionOrders = await loadProductionOrdersFromSupabase();
  const productionOrderStations = await loadProductionOrderStationsFromSupabase();
  const productionOrderEmployees = await loadProductionOrderEmployeesFromSupabase();
  const productionStationCounts = await loadProductionStationCountsFromSupabase();
  const productionStationEvents = await loadProductionStationEventsFromSupabase();
  const productionOrderHistory = await loadProductionOrderHistoryFromSupabase();
  const productionQaCauses = await loadProductionQaCausesFromSupabase();
  const productionChecklistTemplates = await loadProductionChecklistTemplatesFromSupabase();
  const productionOrderChecklist = await loadProductionOrderChecklistFromSupabase();
  const materials = await loadToolMaterialsFromSupabase();
  const planning = await loadPlanningDataFromSupabase();

  applyEmployeesToState(employees);
  applyShiftDefinitionsToState(shiftDefinitions);
  applyDepartmentsToState(departments);
  applyProductionMachinesToState(productionMachines);
  applyProductionCountsToState(productionCounts);
  applyProductionOrdersToState(productionOrders);
  applyProductionOrderStationsToState(productionOrderStations);
  applyProductionOrderEmployeesToState(productionOrderEmployees);
  applyProductionStationCountsToState(productionStationCounts);
  applyProductionStationEventsToState(productionStationEvents);
  applyProductionOrderHistoryToState(productionOrderHistory);
  applyProductionQaCausesToState(productionQaCauses);
  applyProductionChecklistTemplatesToState(productionChecklistTemplates);
  applyProductionOrderChecklistToState(productionOrderChecklist);
  toolMaterials = materials;
  await loadToolPageData();
  state.ui.supabaseReady = true;
  state.ui.toolsInitialLoaded = true;

  if (planning) {
    state.assignments = planning.assignments;
    state.absences = planning.absences;
    state.vacations = planning.vacations;
    state.sickLeaves = planning.sickLeaves;
    state.availability = planning.availability;
    state.swaps = planning.swaps;
    state.shiftCancellations = planning.shiftCancellations;
    state.saturdayEveningRequests = planning.saturdayEveningRequests;
    state.absenceReplacements = planning.absenceReplacements;
    if (Array.isArray(planning.tasks)) {
      state.tasks = planning.tasks;
      console.log("Tasks aus Supabase geladen:", state.tasks);
    } else {
      console.warn(
        "Tasks wurden nicht aktualisiert, alter state.tasks bleibt erhalten.",
      );
    }
  }

  persist();

  document.getElementById("loginBox")?.classList.add("hidden");
  setLoginStatus(
    `Angemeldet als ${currentEmployeeRecord.display_name} (${currentEmployeeRecord.role}).`,
  );

  console.log("Supabase-User:", currentSupabaseUser);
  console.log("Employees-Datensatz:", currentEmployeeRecord);
  console.log("Schichtdefinitionen nach Login geladen:", shiftDefinitions);
  console.log("Abteilungen nach Login geladen:", departments);
  console.log("Produktionsmaschinen nach Login geladen:", productionMachines);
  console.log("Produktionsstückzahlen nach Login geladen:", productionCounts);
  console.log("Produktionsaufträge nach Login geladen:", productionOrders);
  console.log("Produktionsspannungen nach Login geladen:", productionOrderStations);
  console.log("Produktionsmitarbeiter nach Login geladen:", productionOrderEmployees);
  console.log("Produktions-Gutteilzähler nach Login geladen:", productionStationCounts);
  console.log("Produktionsereignisse nach Login geladen:", productionStationEvents);
  console.log("Produktions-Auftragshistorie nach Login geladen:", productionOrderHistory);
  console.log("Produktions-6M-Ursachen nach Login geladen:", productionQaCauses);
  console.log("Produktions-Checklisten-Vorlagen nach Login geladen:", productionChecklistTemplates);
  console.log("Produktions-Auftragschecklisten nach Login geladen:", productionOrderChecklist);
  console.log("Tool-Materials nach Login geladen:", materials);
  console.log("Tools nach Login geladen:", state.tools);
  console.log("Planungsdaten nach Login geladen:", planning);
  console.log("Tasks nach Session-Sync:", state.tasks);

  await startToolRealtimeSubscription();
  render();
}

async function bootSupabase() {
  stopToolAutoRefresh();
  state.ui = state.ui || {};
  state.ui.supabaseReady = false;
  state.ui.toolsInitialLoaded = false;

  supabaseReady = await testSupabaseConnection();

  if (!supabaseReady) {
    console.warn(
      "Supabase ist aktuell nicht bereit. App läuft vorerst lokal weiter.",
    );
    return;
  }

  const {
    data: { session },
  } = await supabaseClient.auth.getSession();

  if (session?.user) {
    await syncSupabaseSessionToApp();
    return;
  }

  const employees = await loadEmployeesFromSupabase();
  const shiftDefinitions = await loadShiftDefinitionsFromSupabase();
  const departments = await loadDepartmentsFromSupabase();
  const productionMachines = await loadProductionMachinesFromSupabase();
  const productionCounts = await loadProductionCountsFromSupabase();
  const productionOrders = await loadProductionOrdersFromSupabase();
  const productionOrderStations = await loadProductionOrderStationsFromSupabase();
  const productionOrderEmployees = await loadProductionOrderEmployeesFromSupabase();
  const productionStationCounts = await loadProductionStationCountsFromSupabase();
  const productionStationEvents = await loadProductionStationEventsFromSupabase();
  const productionOrderHistory = await loadProductionOrderHistoryFromSupabase();
  const productionQaCauses = await loadProductionQaCausesFromSupabase();
  const productionChecklistTemplates = await loadProductionChecklistTemplatesFromSupabase();
  const productionOrderChecklist = await loadProductionOrderChecklistFromSupabase();
  const materials = await loadToolMaterialsFromSupabase();
  const planning = await loadPlanningDataFromSupabase();

  applyEmployeesToState(employees);
  applyShiftDefinitionsToState(shiftDefinitions);
  applyDepartmentsToState(departments);
  applyProductionMachinesToState(productionMachines);
  applyProductionCountsToState(productionCounts);
  applyProductionOrdersToState(productionOrders);
  applyProductionOrderStationsToState(productionOrderStations);
  applyProductionOrderEmployeesToState(productionOrderEmployees);
  applyProductionStationCountsToState(productionStationCounts);
  applyProductionStationEventsToState(productionStationEvents);
  applyProductionOrderHistoryToState(productionOrderHistory);
  applyProductionQaCausesToState(productionQaCauses);
  applyProductionChecklistTemplatesToState(productionChecklistTemplates);
  applyProductionOrderChecklistToState(productionOrderChecklist);
  toolMaterials = materials;

  if (planning) {
    state.assignments = planning.assignments;
    state.absences = planning.absences;
    state.vacations = planning.vacations;
    state.sickLeaves = planning.sickLeaves;
    state.availability = planning.availability;
    state.swaps = planning.swaps;
    state.shiftCancellations = planning.shiftCancellations;
    state.saturdayEveningRequests = planning.saturdayEveningRequests;
    state.absenceReplacements = planning.absenceReplacements;
  }

  persist();

  console.log("Employees aus Supabase:", employees);
  console.log("Schichtdefinitionen aus Supabase:", shiftDefinitions);
  console.log("Abteilungen aus Supabase:", departments);
  console.log("Produktionsmaschinen aus Supabase:", productionMachines);
  console.log("Produktionsstückzahlen aus Supabase:", productionCounts);
  console.log("Produktionsaufträge aus Supabase:", productionOrders);
  console.log("Produktionsspannungen aus Supabase:", productionOrderStations);
  console.log("Produktionsmitarbeiter aus Supabase:", productionOrderEmployees);
  console.log("Produktions-Gutteilzähler aus Supabase:", productionStationCounts);
  console.log("Produktionsereignisse aus Supabase:", productionStationEvents);
  console.log("Produktions-Auftragshistorie aus Supabase:", productionOrderHistory);
  console.log("Produktions-6M-Ursachen aus Supabase:", productionQaCauses);
  console.log("Produktions-Checklisten-Vorlagen aus Supabase:", productionChecklistTemplates);
  console.log("Produktions-Auftragschecklisten aus Supabase:", productionOrderChecklist);
  console.log("Tool-Materials aus Supabase:", materials);
  console.log("Planungsdaten aus Supabase:", planning);

  if (currentUser) render();
}

function allUsers() {
  if (Array.isArray(state.employeesList) && state.employeesList.length) {
    return state.employeesList.map((employee) => ({
      id: employee.id,
      authUserId: employee.authUserId || employee.auth_user_id || "",
      auth_user_id: employee.auth_user_id || employee.authUserId || "",
      name: employee.name || employee.display_name,
      display_name: employee.display_name || employee.name,
      slot: employee.slot || employee.slot_code || "",
      type: employee.type || employee.employee_type || "springer",
      role: employee.role || "employee",
      color_key: employee.color_key || "",
      isActive: employee.isActive !== false && employee.is_active !== false,
    }));
  }

  return [...USERS, ...(state.extraUsers || [])];
}

function activeUsers() {
  return allUsers().filter((u) => {
    const activeByDb = u.isActive !== false;
    const activeByLocalFallback = !state.inactiveUsers?.[u.name];
    return activeByDb && activeByLocalFallback;
  });
}

function loadState() {
  const base = {
    absences: {},
    assignments: {},
    availability: {},
    checklists: {},
    unmanned: {},
    vacations: [],
    sickLeaves: [],
    swaps: [],
    shiftCancellations: {},
    slotAssignments: { ...DEFAULT_SLOT_ASSIGNMENTS },
    saturdayEveningRequests: {},
    tasks: [],
    conflicts: {},
    shiftEndChecks: {},
    shiftStartChecks: {},
    machineDowntime: {},
    machinePromptSeen: {},
    extraUsers: [],
    inactiveUsers: {},
    employees: {},
    employeesList: [],
    shiftDefinitions: [],
    departments: [],
    productionMachines: [],
    productionCounts: [],
    productionOrders: [],
    productionOrderStations: [],
    productionOrderEmployees: [],
    productionStationCounts: [],
    productionStationEvents: [],
    productionOrderHistory: [],
    productionQaCauses: [],
    productionChecklistTemplates: [],
    productionOrderChecklist: [],
    tools: [],
    toolLabelsExtra: [],
    toolManufacturersExtra: [],
    toolJournal: [],
    toolFilters: {
      search: "",
      label: "",
      tNumber: "",
      diameter: "",
      holder: "",
      imageStatus: "",
    },
    toolJournalFilters: {
      search: "",
      action: "",
      range: "",
    },
    toolStatisticsRange: "30d",
    storageSearch: "",
    toolOrderOverrides: {},
    orderArchive: [],
    orderHistory: [],
    orderStatsView: "week",
    orderSuggestionState: {},
    orderListPopupOpen: false,
    selectedOrderListManufacturer: "",
    dashboardSubTab: "overview",
    planningSubTab: "personal",
    personnelManagementSubTab: "employees",
    productionSubTab: "departments",
    toolManagementSubTab: "list",
    storageManagementSubTab: "overview",
    scannerSubTab: "toolScan",
    analyticsSubTab: "toolStats",
    adminSystemSubTab: "roles",
    absenceReplacements: {},
    replacementPlannerSelection: {},
    replacementPlannerChoice: {},
    ui: {
      pendingEmployeeEdits: {},
      pendingSlotAssignments: {},
      personalManagementActionError: "",
      personalManagementActionMessage: "",
      productionActionError: "",
      productionActionMessage: "",
      productionDepartmentsError: "",
      productionMachinesError: "",
      productionCountsError: "",
      productionOrdersError: "",
      productionStationsError: "",
      productionOrderEmployeesError: "",
      productionStationCountsError: "",
      productionStationEventsError: "",
      productionOrderHistoryError: "",
      productionQaCausesError: "",
      productionChecklistTemplatesError: "",
      productionOrderChecklistError: "",
      productionChecklistSavingKey: "",
      productionChecklistPreparingOrderId: "",
      productionChecklistAutoAttemptedOrderId: "",
      productionOrderCompletingId: "",
      productionGoodQtySavingKey: "",
      productionStationAmountSavingKey: "",
      productionCounterSelectedMachineId: "",
      productionCounterActiveOrderId: "",
      productionCompletionOrderId: "",
      productionProtocolSelectedOrderId: "",
      productionProtocolSort: "desc",
      toolsLoading: false,
      supabaseReady: false,
      toolsInitialLoaded: false,
    },
  };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return base;
    const parsed = JSON.parse(raw);
    delete parsed.departments;
    delete parsed.productionMachines;
    delete parsed.productionCounts;
    delete parsed.productionOrders;
    delete parsed.productionOrderStations;
    delete parsed.productionOrderEmployees;
    delete parsed.productionStationCounts;
    delete parsed.productionStationEvents;
    delete parsed.productionOrderHistory;
    delete parsed.productionQaCauses;
    delete parsed.productionChecklistTemplates;
    delete parsed.productionOrderChecklist;
    delete parsed.tools;
    delete parsed.toolJournal;
    return {
      ...base,
      ...parsed,
      tools: base.tools,
      toolJournal: base.toolJournal,
      toolFilters: {
        ...base.toolFilters,
        ...(parsed.toolFilters || {}),
      },
      toolJournalFilters: {
        ...base.toolJournalFilters,
        ...(parsed.toolJournalFilters || {}),
      },
      toolStatisticsRange: parsed.toolStatisticsRange || base.toolStatisticsRange,
      ui: {
        ...base.ui,
        ...(parsed.ui || {}),
      },
    };
  } catch {
    return base;
  }
}

function persist() {
  const snapshot = { ...state };
  delete snapshot.employees;
  delete snapshot.employeesList;
  delete snapshot.shiftDefinitions;
  delete snapshot.departments;
  delete snapshot.productionMachines;
  delete snapshot.productionCounts;
  delete snapshot.productionOrders;
  delete snapshot.productionOrderStations;
  delete snapshot.productionOrderEmployees;
  delete snapshot.productionStationCounts;
  delete snapshot.productionStationEvents;
  delete snapshot.productionOrderHistory;
  delete snapshot.productionQaCauses;
  delete snapshot.productionChecklistTemplates;
  delete snapshot.productionOrderChecklist;
  delete snapshot.tools;
  delete snapshot.toolJournal;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function helpButton(topic) {
  return `<button class="inline-flex items-center justify-center w-7 h-7 rounded-full bg-blue-600 text-white font-bold shadow-sm hover:bg-blue-700" onclick="openHelp('${topic}')" title="Hilfe">?</button>`;
}

function getHelpModalHost() {
  let host = document.getElementById("helpModalHost");
  if (!host) {
    host = document.createElement("div");
    host.id = "helpModalHost";
    document.body.appendChild(host);
  }
  return host;
}

function openHelp(topic) {
  const help = HELP_TEXTS[topic];
  if (!help) return;

  const body = help.body
    .map((paragraph) => `<p class="text-sm text-slate-700">${paragraph}</p>`)
    .join("");

  getHelpModalHost().innerHTML = `<div class="fixed inset-0 z-[60] bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-lg p-4">
      <div class="flex items-start justify-between gap-3 mb-3">
        <h3 class="text-lg font-bold">${help.title}</h3>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeHelp()">Schließen</button>
      </div>
      <div class="space-y-2">${body}</div>
    </div>
  </div>`;
}

function closeHelp() {
  getHelpModalHost().innerHTML = "";
}

function ensureVersionFooter() {
  let footer = document.getElementById("versionFooter");
  if (!footer) {
    footer = document.createElement("div");
    footer.id = "versionFooter";
    footer.className = "fixed bottom-2 right-2 z-40";
    document.body.appendChild(footer);
  }

  footer.innerHTML = `<button class="px-3 py-1 rounded-full bg-white/95 border border-slate-200 shadow text-xs text-slate-600 hover:bg-slate-50" onclick="openVersionLog()">Versionslog · v${APP_VERSION}</button>`;
}

function openVersionLog() {
  const rows = VERSION_LOG.map((entry) => {
    const changes = entry.changes
      .map((change) => `<li>${escapeHtml(change)}</li>`)
      .join("");
    const date = entry.date
      ? `<span class="text-xs text-slate-500">${escapeHtml(entry.date)}</span>`
      : "";

    return `<div class="border-b border-slate-100 pb-3">
      <div class="flex items-center justify-between gap-3 mb-1">
        <h4 class="font-semibold">Version ${escapeHtml(entry.version)}</h4>
        ${date}
      </div>
      <ul class="list-disc pl-5 text-sm text-slate-700 space-y-1">${changes}</ul>
    </div>`;
  }).join("");

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-2xl max-h-[88vh] overflow-auto p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">Versionslog</h3>
          <p class="text-sm text-slate-500">Aktuelle Version: v${APP_VERSION}</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeVersionLog()">Schließen</button>
      </div>
      <div class="space-y-3">${rows}</div>
    </div>
  </div>`;
}

function closeVersionLog() {
  getModalHost().innerHTML = "";
}

function loginAs(name) {
  if (name === "admin") {
    currentUser = { name: "Admin", role: "admin" };
    currentEmployeeRecord =
      (state.employeesList || []).find(
        (employee) => employee.role === "admin" && employee.is_active !== false,
      ) || currentEmployeeRecord;
  } else {
    currentUser = { name, role: "employee" };
    currentEmployeeRecord =
      (state.employeesList || []).find(
        (employee) =>
          employee.is_active !== false &&
          (employee.display_name === name || employee.name === name),
      ) || null;
  }
  document.getElementById("loginBox")?.classList.add("hidden");
  render();
}

function logout() {
  logoutSupabase();
}

function render() {
  if (!currentUser) return;
  cleanupOrderArchive();

  const tabs =
    currentUser.role === "tool_scanner" ? ["scanner"] : ["dashboard"];
  if (currentUser.role === "admin") {
    tabs.push(
      "personalverwaltung",
      "produktion",
      "werkzeugverwaltung",
      "lagerverwaltung",
      "scanner",
      "auswertung",
      "adminsystem",
    );
  }
  if (currentUser.role === "department_admin") tabs.push("produktion");
  if (currentUser.role === "employee") tabs.push("werkzeugverwaltung");

  const tabsEl = document.getElementById("tabs");
  tabsEl.className = "humbel-main-nav flex gap-2 flex-wrap";
  tabsEl.innerHTML =
    tabs
      .map((t) => {
        const statusClass = tabNeedsAttention(t) ? "humbel-tab-attention" : "";
        const activeClass =
          currentTab === t ? "humbel-main-tab-active" : "humbel-main-tab";
        return `<button class="px-3 py-2 rounded border ${activeClass} ${statusClass}" onclick="setTab('${t}')">${labelTab(t)}</button>`;
      })
      .join("") +
    `<button class="px-3 py-2 rounded bg-red-700 text-white" onclick="logout()">Abmelden</button>`;

  tabsEl.classList.remove("hidden");

  document.getElementById("sessionInfo").textContent =
    `Angemeldet: ${currentUser.name} (${currentUser.role})`;

  if (!tabs.includes(currentTab)) currentTab = tabs[0];
  if (["werkzeugverwaltung", "lagerverwaltung"].includes(currentTab)) {
    startToolRealtimeSubscription().catch(console.warn);
    startToolAutoRefresh();
  } else {
    stopToolAutoRefresh();
  }

  const view = document.getElementById("view");
  if (view?.parentElement) {
    view.parentElement.className =
      ["werkzeugverwaltung", "lagerverwaltung"].includes(currentTab)
        ? "w-full max-w-none p-4 space-y-4"
        : "max-w-7xl mx-auto p-4 space-y-4";
  }

  if (currentTab === "dashboard") view.innerHTML = renderDashboard();
  if (currentTab === "produktion") view.innerHTML = renderProduction();
  if (currentTab === "personalverwaltung")
    view.innerHTML = renderPersonalManagement();
  if (currentTab === "werkzeugverwaltung") view.innerHTML = renderToolManagement();
  if (currentTab === "lagerverwaltung") view.innerHTML = renderStorageManagement();
  if (currentTab === "scanner") view.innerHTML = renderScannerModule();
  if (currentTab === "auswertung") view.innerHTML = renderAnalyticsModule();
  if (currentTab === "adminsystem") view.innerHTML = renderAdminSystemModule();
  ensureVersionFooter();
  if (currentUser.role === "tool_scanner") return;
  maybeShowMachinePrompt();
  maybeTaskReminder();
  maybeShowShiftStartChecklist();
  maybeShowShiftEndChecklist();
}

function labelTab(tab) {
  return {
    dashboard: "Start / Dashboard",
    produktion: "Produktion",
    personalverwaltung: "Personalverwaltung",
    werkzeugverwaltung: "Werkzeugverwaltung",
    lagerverwaltung: "Lagerverwaltung",
    scanner: "Scanner",
    auswertung: "Auswertung",
    adminsystem: "Admin / System",
  }[tab];
}

function setTab(tab) {
  const legacyTabMap = {
    schichtplan: "dashboard",
    meine: "dashboard",
    werkzeuge: "werkzeugverwaltung",
    toolscanner: "scanner",
    bestellstatistik: "auswertung",
    statistik: "auswertung",
    planung: "personalverwaltung",
    todo: "dashboard",
    konflikte: "auswertung",
  };
  if (legacyTabMap[tab]) {
    currentTab = legacyTabMap[tab];
    render();
    return;
  }
  currentTab = tab;
  render();
}

function setStatsView(period) {
  statsViewPeriod = period;
  render();
}

function setPlanningSubTab(subTab) {
  if (!PLANNING_SUBTABS.some((t) => t.id === subTab)) return;
  state.planningSubTab = subTab;
  persist();
  render();
}

function setPersonnelManagementSubTab(subTab) {
  if (!PERSONNEL_MANAGEMENT_SUBTABS.some((t) => t.id === subTab)) return;
  state.personnelManagementSubTab = subTab;
  persist();
  render();
}

function setProductionSubTab(subTab) {
  if (!PRODUCTION_SUBTABS.some((t) => t.id === subTab)) return;
  state.productionSubTab = subTab;
  persist();
  render();
}

function setModuleSubTab(stateKey, subTab, allowedTabs) {
  if (!allowedTabs.some((t) => t.id === subTab)) return;
  state[stateKey] = subTab;
  persist();
  render();
}

function setToolManagementSubTab(subTab) {
  if (!TOOL_MANAGEMENT_SUBTABS.some((t) => t.id === subTab)) return;
  state.toolManagementSubTab = subTab;
  persist();
  render();
}

function setStorageManagementSubTab(subTab) {
  if (!STORAGE_MANAGEMENT_SUBTABS.some((t) => t.id === subTab)) return;
  state.storageManagementSubTab = subTab;
  persist();
  render();
}

function setScannerSubTab(subTab) {
  setModuleSubTab("scannerSubTab", subTab, SCANNER_SUBTABS);
}

function setAnalyticsSubTab(subTab) {
  setModuleSubTab("analyticsSubTab", subTab, ANALYTICS_SUBTABS);
}

function setAdminSystemSubTab(subTab) {
  setModuleSubTab("adminSystemSubTab", subTab, ADMIN_SYSTEM_SUBTABS);
}

function setDashboardSubTab(subTab) {
  setModuleSubTab("dashboardSubTab", subTab, DASHBOARD_SUBTABS);
}

function renderModuleSubTabs(tabs, activeId, setterName) {
  return tabs
    .map((tab) => {
      const active = activeId === tab.id;
      return `<button class='px-3 py-2 rounded border ${active ? "humbel-subtab-active" : "humbel-subtab"}' onclick="${setterName}('${tab.id}')">${tab.label}</button>`;
    })
    .join("");
}

function renderModulePlaceholder(title, text) {
  return `<div class='border rounded-lg p-4 bg-slate-50'>
    <h3 class='font-semibold mb-2'>${escapeHtml(title)}</h3>
    <p class='text-sm text-slate-600'>${escapeHtml(text)}</p>
  </div>`;
}

function renderDashboard() {
  const subTab = DASHBOARD_SUBTABS.some((tab) => tab.id === state.dashboardSubTab)
    ? state.dashboardSubTab
    : "overview";
  const cards = [
    ["Personalverwaltung", "Mitarbeiter, Schichtmodell und spätere Schichtplanung."],
    ["Produktion", "Abteilungen, Maschinen, Stückzahl und künftige Produktionsmodule."],
    ["Werkzeugverwaltung", "Werkzeugliste, Bestand, Buchungen, Nachbestellung und Journal."],
    ["Lagerverwaltung", "Lagerübersicht, Fachsuche, Umlagerung, QR und Inventur-Platz."],
    ["Scanner", "Scanner-Flows für Werkzeug, Lagerfach und spätere Stückzahl-Scans."],
    ["Auswertung", "Statistiken, Produktionszahlen, Protokolle und Export."],
    ["Admin / System", "Rollen, Login-Konten, Versionen, Backup und Setup."],
  ];
  let content = "";
  if (subTab === "overview") {
    content = `<div class='grid md:grid-cols-2 xl:grid-cols-3 gap-3'>
      ${cards.map(([title, text]) => `<div class='border rounded-lg p-3 bg-slate-50'><h3 class='font-semibold'>${escapeHtml(title)}</h3><p class='text-sm text-slate-600 mt-1'>${escapeHtml(text)}</p></div>`).join("")}
    </div>`;
  }
  if (subTab === "myShifts") {
    content =
      currentUser?.role === "employee"
        ? renderMyShifts()
        : renderModulePlaceholder(
            "Meine Schichten",
            "Meine Schichten ist für Mitarbeiter vorgesehen.",
          );
  }
  if (subTab === "todo") content = renderTodo();
  if (subTab === "oldSchedule") {
    content = renderModulePlaceholder(
      "Alt-Schichtplan",
      "Alter Schichtplan bleibt Alt / deaktiviert und wird nicht als neue Struktur weiterverwendet.",
    );
  }
  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-lg font-semibold'>Start / Dashboard</h2>
      <p class='text-sm text-slate-500 mt-1'>Modulstruktur der Humbel Schichtplan- und Werkzeug-App.</p>
    </div>
    <div class='flex gap-2 flex-wrap'>${renderModuleSubTabs(DASHBOARD_SUBTABS, subTab, "setDashboardSubTab")}</div>
    ${content}
    <div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Alter Schichtplan bleibt Alt / deaktiviert und wird nicht als neue Struktur weiterverwendet.</div>
  </div>`;
}

function renderToolManagement() {
  const subTab = TOOL_MANAGEMENT_SUBTABS.some(
    (tab) => tab.id === state.toolManagementSubTab,
  )
    ? state.toolManagementSubTab
    : "list";
  let content = "";
  if (subTab === "list") content = renderTools("list");
  if (subTab === "stock") {
    content = `${renderToolManagementSubtabHint("Bestand", "Vorhandene Bestands-, Lager- und Mindestbestandsansicht.")}
      ${renderTools("stock")}`;
  }
  if (subTab === "movement") content = renderToolScanner();
  if (subTab === "reorder") {
    content = `${renderToolManagementSubtabHint("Nachbestellen", "Vorhandene Nachbestellprüfung, Bestellliste und Bestellmengenvorschläge.")}
      ${renderTools("reorder")}`;
  }
  if (subTab === "journal") {
    content = `${renderToolManagementSubtabHint("Journal", "Vorhandenes Werkzeugjournal mit Filterleiste.")}
      ${renderTools("journal")}`;
  }
  if (subTab === "qrLabels") {
    content = renderToolQrLabelsTab();
  }
  if (subTab === "settings") {
    content = renderModulePlaceholder(
      "Werkzeug-Einstellungen",
      "Werkzeugeinstellungen folgen später.",
    );
  }
  return `<div class='space-y-4'>
    <div class='bg-white rounded-xl shadow p-4 space-y-3'>
      <div>
        <h2 class='text-lg font-semibold'>Werkzeugverwaltung</h2>
        <p class='text-sm text-slate-500 mt-1'>Bestehende Werkzeugfunktionen bleiben erhalten und sind hier einsortiert.</p>
      </div>
      <div class='flex gap-2 flex-wrap'>${renderModuleSubTabs(TOOL_MANAGEMENT_SUBTABS, subTab, "setToolManagementSubTab")}</div>
    </div>
    ${content}
  </div>`;
}

function renderToolManagementSubtabHint(title, text) {
  return `<div class='rounded border border-sky-200 bg-sky-50 p-3 text-sm text-sky-800'>
    <span class='font-semibold'>${escapeHtml(title)}:</span> ${escapeHtml(text)}
  </div>`;
}

function renderToolQrLabelsTab() {
  const tools = (state.tools || []).slice().sort((a, b) =>
    String(a.tNumber || "").localeCompare(String(b.tNumber || ""), "de", {
      numeric: true,
    }),
  );
  const rows = tools
    .map((tool) => {
      const label = [tool.label, formatToolSize(tool), tool.holder]
        .filter(Boolean)
        .join(" · ");
      return `<tr class='border-b'>
        <td class='p-2 whitespace-nowrap'>T ${escapeHtml(tool.tNumber || "-")}</td>
        <td class='p-2'>${escapeHtml(label || "-")}</td>
        <td class='p-2'>${escapeHtml(tool.articleNo || "-")}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="openToolQrPopup('${tool.id}')">QR anzeigen</button>
        </td>
      </tr>`;
    })
    .join("");
  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-xl font-bold'>QR-Etiketten</h2>
      <p class='text-sm text-slate-500 mt-1'>Vorhandene Werkzeug-QR-Funktionen pro Werkzeug öffnen.</p>
    </div>
    <div class='border rounded-lg overflow-auto max-h-[65vh]'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Werkzeug</th>
            <th class='p-2 text-left'>Bezeichnung</th>
            <th class='p-2 text-left'>Artikel</th>
            <th class='p-2 text-left'>Aktion</th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td class="p-3 text-slate-500" colspan="4">Keine Werkzeuge geladen.</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;
}

function renderStorageManagement() {
  const subTab = STORAGE_MANAGEMENT_SUBTABS.some(
    (tab) => tab.id === state.storageManagementSubTab,
  )
    ? state.storageManagementSubTab
    : "overview";
  let content = "";
  if (subTab === "overview") {
    content = `${renderStorageManagementSubtabHint("Lagerübersicht", "Regal-/Fachstruktur mit belegten Fächern und Fach-Popup. Für gezielte Suche bitte den Untertab Fachsuche verwenden.")}
      ${renderTools("storageOverview")}`;
  }
  if (subTab === "search") {
    content = `${renderStorageManagementSubtabHint("Fachsuche", "Sucheingabe, Trefferliste, Blinkmarkierung und Fach-Popup der vorhandenen Lagerfachansicht.")}
      ${renderTools("storageSearch")}`;
  }
  if (subTab === "move") {
    content = renderStorageMoveTab();
  }
  if (subTab === "qr") {
    content = renderStorageQrTab();
  }
  if (subTab === "inventory") {
    content = renderModulePlaceholder("Inventur", "Inventur bleibt aktuell deaktiviert.");
  }
  if (subTab === "settings") {
    content = renderModulePlaceholder("Lager-Einstellungen", "Lagereinstellungen folgen später.");
  }
  return `<div class='space-y-4'>
    <div class='bg-white rounded-xl shadow p-4 space-y-3'>
      <div>
        <h2 class='text-lg font-semibold'>Lagerverwaltung</h2>
        <p class='text-sm text-slate-500 mt-1'>Bestehende Lagerbezüge bleiben erhalten und bekommen feste Unterbereiche.</p>
      </div>
      <div class='flex gap-2 flex-wrap'>${renderModuleSubTabs(STORAGE_MANAGEMENT_SUBTABS, subTab, "setStorageManagementSubTab")}</div>
    </div>
    ${content}
  </div>`;
}

function renderStorageManagementSubtabHint(title, text) {
  return `<div class='rounded border border-sky-200 bg-sky-50 p-3 text-sm text-sky-800'>
    <span class='font-semibold'>${escapeHtml(title)}:</span> ${escapeHtml(text)}
  </div>`;
}

function renderStorageMoveTab() {
  const tools = (state.tools || []).slice().sort((a, b) =>
    String(a.tNumber || "").localeCompare(String(b.tNumber || ""), "de", {
      numeric: true,
    }),
  );
  const rows = tools
    .map((tool) => {
      const location = getToolStorageLocationKey(tool);
      return `<tr class='border-b'>
        <td class='p-2 whitespace-nowrap'>T ${escapeHtml(tool.tNumber || "-")}</td>
        <td class='p-2'>${escapeHtml(tool.label || "-")}</td>
        <td class='p-2'>${escapeHtml(formatToolSize(tool))}</td>
        <td class='p-2 whitespace-nowrap'>${escapeHtml(location || "-")}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-3 py-2 rounded bg-blue-700 text-white text-sm' onclick="openMoveToolModal('${tool.id}')">Umlagern</button>
        </td>
      </tr>`;
    })
    .join("");
  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-xl font-bold'>Umlagern</h2>
      <p class='text-sm text-slate-500 mt-1'>Vorhandene Werkzeug-Umlagerung aus der Lagerfachansicht verwenden.</p>
    </div>
    <div class='border rounded-lg overflow-auto max-h-[65vh]'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Werkzeug</th>
            <th class='p-2 text-left'>Bezeichnung</th>
            <th class='p-2 text-left'>Größe</th>
            <th class='p-2 text-left'>Fach</th>
            <th class='p-2 text-left'>Aktion</th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td class="p-3 text-slate-500" colspan="5">Keine Werkzeuge geladen.</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;
}

function renderStorageQrTab() {
  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-xl font-bold'>Lagerfach-QR</h2>
      <p class='text-sm text-slate-500 mt-1'>Vorhandene STORAGE:XX QR-Logik für Lagerfächer verwenden.</p>
    </div>
    <div class='border rounded-lg p-3 bg-slate-50 space-y-3'>
      <label class='text-sm font-medium'>Lagerfach
        <input id='storageQrLocationInput' class='border rounded p-2 w-full mt-1 bg-white' placeholder='12N oder STORAGE:12N' />
      </label>
      <div class='flex gap-2 flex-wrap'>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='openStorageQrFromInput()'>QR anzeigen</button>
        <button class='px-3 py-2 rounded bg-blue-700 text-white' onclick='printStorageQrFromInput()'>QR drucken</button>
        <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='openStorageLocationFromInput()'>Fach öffnen</button>
      </div>
    </div>
  </div>`;
}

function readStorageQrInputLocation() {
  const rawValue = document.getElementById("storageQrLocationInput")?.value || "";
  const locationKey = rawValue.toUpperCase().startsWith("STORAGE:")
    ? parseStorageQrCode(rawValue)
    : normalizeStorageLocation(rawValue);
  return isValidStorageLocation(locationKey) ? locationKey : "";
}

function openStorageQrFromInput() {
  const locationKey = readStorageQrInputLocation();
  if (!locationKey) {
    alert("Bitte ein gültiges Lagerfach eingeben, z. B. 12N oder STORAGE:12N.");
    return;
  }
  openStorageQrModal(locationKey);
}

function openStorageLocationFromInput() {
  const locationKey = readStorageQrInputLocation();
  if (!locationKey) {
    alert("Bitte ein gültiges Lagerfach eingeben, z. B. 12N oder STORAGE:12N.");
    return;
  }
  openStorageLocationModal(locationKey);
}

function printStorageQrFromInput() {
  const locationKey = readStorageQrInputLocation();
  if (!locationKey) {
    alert("Bitte ein gültiges Lagerfach eingeben, z. B. 12N oder STORAGE:12N.");
    return;
  }
  printStorageQrLabel(locationKey);
}

function renderScannerModule() {
  const subTab = SCANNER_SUBTABS.some((tab) => tab.id === state.scannerSubTab)
    ? state.scannerSubTab
    : "toolScan";
  let content = "";
  if (subTab === "toolScan") content = renderToolScanner("all");
  if (subTab === "withdraw") content = renderToolScanner("withdraw");
  if (subTab === "restock") content = renderToolScanner("restock");
  if (subTab === "storageScan") {
    content = renderModulePlaceholder("Lagerfach scannen", "Lagerfach-Scanner wird später ergänzt.");
  }
  if (subTab === "return") {
    content = renderModulePlaceholder("Rückgabe", "Rückgabe wird später als Scanner-Flow ergänzt.");
  }
  if (subTab === "countScan") {
    content = renderModulePlaceholder("Stückzahl-Scan", "Stückzahl-Scan wird später mit dem Produktionszähler verbunden.");
  }
  return `<div class='space-y-4'>
    <div class='bg-white rounded-xl shadow p-4 space-y-3'>
      <div>
        <h2 class='text-lg font-semibold'>Scanner</h2>
        <p class='text-sm text-slate-500 mt-1'>Scanner-Flows für Werkzeug, Lager und Produktion.</p>
      </div>
      <div class='flex gap-2 flex-wrap'>${renderModuleSubTabs(SCANNER_SUBTABS, subTab, "setScannerSubTab")}</div>
    </div>
    ${content}
  </div>`;
}

function renderAnalyticsModule() {
  if (currentUser?.role !== "admin") {
    return `<div class='bg-white rounded-xl shadow p-4'><p>Kein Zugriff.</p></div>`;
  }
  const subTab = ANALYTICS_SUBTABS.some((tab) => tab.id === state.analyticsSubTab)
    ? state.analyticsSubTab
    : "toolStats";
  let content = "";
  if (subTab === "toolStats") content = renderStats();
  if (subTab === "orderStats") content = renderOrderStats();
  if (subTab === "productionStats") {
    content = renderModulePlaceholder("Produktionszahlen", "Produktionsauswertungen folgen später auf Basis der Stückzahl-Zähler.");
  }
  if (subTab === "employeesDepartments") {
    content = renderModulePlaceholder("Mitarbeiter / Abteilung", "Auswertung nach Mitarbeitern und Abteilungen folgt später.");
  }
  if (subTab === "protocols") content = renderConflicts();
  if (subTab === "export") {
    content = renderModulePlaceholder("Export", "Exportfunktionen werden später ergänzt.");
  }
  return `<div class='space-y-4'>
    <div class='bg-white rounded-xl shadow p-4 space-y-3'>
      <div>
        <h2 class='text-lg font-semibold'>Auswertung</h2>
        <p class='text-sm text-slate-500 mt-1'>Statistiken, Protokolle und spätere Exporte.</p>
      </div>
      <div class='flex gap-2 flex-wrap'>${renderModuleSubTabs(ANALYTICS_SUBTABS, subTab, "setAnalyticsSubTab")}</div>
    </div>
    ${content}
  </div>`;
}

function renderAdminSystemModule() {
  if (currentUser?.role !== "admin") {
    return `<div class='bg-white rounded-xl shadow p-4'><p>Kein Zugriff.</p></div>`;
  }
  const subTab = ADMIN_SYSTEM_SUBTABS.some(
    (tab) => tab.id === state.adminSystemSubTab,
  )
    ? state.adminSystemSubTab
    : "roles";
  let content = "";
  if (subTab === "roles") {
    content = renderModulePlaceholder("Rollen", "Rollenverwaltung bleibt vorbereitet und wird später ausgebaut.");
  }
  if (subTab === "accounts") {
    content = renderModulePlaceholder("Benutzer / Login-Konten", "Login-Konten werden weiterhin nicht automatisch verändert.");
  }
  if (subTab === "company") {
    content = renderModulePlaceholder("Firmen-Einstellungen", "Firmen-Einstellungen folgen später.");
  }
  if (subTab === "version") {
    content = `<div class='border rounded-lg p-4 bg-slate-50 space-y-3'>
      <h3 class='font-semibold'>Version / Logs</h3>
      <p class='text-sm text-slate-600'>Aktuelle Version: v${APP_VERSION}</p>
      <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='openVersionLog()'>Versionslog öffnen</button>
    </div>`;
  }
  if (subTab === "backup") {
    content = renderModulePlaceholder("Backup", "Backup-Funktionen werden später ergänzt.");
  }
  if (subTab === "setup") {
    content = renderModulePlaceholder("Setup-Assistent", "Setup-Assistent wird später ergänzt.");
  }
  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-lg font-semibold'>Admin / System</h2>
      <p class='text-sm text-slate-500 mt-1'>Systembereiche als feste Funktionslandkarte.</p>
    </div>
    <div class='flex gap-2 flex-wrap'>${renderModuleSubTabs(ADMIN_SYSTEM_SUBTABS, subTab, "setAdminSystemSubTab")}</div>
    ${content}
  </div>`;
}

function canAccessProduction() {
  return ["admin", "department_admin"].includes(currentUser?.role);
}

function canManageProduction() {
  return currentUser?.role === "admin";
}

function isToolBelowMinStock(tool) {
  return getToolStockWarningLevel(tool) !== "";
}

function getToolStockWarningLevel(tool) {
  const stock = Number(tool.stock || 0);
  const minStock = Number(tool.minStock || 0);

  if (minStock <= 0) return "";
  if (stock < minStock) return "critical";
  if (stock === minStock) return "warning";
  return "";
}

function getDefaultToolBookingQty(tool) {
  const insertEdges = Number(tool?.insertEdges || tool?.insert_edges || 0);
  if (tool?.insertTool && Number.isFinite(insertEdges) && insertEdges > 0) {
    return insertEdges;
  }
  return 1;
}

function readPositiveQtyInput(inputId) {
  const qty = Number(document.getElementById(inputId)?.value || 0);
  return Number.isFinite(qty) && qty > 0 ? qty : 0;
}

function normalizeStorageLocation(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "")
    .toUpperCase();
}

function isValidStorageLocation(value) {
  const normalized = normalizeStorageLocation(value);
  const match = normalized.match(/^(\d{1,2})([A-Y])$/);
  if (!match) return false;

  const rack = Number(match[1]);
  return rack >= 1 && rack <= 26;
}

function parseStorageQrCode(value) {
  const normalized = String(value || "")
    .trim()
    .toUpperCase();
  if (!normalized.startsWith("STORAGE:")) return null;

  const locationKey = normalizeStorageLocation(
    normalized.slice("STORAGE:".length),
  );
  return isValidStorageLocation(locationKey) ? locationKey : null;
}

function buildStorageQrValue(locationKey) {
  return `STORAGE:${normalizeStorageLocation(locationKey)}`;
}

function openStorageLocationByQr(value) {
  const locationKey = parseStorageQrCode(value);
  if (!locationKey) {
    alert("Ungültiger Lagerfach-QR-Code");
    return;
  }

  openStorageLocationModal(locationKey);
}

function getToolStorageLocationKey(tool) {
  return normalizeStorageLocation(
    tool?.location || tool?.storageLocation || tool?.shelf || "",
  );
}

function findToolByStorageLocation(locationKey, excludeToolId = null) {
  const normalizedKey = normalizeStorageLocation(locationKey);
  if (!normalizedKey) return null;

  const tools = Array.isArray(state.tools) ? state.tools : [];
  return (
    tools.find((tool) => {
      return (
        tool.id !== excludeToolId &&
        getToolStorageLocationKey(tool) === normalizedKey
      );
    }) || null
  );
}

function buildToolStorageMap() {
  const storageMap = {};
  const tools = Array.isArray(state.tools) ? state.tools : [];

  tools.forEach((tool) => {
    const locationKey = getToolStorageLocationKey(tool);
    if (!locationKey) return;
    if (!storageMap[locationKey]) storageMap[locationKey] = [];
    storageMap[locationKey].push(tool);
  });

  return storageMap;
}

function getStorageCellState(tools) {
  if (!Array.isArray(tools) || tools.length === 0) return "empty";
  if (tools.some((tool) => getToolStockWarningLevel(tool) === "critical")) {
    return "critical";
  }
  if (tools.some((tool) => getToolStockWarningLevel(tool) === "warning")) {
    return "warning";
  }
  return "ok";
}

function ensureStorageHighlightStyles() {
  if (document.getElementById("storageHighlightStyles")) return;

  const style = document.createElement("style");
  style.id = "storageHighlightStyles";
  style.textContent = `
    @keyframes storageHitBlink {
      0%, 100% {
        background-color: #dbeafe;
        border-color: #1d4ed8;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.9), 0 0 14px rgba(37, 99, 235, 0.75);
        transform: scale(1);
      }
      50% {
        background-color: #60a5fa;
        border-color: #1e3a8a;
        box-shadow: 0 0 0 4px rgba(29, 78, 216, 1), 0 0 22px rgba(29, 78, 216, 0.95);
        transform: scale(1.08);
      }
    }

    .storage-hit-blink {
      animation: storageHitBlink 0.9s ease-in-out infinite;
      position: relative;
      z-index: 1;
    }
  `;
  document.head.appendChild(style);
}

function findStorageSearchMatches() {
  const raw = String(state.storageSearch || "").trim();
  const normalized = raw.toLowerCase();
  if (!normalized) return [];

  const tools = Array.isArray(state.tools) ? state.tools : [];
  const tNumberMatch = normalized.match(/^t?\s*(\d+)$/);

  if (tNumberMatch) {
    const target = Number(tNumberMatch[1]);
    return tools.filter((tool) => {
      const toolNumber = Number(
        String(tool.tNumber || "")
          .replace(/^t/i, "")
          .trim(),
      );
      return toolNumber === target;
    });
  }

  return tools.filter((tool) =>
    [tool.tNumber, tool.label, tool.articleNo, tool.manufacturer].some((value) =>
      String(value || "")
        .toLowerCase()
        .includes(normalized),
    ),
  );
}

function applyStorageSearch() {
  state.storageSearch =
    document.getElementById("storageSearchInput")?.value || "";
  persist();
  render();
}

function resetStorageSearch() {
  state.storageSearch = "";
  persist();
  render();
}

function shouldOrderTool(tool) {
  return isToolBelowMinStock(tool);
}

function tabNeedsAttention(tab) {
  if (tab === "planung") return generateThreeMonths().some((s) => s.open);
  if (tab === "todo") return state.tasks.some((t) => t.status !== "done");
  if (tab === "dashboard") return state.tasks.some((t) => t.status !== "done");
  if (tab === "werkzeugverwaltung" && currentUser?.role === "admin")
    return state.tools.some((t) => shouldOrderTool(t));
  if (tab === "auswertung" && currentUser?.role === "admin")
    return (state.orderHistory || []).length > 0;
  if (tab === "werkzeuge" && currentUser?.role === "admin")
    return state.tools.some((t) => shouldOrderTool(t));
  if (tab === "bestellstatistik" && currentUser?.role === "admin")
    return (state.orderHistory || []).length > 0;
  if (tab === "konflikte")
    return Object.values(state.conflicts).some((c) => c.resolved !== true);
  if (tab === "meine" && currentUser?.role === "employee")
    return generateThreeMonths().some(
      (s) => s.assigned === currentUser.name && s.open,
    );
  return false;
}

function generateThreeMonths() {
  const shifts = [];
  const today = new Date();
  const start = new Date(today);
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
  const anchor = new Date(`${ROTATION_ANCHOR_MONDAY}T00:00:00`);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;

  for (let i = 0; i < 90; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const iso = d.toISOString().slice(0, 10);
    const weekIndex = ((Math.floor((d - anchor) / msPerWeek) % 6) + 6) % 6;
    const weekday = d.getDay();
    const template = WEEK_TEMPLATES[weekIndex];

    if (weekday >= 1 && weekday <= 5) {
      template.mondayToFriday.forEach((s, idx) =>
        shifts.push(buildShift(iso, `${iso}-mf-${idx}`, s)),
      );
    }
    if (weekday === 6) {
      template.saturday.forEach((s, idx) =>
        shifts.push(buildShift(iso, `${iso}-sa-${idx}`, s)),
      );
    }
    if (weekday === 0) {
      template.sunday.forEach((s, idx) =>
        shifts.push(buildShift(iso, `${iso}-su-${idx}`, s)),
      );
    }
  }

  return shifts;
}

function getReplacementForShift(shiftId) {
  return state.absenceReplacements?.[shiftId] || null;
}

function getReplacementEntriesForSource(sourceType, sourceId) {
  return Object.entries(state.absenceReplacements || {}).filter(([, entry]) => {
    return (
      entry?.sourceType === sourceType &&
      String(entry?.sourceId) === String(sourceId)
    );
  });
}

function getReplacementSummaryForSource(sourceType, sourceId) {
  const entries = getReplacementEntriesForSource(sourceType, sourceId);
  if (!entries.length) return "Kein Ersatz geplant";

  const normalized = entries.map(([, entry]) => {
    if (entry.mode === "cancel") {
      return `${formatDateDisplay(entry.weekFrom)} bis ${formatDateDisplay(entry.weekTo)}: AUSFALL`;
    }
    return `${formatDateDisplay(entry.weekFrom)} bis ${formatDateDisplay(entry.weekTo)}: ${entry.replacementUser}`;
  });

  return normalized.join("<br>");
}

function applyReplacementToShift(shiftId, replacementEntry) {
  if (!state.absenceReplacements) state.absenceReplacements = {};
  state.absenceReplacements[shiftId] = replacementEntry;

  if (replacementEntry.mode === "cancel") {
    state.shiftCancellations[shiftId] = true;
    delete state.assignments[shiftId];
    return;
  }

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = replacementEntry.replacementUser;
}

function getReplacementShiftIdsForSource(sourceType, sourceId) {
  return getReplacementEntriesForSource(sourceType, sourceId).map(
    ([shiftId]) => shiftId,
  );
}

function clearReplacementPlanForSource(sourceType, sourceId) {
  if (!state.absenceReplacements) return;
  Object.keys(state.absenceReplacements).forEach((shiftId) => {
    const entry = state.absenceReplacements[shiftId];
    if (
      entry?.sourceType === sourceType &&
      String(entry?.sourceId) === String(sourceId)
    ) {
      delete state.absenceReplacements[shiftId];
      delete state.shiftCancellations[shiftId];
      delete state.assignments[shiftId];
    }
  });
}

async function deleteReplacementPlanFromSupabase(
  sourceType,
  sourceId,
  shiftIds = null,
) {
  if (!supabaseReady) return null;

  const targetShiftIds =
    shiftIds || getReplacementShiftIdsForSource(sourceType, sourceId);

  const operations = [
    supabaseClient
      .from("planner_absence_replacements")
      .delete()
      .eq("source_type", sourceType)
      .eq("source_id", sourceId),
  ];

  if (targetShiftIds.length) {
    operations.push(
      supabaseClient
        .from("planner_assignments")
        .delete()
        .in("shift_id", targetShiftIds),
    );
    operations.push(
      supabaseClient
        .from("planner_shift_cancellations")
        .delete()
        .in("shift_id", targetShiftIds),
    );
  }

  const results = await Promise.all(operations);
  const firstError = results.find((res) => res.error)?.error;

  return firstError || null;
}

async function persistReplacementShiftEffects(replacementEntries) {
  if (!supabaseReady || !replacementEntries.length) return null;

  const assignmentRows = [];
  const cancellationRows = [];
  const replacedShiftIds = [];
  const canceledShiftIds = [];

  replacementEntries.forEach(([shiftId, entry]) => {
    const shift = getShiftById(shiftId);
    if (!shift) return;

    if (entry.mode === "cancel") {
      canceledShiftIds.push(shiftId);
      cancellationRows.push({
        shift_id: shiftId,
        shift_date: shift.date,
        created_by_employee_id: currentEmployeeRecord?.id || null,
      });
      return;
    }

    if (entry.mode === "replace" && entry.replacementUser) {
      replacedShiftIds.push(shiftId);
      assignmentRows.push({
        shift_id: shiftId,
        shift_date: shift.date,
        assigned_user: entry.replacementUser,
        created_by_employee_id: currentEmployeeRecord?.id || null,
      });
    }
  });

  const operations = [];

  if (replacedShiftIds.length) {
    operations.push(
      supabaseClient
        .from("planner_shift_cancellations")
        .delete()
        .in("shift_id", replacedShiftIds),
    );
  }

  if (canceledShiftIds.length) {
    operations.push(
      supabaseClient
        .from("planner_assignments")
        .delete()
        .in("shift_id", canceledShiftIds),
    );
  }

  if (assignmentRows.length) {
    operations.push(
      supabaseClient
        .from("planner_assignments")
        .upsert(assignmentRows, { onConflict: "shift_id" }),
    );
  }

  if (cancellationRows.length) {
    operations.push(
      supabaseClient
        .from("planner_shift_cancellations")
        .upsert(cancellationRows, { onConflict: "shift_id" }),
    );
  }

  const results = await Promise.all(operations);
  const firstError = results.find((res) => res.error)?.error;

  return firstError || null;
}

async function deleteReplacementEntriesFromSupabase(sourceType, sourceId) {
  if (!supabaseReady) return null;

  const { error } = await supabaseClient
    .from("planner_absence_replacements")
    .delete()
    .eq("source_type", sourceType)
    .eq("source_id", sourceId);

  return error || null;
}

async function saveReplacementPlanToSupabase(replacementEntries) {
  if (!supabaseReady || !replacementEntries.length) return null;

  const rows = replacementEntries.map(([shiftId, entry]) => {
    const shift = getShiftById(shiftId);
    return {
      shift_id: shiftId,
      shift_date: shift?.date || entry.weekFrom || entry.from,
      source_type: entry.sourceType,
      source_id: entry.sourceId,
      absent_user: entry.absentUser,
      from_date: entry.from,
      to_date: entry.to,
      mode: entry.mode,
      replacement_user: entry.replacementUser,
      week_from: entry.weekFrom,
      week_to: entry.weekTo,
    };
  });

  const { error } = await supabaseClient
    .from("planner_absence_replacements")
    .insert(rows);

  if (error) return error;

  return persistReplacementShiftEffects(replacementEntries);
}

async function clearAbsenceReplacementPlan(sourceType, sourceId) {
  const shiftIds = getReplacementShiftIdsForSource(sourceType, sourceId);
  const error = await deleteReplacementPlanFromSupabase(sourceType, sourceId);
  if (error) {
    console.error("Fehler beim Löschen der Ersatzplanung:", error);
    return alert(
      `Ersatzplanung konnte nicht gelöscht werden: ${error.message}`,
    );
  }

  clearReplacementPlanForSource(sourceType, sourceId);
  persist();
  render();
}

function getAbsenceEntry(type, entryId) {
  const list = type === "vacation" ? state.vacations : state.sickLeaves;
  return list.find((entry) => String(entry.id) === String(entryId)) || null;
}

function getPendingReplacementShifts(type, entryId) {
  const entry = getAbsenceEntry(type, entryId);
  if (!entry) return [];

  return getShiftsOfUserInRange(entry.user, entry.from, entry.to).filter(
    (shift) => {
      const replacement = state.absenceReplacements?.[shift.id];
      return !(
        replacement?.sourceType === type &&
        String(replacement?.sourceId) === String(entryId)
      );
    },
  );
}

function groupReplacementShiftsByWeek(shifts) {
  const grouped = {};

  shifts.forEach((shift) => {
    const key = weekKey(shift.date);
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(shift);
  });

  return Object.entries(grouped).map(([weekStart, weekShifts]) => ({
    weekStart,
    shifts: weekShifts,
  }));
}

function getReplacementPlannerSelectionKey(type, entryId) {
  return `${type}:${entryId}`;
}

function getReplacementPlannerChoice(type, entryId) {
  const key = getReplacementPlannerSelectionKey(type, entryId);
  return state.replacementPlannerChoice?.[key] || "AUSFALL";
}

function setReplacementPlannerChoice(type, entryId, value) {
  const key = getReplacementPlannerSelectionKey(type, entryId);
  if (!state.replacementPlannerChoice) state.replacementPlannerChoice = {};
  state.replacementPlannerChoice[key] = value || "AUSFALL";
  persist();
}

function ensureReplacementPlannerSelection(type, entryId) {
  const key = getReplacementPlannerSelectionKey(type, entryId);
  if (!state.replacementPlannerSelection)
    state.replacementPlannerSelection = {};
  if (!state.replacementPlannerSelection[key]) {
    state.replacementPlannerSelection[key] = {};
  }
  return state.replacementPlannerSelection[key];
}

function renderReplacementCalendar(type, entryId) {
  const pendingShifts = getPendingReplacementShifts(type, entryId);
  const selection = ensureReplacementPlannerSelection(type, entryId);
  const groupedWeeks = groupReplacementShiftsByWeek(pendingShifts);

  if (!pendingShifts.length) {
    return `<div class="text-sm rounded border border-emerald-200 bg-emerald-50 p-3 text-emerald-800">
      Alle Schichten in diesem Zeitraum sind ersetzt oder als Ausfall markiert.
    </div>`;
  }

  return groupedWeeks
    .map(({ weekStart, shifts }, index) => {
      const weekEnd = new Date(`${weekStart}T00:00:00`);
      weekEnd.setDate(weekEnd.getDate() + 6);

      const shiftButtons = shifts
        .map((shift) => {
          const selected = !!selection[shift.id];
          return `<button class="text-left border rounded p-2 ${selected ? "bg-slate-900 text-white border-slate-900" : "bg-white border-slate-300"}" onclick="toggleReplacementDay('${type}', '${entryId}', '${shift.id}')">
            <div class="font-semibold">${formatDateWithWeekday(shift.date)}</div>
            <div class="text-xs">${shift.label} · ${shift.start}-${shift.end}</div>
          </button>`;
        })
        .join("");

      return `<div class="border rounded-lg p-3 bg-slate-50">
        <div class="flex items-center justify-between gap-2 mb-2">
          <div class="font-semibold">KW ${index + 1}: ${formatDateDisplay(weekStart)} bis ${formatDateDisplay(isoDate(weekEnd))}</div>
          <button class="px-2 py-1 rounded bg-slate-700 text-white text-sm" onclick="applyReplacementForWeek('${type}', '${entryId}', '${weekStart}')">Diese Woche</button>
        </div>
        <div class="grid md:grid-cols-2 gap-2">${shiftButtons}</div>
      </div>`;
    })
    .join("");
}

function toggleReplacementDay(type, entryId, shiftId) {
  const selection = ensureReplacementPlannerSelection(type, entryId);
  selection[shiftId] = !selection[shiftId];
  openAbsenceReplacementPlanner(type, entryId);
}

async function applyReplacementForSelectedDays(
  type,
  entryId,
  selectedShiftIds,
) {
  if (!supabaseReady) return;

  const entry = getAbsenceEntry(type, entryId);
  if (!entry) return;

  const modeValue =
    document.getElementById("replacementUserSelect")?.value ||
    getReplacementPlannerChoice(type, entryId);
  setReplacementPlannerChoice(type, entryId, modeValue);
  const mode = modeValue === "AUSFALL" ? "cancel" : "replace";
  const replacementUser = mode === "replace" ? modeValue : null;

  if (!selectedShiftIds.length) {
    alert("Bitte mindestens eine Schicht auswählen.");
    return;
  }

  const newEntries = selectedShiftIds
    .map((shiftId) => {
      const shift = getShiftById(shiftId);
      if (!shift) return null;

      return [
        shiftId,
        {
          sourceType: type,
          sourceId: entryId,
          absentUser: entry.user,
          from: entry.from,
          to: entry.to,
          mode,
          replacementUser,
          weekFrom: weekKey(shift.date),
          weekTo: shift.date,
        },
      ];
    })
    .filter(Boolean);

  const deleteError = await deleteReplacementEntriesFromSupabase(type, entryId);
  if (deleteError) {
    console.error("Fehler beim Vorbereiten der Ersatzplanung:", deleteError);
    return alert(
      `Ersatzplanung konnte nicht vorbereitet werden: ${deleteError.message}`,
    );
  }

  newEntries.forEach(([shiftId, replacementEntry]) => {
    applyReplacementToShift(shiftId, replacementEntry);
  });

  const allEntries = getReplacementEntriesForSource(type, entryId);
  const saveError = await saveReplacementPlanToSupabase(allEntries);
  if (saveError) {
    console.error("Fehler beim Speichern der Ersatzplanung:", saveError);
    return alert(
      `Ersatzplanung konnte nicht gespeichert werden: ${saveError.message}`,
    );
  }

  const selectionKey = getReplacementPlannerSelectionKey(type, entryId);
  if (state.replacementPlannerSelection) {
    state.replacementPlannerSelection[selectionKey] = {};
  }

  persist();

  if (getPendingReplacementShifts(type, entryId).length) {
    openAbsenceReplacementPlanner(type, entryId);
    return;
  }

  if (state.replacementPlannerSelection) {
    delete state.replacementPlannerSelection[selectionKey];
  }
  if (state.replacementPlannerChoice) {
    delete state.replacementPlannerChoice[selectionKey];
  }

  getModalHost().innerHTML = "";
  render();
}

function applyReplacementForWeek(type, entryId, weekStart) {
  const pendingShiftIds = getPendingReplacementShifts(type, entryId)
    .filter((shift) => weekKey(shift.date) === weekStart)
    .map((shift) => shift.id);

  applyReplacementForSelectedDays(type, entryId, pendingShiftIds);
}

function applyReplacementForAll(type, entryId) {
  const pendingShiftIds = getPendingReplacementShifts(type, entryId).map(
    (shift) => shift.id,
  );

  applyReplacementForSelectedDays(type, entryId, pendingShiftIds);
}

function openAbsenceReplacementPlanner(type, entryId) {
  const entry = getAbsenceEntry(type, entryId);
  if (!entry) return;

  const host = getModalHost();
  const pendingShifts = getPendingReplacementShifts(type, entryId);
  const selection = ensureReplacementPlannerSelection(type, entryId);
  const selectedReplacementValue = getReplacementPlannerChoice(type, entryId);
  const selectedShiftIds = Object.entries(selection)
    .filter(([, selected]) => selected)
    .map(([shiftId]) => shiftId)
    .filter((shiftId) => pendingShifts.some((shift) => shift.id === shiftId));

  const replacementOptions = [
    `<option value="AUSFALL" ${selectedReplacementValue === "AUSFALL" ? "selected" : ""}>Ausfall</option>`,
    ...activeUsers()
      .filter((user) => user.name !== entry.user)
      .map(
        (user) =>
          `<option value="${user.name}" ${selectedReplacementValue === user.name ? "selected" : ""}>${user.name}</option>`,
      ),
  ].join("");

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl max-h-[90vh] overflow-auto p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">Ersatzplanung</h3>
          <p class="text-sm text-slate-600">${entry.user}: ${formatDateDisplay(entry.from)} bis ${formatDateDisplay(entry.to)}</p>
        </div>
        <div class="flex items-center gap-2">
          ${helpButton("ersatzplanung")}
          <button class="px-3 py-1 rounded bg-slate-200" onclick="closeReplacementPlanner('${type}', '${entryId}')">Abbrechen</button>
        </div>
      </div>

      <div class="grid md:grid-cols-[minmax(180px,260px)_1fr] gap-4">
        <div class="space-y-3">
          <label class="block text-sm font-medium">
            Ersatz-Mitarbeiter
            <select id="replacementUserSelect" class="border rounded p-2 w-full mt-1" onchange="setReplacementPlannerChoice('${type}', '${entryId}', this.value)">${replacementOptions}</select>
          </label>
          <button class="px-3 py-2 rounded bg-slate-900 text-white w-full" onclick="applyReplacementForAll('${type}', '${entryId}')">Alle</button>
          <button class="px-3 py-2 rounded bg-emerald-700 text-white w-full" onclick="applyReplacementForSelectedDays('${type}', '${entryId}', ${JSON.stringify(selectedShiftIds).replaceAll('"', "&quot;")})">OK</button>
          <div class="text-xs text-slate-500">${pendingShifts.length} offene Schicht(en), ${selectedShiftIds.length} ausgewählt.</div>
        </div>
        <div class="space-y-3">${renderReplacementCalendar(type, entryId)}</div>
      </div>
    </div>
  </div>`;
}

function closeReplacementPlanner(type = null, entryId = null) {
  if (type && entryId && state.replacementPlannerSelection) {
    const key = getReplacementPlannerSelectionKey(type, entryId);
    delete state.replacementPlannerSelection[key];
    delete state.replacementPlannerChoice?.[key];
    persist();
  }

  getModalHost().innerHTML = "";
  render();
}

function buildShift(date, id, template) {
  const defaultAssigned = chooseDefault(template.options);
  const swappedDefault = defaultAssigned
    ? applySwap(date, defaultAssigned)
    : defaultAssigned;

  const isOptional = template.options.length > 1;
  const manualAssigned = state.assignments[id] || null;
  const replacement = getReplacementForShift(id);
  const specialDay = getSpecialDay(date);
  const blocked = isBlockedDay(date);

  const absenceKey = swappedDefault ? `${date}:${swappedDefault}` : null;
  const calendarAbsenceType = swappedDefault
    ? getCalendarAbsenceType(swappedDefault, date)
    : null;
  const manualAbsent = absenceKey ? !!state.absences[absenceKey] : false;
  const absent = manualAbsent || !!calendarAbsenceType;

  const canceled =
    !!state.shiftCancellations[id] || replacement?.mode === "cancel";

  let assigned = null;

  if (manualAssigned) {
    assigned = manualAssigned;
  } else if (blocked || canceled) {
    assigned = null;
  } else if (replacement?.mode === "replace" && replacement?.replacementUser) {
    assigned = replacement.replacementUser;
  } else if (!isOptional && !absent) {
    assigned = swappedDefault;
  } else {
    assigned = null;
  }

  const open = !blocked && (canceled || !assigned);

  return {
    id,
    date,
    label: template.label,
    start: template.start,
    end: template.end,
    options: template.options,
    assigned,
    originalAssigned: blocked ? null : swappedDefault,
    absenceType: calendarAbsenceType || (manualAbsent ? "abwesend" : null),
    replacement,
    open,
    blocked,
    specialDay,
  };
}

function chooseDefault(options) {
  const slot = options[0];
  if (slot === "NONE") return null;
  return slotToName(slot);
}

function slotToName(slot) {
  return (
    state.slotAssignments?.[slot] ||
    allUsers().find((u) => u.slot === slot)?.name ||
    null
  );
}

function userByName(name) {
  return allUsers().find((u) => u.name === name);
}

function isCoreEmployee(name) {
  return userByName(name)?.type === "core";
}

function isSpringer(name) {
  return userByName(name)?.type === "springer";
}

function slotOfUser(name) {
  const mappedSlot = Object.entries(state.slotAssignments || {}).find(
    ([, assignedName]) => assignedName === name,
  )?.[0];
  return mappedSlot || allUsers().find((u) => u.name === name)?.slot || "-";
}

const PERSON_COLORS = {
  Lavdrim: "bg-green-100 text-green-900",
  Roger: "bg-violet-100 text-violet-900",
  Dashmir: "bg-rose-100 text-rose-900",
  Thomas: "bg-cyan-100 text-cyan-900",
  Musa: "bg-amber-100 text-amber-900",
  Ardian: "bg-indigo-100 text-indigo-900",
};

function personColorClasses(name) {
  const fixedColors = {
    Lavdrim: "bg-green-200 text-green-900",
    Roger: "bg-blue-200 text-blue-900",
    Dashmir: "bg-yellow-200 text-yellow-900",
    Thomas: "bg-purple-200 text-purple-900",
    Musa: "bg-orange-200 text-orange-900",
    Ardian: "bg-teal-200 text-teal-900",
  };

  const colorByKey = {
    green: "bg-green-200 text-green-900",
    blue: "bg-blue-200 text-blue-900",
    yellow: "bg-yellow-200 text-yellow-900",
    purple: "bg-purple-200 text-purple-900",
    orange: "bg-orange-200 text-orange-900",
    teal: "bg-teal-200 text-teal-900",
    cyan: "bg-cyan-200 text-cyan-900",
    pink: "bg-pink-200 text-pink-900",
    rose: "bg-rose-200 text-rose-900",
    lime: "bg-lime-200 text-lime-900",
    gray: "bg-gray-200 text-gray-900",
  };

  const emp = Object.values(state.employees || {}).find(
    (e) => e.display_name === name || e.name === name,
  );

  const fromDb = emp?.color_key ? colorByKey[emp.color_key] : null;
  const raw = fromDb || fixedColors[name] || "bg-gray-200 text-gray-900";

  const parts = raw.split(" ");

  return {
    bg: parts[0],
    text: parts[1],
    raw,
  };
}

function personBorderClass(name) {
  const slot = slotOfUser(name);
  if (slot === "A") return "border-green-500";
  if (slot === "B") return "border-violet-500";
  if (slot === "C") return "border-rose-500";
  if (slot === "D") return "border-cyan-500";
  if (slot === "E") return "border-amber-500";
  if (slot === "F") return "border-indigo-500";
  return "border-slate-400";
}

function formatSlot(options) {
  return options.map((slot) => (slot === "NONE" ? "0" : slot)).join("/");
}

function slotColor(slotText) {
  if (slotText.includes("A")) return "bg-green-100 text-green-900";
  if (slotText.includes("B")) return "bg-violet-100 text-violet-900";
  if (slotText.includes("C")) return "bg-rose-100 text-rose-900";
  return "bg-slate-100 text-slate-800";
}

function weekStartFromMonday(baseDate, addWeeks) {
  const d = new Date(baseDate);
  d.setDate(d.getDate() + addWeeks * 7);
  return d;
}

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}

function getSpecialDay(date) {
  return state.specialDays?.[date] || null;
}

function isBlockedDay(date) {
  const d = getSpecialDay(date);
  if (!d) return false;

  return d.type === "holiday" || d.type === "bridge" || d.type === "company";
}

function todayIso() {
  return isoDate(new Date());
}

function formatDateDisplay(iso) {
  if (!iso || !iso.includes("-")) return iso || "";
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y.slice(2)}`;
}

function getWeekdayName(iso) {
  if (!iso) return "";
  const date = new Date(`${iso}T00:00:00`);
  return ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"][date.getDay()];
}

function formatDateWithWeekday(iso) {
  return `${getWeekdayName(iso)} ${formatDateDisplay(iso)}`;
}

function formatDateTime(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString("de-CH", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function inRange(date, from, to) {
  return date >= from && date <= to;
}

function applySwap(date, name) {
  let resolved = name;
  state.swaps.forEach((swap) => {
    if (!swap || date < swap.startDate) return;
    if (swap.endDate && date > swap.endDate) return;
    if (resolved === swap.userA) resolved = swap.userB;
    else if (resolved === swap.userB) resolved = swap.userA;
  });
  return resolved;
}

function hasCalendarAbsence(name, date) {
  const inVacation = state.vacations.some(
    (v) => v.user === name && inRange(date, v.from, v.to),
  );
  const inSick = state.sickLeaves.some(
    (s) => s.user === name && inRange(date, s.from, s.to),
  );
  return inVacation || inSick;
}

function getCalendarAbsenceType(name, date) {
  if (!name) return null;
  if (
    state.sickLeaves.some((s) => s.user === name && inRange(date, s.from, s.to))
  )
    return "krank";
  if (
    state.vacations.some((v) => v.user === name && inRange(date, v.from, v.to))
  )
    return "urlaub";
  return null;
}

function resolveAssigned(shiftId, options) {
  const replacement = getReplacementForShift(shiftId);
  if (replacement?.mode === "cancel") return null;
  if (replacement?.replacementUser) return replacement.replacementUser;

  const date = shiftId.slice(0, 10);
  const defaultAssigned = chooseDefault(options);
  const swappedDefault = defaultAssigned
    ? applySwap(date, defaultAssigned)
    : defaultAssigned;
  const isOptional = options.length > 1;
  const absent =
    !!state.absences[`${date}:${swappedDefault}`] ||
    hasCalendarAbsence(swappedDefault, date);
  const manualAssigned = state.assignments[shiftId] || null;
  if (state.shiftCancellations[shiftId]) return null;
  return !isOptional && !absent
    ? swappedDefault
    : manualAssigned || (isOptional ? swappedDefault : null);
}

function resolveAbsenceType(shiftId, options) {
  const date = shiftId.slice(0, 10);
  const defaultAssigned = chooseDefault(options);
  const swappedDefault = defaultAssigned
    ? applySwap(date, defaultAssigned)
    : defaultAssigned;
  if (getCalendarAbsenceType(swappedDefault, date))
    return getCalendarAbsenceType(swappedDefault, date);
  if (state.absences[`${date}:${swappedDefault}`]) return "abwesend";
  return null;
}

function assignedDisplay(shiftId, options) {
  const assignedName = resolveAssigned(shiftId, options);
  if (!assignedName) return "-";
  return `${slotOfUser(assignedName)} • ${assignedName}`;
}

function getPrimaryAbsenceInfo(shiftId, options) {
  const date = shiftId.slice(0, 10);
  const defaultAssigned = chooseDefault(options);
  if (!defaultAssigned) return null;
  const swappedDefault = applySwap(date, defaultAssigned);
  const type = getCalendarAbsenceType(swappedDefault, date);
  if (type) return { name: swappedDefault, type };
  if (state.absences[`${date}:${swappedDefault}`])
    return { name: swappedDefault, type: "abwesend" };
  return null;
}

function assignedMeta(shiftId, options) {
  const date = shiftId.slice(0, 10);
  const specialDay = getSpecialDay(date);

  if (specialDay) {
    const label =
      specialDay.type === "holiday"
        ? `FEIERTAG${specialDay.label ? ` – ${specialDay.label}` : ""}`
        : specialDay.type === "bridge"
          ? `BRÜCKENTAG${specialDay.label ? ` – ${specialDay.label}` : ""}`
          : `BETRIEBSFERIEN${specialDay.label ? ` – ${specialDay.label}` : ""}`;

    return {
      label,
      cls: "bg-red-100 text-red-900",
      borderCls: "border-red-500",
      ringCls: "border-2",
    };
  }

  if (
    state.shiftCancellations?.[shiftId] ||
    getReplacementForShift(shiftId)?.mode === "cancel"
  ) {
    return {
      label: "AUSFALL",
      cls: "bg-rose-200 text-rose-900",
      borderCls: "border-rose-500",
      ringCls: "border-2",
    };
  }

  const shift = getShiftById(shiftId);
  const assignedName = shift?.assigned || resolveAssigned(shiftId, options);
  const absence = getPrimaryAbsenceInfo(shiftId, options);

  if (!assignedName) {
    const abs = absence?.type || resolveAbsenceType(shiftId, options);
    if (abs) {
      return {
        label: "-",
        cls: "bg-slate-100 text-slate-700",
        borderCls: absence?.name
          ? personBorderClass(absence.name)
          : "border-slate-400",
        ringCls: "border-2",
      };
    }

    return {
      label: "-",
      cls: "bg-slate-100 text-slate-700",
      borderCls: "border-slate-300",
      ringCls: "border",
    };
  }

  const assignedColors = personColorClasses(assignedName);

  if (absence?.name) {
    return {
      label: `${slotOfUser(assignedName)} • ${assignedName}`,
      cls: `${assignedColors.bg} ${assignedColors.text}`,
      borderCls: personBorderClass(absence.name),
      ringCls: "border-2",
    };
  }

  return {
    label: `${slotOfUser(assignedName)} • ${assignedName}`,
    cls: assignedColors.raw,
    borderCls: personBorderClass(assignedName),
    ringCls: "border",
  };
}

function renderOverviewPlan(weeksToShow = 12) {
  const today = new Date();
  const monday = new Date(today);
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
  const anchor = new Date(`${ROTATION_ANCHOR_MONDAY}T00:00:00`);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  const isAdmin = currentUser?.role === "admin";

  const dayNames = [
    "Montag",
    "Dienstag",
    "Mittwoch",
    "Donnerstag",
    "Freitag",
    "Samstag",
    "Sonntag",
  ];

  function renderShiftCell(shiftId, options) {
    const meta = assignedMeta(shiftId, options);
    const hasManualAssignment = !!state.assignments?.[shiftId];
    const hasManualCancellation = !!state.shiftCancellations?.[shiftId];
    const hasReplacement = !!state.absenceReplacements?.[shiftId];
    const isManual =
      hasManualAssignment || hasManualCancellation || hasReplacement;

    return `<div class="flex flex-col items-center gap-1">
      <span>${meta.label}</span>
      ${
        isAdmin && isManual
          ? `<button class="px-1.5 py-0.5 rounded bg-slate-800 text-white text-[10px] leading-none" onclick="clearManualShift('${shiftId}')">Zurück</button>`
          : ""
      }
    </div>`;
  }

  let body = "";

  for (let w = 0; w < weeksToShow; w++) {
    const weekStart = weekStartFromMonday(monday, w);
    const templateWeekIndex =
      ((Math.floor((weekStart - anchor) / msPerWeek) % 6) + 6) % 6;
    const template = WEEK_TEMPLATES[templateWeekIndex];

    if (w === 0) {
      const startSunday = new Date(weekStart);
      startSunday.setDate(startSunday.getDate() - 1);
      const prevTemplateIndex = (templateWeekIndex + 5) % 6;
      const prevTemplate = WEEK_TEMPLATES[prevTemplateIndex];
      const startShiftId = `${isoDate(startSunday)}-su-1`;

      body += `<tr class="border-b bg-amber-50">
        <td class="p-2"></td>
        <td class="p-2 font-semibold">Start (Sonntag)</td>
        <td class="p-2 bg-slate-50" colspan="4"></td>
        <td class="p-2 font-semibold text-center ${assignedMeta(startShiftId, prevTemplate.sunday[1].options).cls} ${assignedMeta(startShiftId, prevTemplate.sunday[1].options).borderCls} ${assignedMeta(startShiftId, prevTemplate.sunday[1].options).ringCls}">
          ${renderShiftCell(startShiftId, prevTemplate.sunday[1].options)}
        </td>
        <td class="p-2 font-semibold text-center">18:00-24:00</td>
      </tr>`;
    }

    dayNames.forEach((dayName, dayIndex) => {
      const date = new Date(weekStart);
      date.setDate(weekStart.getDate() + dayIndex);
      const weekLabel = dayIndex === 0 ? `Woche ${w + 1}` : "";

      if (dayIndex <= 4) {
        const s1 = template.mondayToFriday[0];
        const s2 = template.mondayToFriday[1];
        const s3 = template.mondayToFriday[2];
        const dateIso = isoDate(date);

        const id1 = `${dateIso}-mf-0`;
        const id2 = `${dateIso}-mf-1`;
        const id3 = `${dateIso}-mf-2`;

        const m1 = assignedMeta(id1, s1.options);
        const m2 = assignedMeta(id2, s2.options);
        const m3 = assignedMeta(id3, s3.options);

        body += `<tr class="border-b">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${renderShiftCell(id1, s1.options)}</td>
          <td class="p-2 font-semibold text-center bg-white">05:00-11:00</td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${renderShiftCell(id2, s2.options)}</td>
          <td class="p-2 font-semibold text-center bg-white">13:00-19:00</td>
          <td class="p-2 ${m3.cls} ${m3.borderCls} ${m3.ringCls} font-semibold text-center">${renderShiftCell(id3, s3.options)}</td>
          <td class="p-2 font-semibold text-center bg-white">21:00-03:00</td>
        </tr>`;
      } else if (dayIndex === 5) {
        const s1 = template.saturday[0];
        const s2 = template.saturday[1];
        const dateIso = isoDate(date);

        const id1 = `${dateIso}-sa-0`;
        const id2 = `${dateIso}-sa-1`;

        const m1 = assignedMeta(id1, s1.options);
        const m2 = assignedMeta(id2, s2.options);

        body += `<tr class="border-b bg-amber-50">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${renderShiftCell(id1, s1.options)}</td>
          <td class="p-2 font-semibold text-center bg-amber-100">05:00-11:00 (Sa Morgen)</td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${renderShiftCell(id2, s2.options)}</td>
          <td class="p-2 font-semibold text-center bg-amber-200">16:00-22:00 (Sa Abend)</td>
          <td class="p-2 bg-slate-50" colspan="2"></td>
        </tr>`;
      } else {
        const s1 = template.sunday[0];
        const s2 = template.sunday[1];
        const dateIso = isoDate(date);

        const id1 = `${dateIso}-su-0`;
        const id2 = `${dateIso}-su-1`;

        const m1 = assignedMeta(id1, s1.options);
        const m2 = assignedMeta(id2, s2.options);

        body += `<tr class="border-b bg-amber-50">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${renderShiftCell(id1, s1.options)}</td>
          <td class="p-2 font-semibold text-center bg-blue-100">06:00-12:00 (So Morgen)</td>
          <td class="p-2 bg-slate-50" colspan="2"></td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${renderShiftCell(id2, s2.options)}</td>
          <td class="p-2 font-semibold text-center bg-white">18:00-24:00</td>
        </tr>`;
      }
    });
  }

  return `<div class='overflow-auto max-h-[70vh] border rounded-lg'>
    <table class='min-w-[1200px] w-full text-sm'>
      <thead class='sticky top-0 bg-slate-200 z-10'>
        <tr>
          <th class='p-2 text-left'>Woche</th>
          <th class='p-2 text-left'>Tag</th>
          <th class='p-2 text-center'>S1</th>
          <th class='p-2 text-center'>Frühschicht</th>
          <th class='p-2 text-center'>S2</th>
          <th class='p-2 text-center'>Spätschicht</th>
          <th class='p-2 text-center'>S3</th>
          <th class='p-2 text-center'>Nachtschicht</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
  </div>`;
}

function renderSchedule() {
  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-1'>Gesamt-Schichtplan (Alt / deaktiviert)</h2>
    <p class='text-sm text-slate-500 mb-3'>Dieser alte Schichtplan bleibt vorerst erhalten, ist aber nicht Grundlage der neuen Personalverwaltung.</p>
    ${renderOverviewPlan(Math.ceil(90 / 7))}
  </div>`;
}

async function assignShift(shiftId) {
  if (!supabaseReady) return;

  const select = document.getElementById(`sel-${shiftId}`);
  if (!select) return;

  const name = select.value;
  if (!canAssignUserToShift(name, shiftId)) return;

  const shift = getShiftById(shiftId);
  if (!shift) return;

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: name,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler bei Zuweisung:", error);
    return alert(`Zuweisung fehlgeschlagen: ${error.message}`);
  }

  // falls vorher storniert → entfernen
  await supabaseClient
    .from("planner_shift_cancellations")
    .delete()
    .eq("shift_id", shiftId);

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = name;

  persist();
  render();
}

async function cancelShift(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return;

  const { error } = await supabaseClient
    .from("planner_shift_cancellations")
    .upsert(
      {
        shift_id: shiftId,
        shift_date: shift.date,
        created_by_employee_id: currentEmployeeRecord?.id || null,
      },
      { onConflict: "shift_id" },
    );

  if (error) {
    console.error("Fehler beim Stornieren:", error);
    return alert(`Stornierung fehlgeschlagen: ${error.message}`);
  }

  const { error: deleteError } = await supabaseClient
    .from("planner_assignments")
    .delete()
    .eq("shift_id", shiftId);

  if (deleteError) {
    console.error("Fehler beim Entfernen der Zuweisung:", deleteError);
    return alert(
      `Ausfall gespeichert, aber Zuweisung konnte nicht entfernt werden: ${deleteError.message}`,
    );
  }

  delete state.assignments[shiftId];
  state.shiftCancellations[shiftId] = true;

  persist();
  render();
}

async function clearManualShift(shiftId) {
  if (currentUser?.role !== "admin") return;
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) {
    return alert("Schicht konnte nicht gefunden werden.");
  }

  const ok = confirm(
    `Manuelle Planung für ${formatDateWithWeekday(shift.date)} – ${shift.label} zurücksetzen?`,
  );
  if (!ok) return;

  const [assignmentDelete, cancellationDelete, replacementDelete] =
    await Promise.all([
      supabaseClient
        .from("planner_assignments")
        .delete()
        .eq("shift_id", shiftId),
      supabaseClient
        .from("planner_shift_cancellations")
        .delete()
        .eq("shift_id", shiftId),
      supabaseClient
        .from("planner_absence_replacements")
        .delete()
        .eq("shift_id", shiftId),
    ]);

  const firstError =
    assignmentDelete.error ||
    cancellationDelete.error ||
    replacementDelete.error;

  if (firstError) {
    console.error("Fehler beim Zurücksetzen der Schicht:", firstError);
    return alert(
      `Schicht konnte nicht zurückgesetzt werden: ${firstError.message}`,
    );
  }

  delete state.assignments[shiftId];
  delete state.shiftCancellations[shiftId];
  if (state.absenceReplacements) {
    delete state.absenceReplacements[shiftId];
  }

  persist();
  render();
}

function renderMyShifts() {
  const isSpringerUser = isSpringer(currentUser.name);

  const shifts = generateThreeMonths()
    .filter((s) => s.assigned === currentUser.name)
    .slice(0, 180);

  const rows = shifts
    .map((s) => {
      const status =
        s.assigned === currentUser.name
          ? '<span class="text-emerald-700 font-semibold">Eingeplant</span>'
          : s.open
            ? '<span class="text-red-600 font-semibold">OFFEN</span>'
            : '<span class="text-emerald-700">Besetzt</span>';

      return `<tr class="border-b">
        <td class="p-2">${formatDateWithWeekday(s.date)}</td>
        <td class="p-2">${s.label}</td>
        <td class="p-2">${s.start}–${s.end}</td>
        <td class="p-2">${status}</td>
        <td class="p-2">
          ${
            !isSpringerUser
              ? `<button class='px-2 py-1 bg-amber-200 rounded mr-2' onclick="markAbsent('${s.id}','${s.date}','${currentUser.name}')">Abwesenheit</button>`
              : "-"
          }
        </td>
      </tr>`;
    })
    .join("");

  const weekendRows = isSpringerUser
    ? generateThreeMonths()
        .filter((s) => s.id.includes("-sa-1") || s.id.includes("-su-0"))
        .slice(0, 180)
        .map((s) => {
          const key = `${s.id}:${currentUser.name}`;
          const val = state.availability[key] || "";

          return `<tr class='border-b'>
            <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
            <td class='p-2'>${s.label}</td>
            <td class='p-2'>${s.start}–${s.end}</td>
            <td class='p-2'>
              <select class='border rounded p-1' onchange="setWeekendAvailability('${s.id}', '${currentUser.name}', '${s.date}', this.value)">
                <option value='' ${val === "" ? "selected" : ""}>-</option>
                <option value='yes' ${val === "yes" ? "selected" : ""}>Kann</option>
                <option value='no' ${val === "no" ? "selected" : ""}>Kann nicht</option>
              </select>
            </td>
          </tr>`;
        })
        .join("")
    : "";

  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-lg font-semibold mb-3'>Meine Schichten (90 Tage)</h2>
      <div class='overflow-auto border rounded-lg'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100'>
            <tr>
              <th class='p-2 text-left'>Datum</th>
              <th class='p-2 text-left'>Schicht</th>
              <th class='p-2 text-left'>Zeit</th>
              <th class='p-2 text-left'>Status</th>
              <th class='p-2 text-left'>Aktionen</th>
            </tr>
          </thead>
          <tbody>${rows || `<tr><td class='p-2' colspan='5'>Keine Schichten gefunden.</td></tr>`}</tbody>
        </table>
      </div>
    </div>

    ${
      isSpringerUser
        ? `<div>
            <h3 class='text-lg font-semibold mb-3'>Wochenend-Verfügbarkeit</h3>
            <div class='overflow-auto border rounded-lg'>
              <table class='w-full text-sm'>
                <thead class='bg-slate-100'>
                  <tr>
                    <th class='p-2 text-left'>Datum</th>
                    <th class='p-2 text-left'>Schicht</th>
                    <th class='p-2 text-left'>Zeit</th>
                    <th class='p-2 text-left'>Verfügbarkeit</th>
                  </tr>
                </thead>
                <tbody>${weekendRows || `<tr><td class='p-2' colspan='4'>Keine relevanten Wochenendschichten gefunden.</td></tr>`}</tbody>
              </table>
            </div>
          </div>`
        : ""
    }
  </div>`;
}

async function requestSaturdayEvening(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  const requestKey = `${shiftId}:${currentUser.name}`;

  const { error } = await supabaseClient
    .from("planner_saturday_requests")
    .upsert(
      {
        request_key: requestKey,
        shift_id: shiftId,
        shift_date: shift.date,
        user_name: currentUser.name,
      },
      { onConflict: "request_key" },
    );

  if (error) {
    console.error("Fehler beim Speichern der Samstags-Anfrage:", error);
    return alert(
      `Samstags-Anfrage konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.saturdayEveningRequests[requestKey] = true;
  persist();
  render();
}

async function setWeekendAvailability(shiftId, userName, date, status) {
  if (!supabaseReady) return;

  const availabilityKey = `${shiftId}:${userName}`;

  if (!status) {
    const { error } = await supabaseClient
      .from("planner_availability")
      .delete()
      .eq("availability_key", availabilityKey);

    if (error) {
      console.error("Fehler beim Löschen der Verfügbarkeit:", error);
      return alert(
        `Verfügbarkeit konnte nicht gelöscht werden: ${error.message}`,
      );
    }

    delete state.availability[availabilityKey];
    persist();
    render();
    return;
  }

  const { error } = await supabaseClient.from("planner_availability").upsert(
    {
      availability_key: availabilityKey,
      shift_id: shiftId,
      shift_date: date,
      user_name: userName,
      status,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "availability_key" },
  );

  if (error) {
    console.error("Fehler bei Verfügbarkeit:", error);
    return alert(
      `Verfügbarkeit konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.availability[availabilityKey] = status;
  persist();
  render();
}

async function loadAvailabilityFromSupabase() {
  const { data, error } = await supabaseClient
    .from("planner_availability")
    .select("*");

  if (error) {
    console.error("Fehler beim Laden der Verfügbarkeit:", error);
    return {};
  }

  const map = {};
  (data || []).forEach((row) => {
    map[row.availability_key] = row.status;
  });

  return map;
}

function renderPlanning() {
  const subTab = PLANNING_SUBTABS.some((t) => t.id === state.planningSubTab)
    ? state.planningSubTab
    : "personal";

  const subTabButtons = PLANNING_SUBTABS.map((tab) => {
    const active = subTab === tab.id;
    return `<button class='px-3 py-2 rounded border ${active ? "humbel-subtab-active" : "humbel-subtab"}' onclick="setPlanningSubTab('${tab.id}')">${tab.label}</button>`;
  }).join("");

  let content = "";
  if (subTab === "personal") content = renderPlanningPersonal();
  if (subTab === "abstinenz") content = renderPlanningAbstinenz();
  if (subTab === "wochenende") content = renderPlanningWochenende();
  if (subTab === "schichttausch") content = renderPlanningSchichttausch();

  return `<div class='bg-white rounded-xl shadow p-4'>
    <div class='flex items-center justify-between mb-4 gap-2 flex-wrap'>
      <div>
        <h2 class='text-lg font-semibold'>Planung (Admin)</h2>
        <p class='text-sm text-slate-500 mt-1'>Struktur: Unterregister für Personal, Abstinenz, Wochenendeinsätze und Schichttausch.</p>
      </div>
      <button class='px-2 py-1 rounded bg-red-700 text-white text-sm' onclick='resetPlanCurrentFuture()'>Gesamtplan zurücksetzen (aktuelle+zukünftige)</button>
    </div>
    <div class='flex gap-2 flex-wrap mb-5'>
      ${subTabButtons}
    </div>
    ${content}
  </div>`;
}

function ensurePersonnelPendingState() {
  if (!state.ui) state.ui = {};
  if (!state.ui.pendingEmployeeEdits) state.ui.pendingEmployeeEdits = {};
  if (!state.ui.pendingSlotAssignments) state.ui.pendingSlotAssignments = {};
  return state.ui;
}

function queueEmployeeEdit(id, field, value) {
  if (!id) return;
  const ui = ensurePersonnelPendingState();
  ui.pendingEmployeeEdits[id] = {
    ...(ui.pendingEmployeeEdits[id] || {}),
    [field]: value,
  };
  persist();
  render();
}

function queueEmployeeActive(id, isActive) {
  if (!id) return;
  const employee = state.employees?.[id];
  const name = employee?.display_name || employee?.name || "Mitarbeiter";
  if (!isActive && !confirm(`${name} als inaktiv vormerken?`)) return;
  queueEmployeeEdit(id, "is_active", isActive);
}

function queueSlotAssignment(slot, value) {
  const ui = ensurePersonnelPendingState();
  ui.pendingSlotAssignments[slot] = value;
  persist();
  render();
}

function hasPersonnelPendingChanges() {
  const ui = ensurePersonnelPendingState();
  return (
    Object.keys(ui.pendingEmployeeEdits).length > 0 ||
    Object.keys(ui.pendingSlotAssignments).length > 0
  );
}

function renderPlanningPersonal() {
  const ui = ensurePersonnelPendingState();
  const pendingEmployees = ui.pendingEmployeeEdits;
  const pendingSlots = ui.pendingSlotAssignments;
  const hasPending = hasPersonnelPendingChanges();

  const personnelCards = allUsers()
    .map((u) => {
      const pending = pendingEmployees[u.id] || {};
      const name = pending.display_name ?? u.name;
      const type = pending.employee_type ?? u.type;
      const isActive =
        pending.is_active ??
        (u.isActive !== false && !state.inactiveUsers?.[u.name]);
      const colorKey = pending.color_key ?? u.color_key ?? "gray";
      const role = pending.role ?? u.role ?? "employee";
      const changed = Object.keys(pending).length > 0;
      const cardClass = changed
        ? "border-amber-300 bg-amber-50"
        : "border-slate-200 bg-white";
      const statusClass = isActive
        ? "bg-emerald-100 text-emerald-800"
        : "bg-slate-200 text-slate-700";

      return `<div class='border ${cardClass} rounded-lg p-3 shadow-sm'>
        <div class='flex items-start justify-between gap-3 mb-3'>
          <div class='min-w-0'>
            <div class='font-semibold text-slate-900 truncate'>${escapeHtml(name)}</div>
            ${changed ? `<div class='text-xs text-amber-700 mt-1'>Ungespeichert</div>` : ""}
          </div>
          <span class='shrink-0 px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${isActive ? "Aktiv" : "Inaktiv"}</span>
        </div>
        <div class='grid sm:grid-cols-2 gap-3'>
          <label class='text-xs text-slate-500'>Name
            <input id='employee-name-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' value='${escapeHtml(name)}' onchange="queueEmployeeEdit('${u.id}', 'display_name', this.value.trim())" />
          </label>
          <label class='text-xs text-slate-500'>Typ
            <select id='employee-type-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' onchange="queueEmployeeEdit('${u.id}', 'employee_type', this.value)">
              <option value='core' ${type === "core" ? "selected" : ""}>A/B/C</option>
              <option value='springer' ${type === "springer" ? "selected" : ""}>Springer</option>
            </select>
          </label>
          <label class='text-xs text-slate-500'>Farbe
            <select id='employee-color-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' onchange="queueEmployeeEdit('${u.id}', 'color_key', this.value)">
              ${EMPLOYEE_COLOR_OPTIONS.map((color) => `<option value='${color.key}' ${colorKey === color.key ? "selected" : ""}>${color.label}</option>`).join("")}
            </select>
          </label>
          <label class='text-xs text-slate-500'>Rolle
            <select id='employee-role-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' onchange="queueEmployeeEdit('${u.id}', 'role', this.value)">
              <option value='employee' ${role === "employee" ? "selected" : ""}>Mitarbeiter</option>
              <option value='admin' ${role === "admin" ? "selected" : ""}>Admin</option>
            </select>
          </label>
        </div>
        <div class='flex justify-end mt-3'>
          ${
            isActive
              ? `<button class='px-3 py-2 rounded bg-rose-700 text-white text-sm' onclick="queueEmployeeActive('${u.id}', false)">Inaktiv vormerken</button>`
              : `<button class='px-3 py-2 rounded bg-emerald-700 text-white text-sm' onclick="queueEmployeeActive('${u.id}', true)">Reaktivieren vormerken</button>`
          }
        </div>
      </div>`;
    })
    .join("");

  const slotAssignmentRows = SLOT_CODES.map((slot) => {
    const currentValue = state.slotAssignments?.[slot] || "";
    const selectedValue = pendingSlots[slot] ?? currentValue;
    const changed = Object.prototype.hasOwnProperty.call(pendingSlots, slot);
    const rowClass = changed ? "border-b bg-amber-50" : "border-b";
    const options = activeUsers()
      .map(
        (u) =>
          `<option value='${escapeHtml(u.name)}' ${selectedValue === u.name ? "selected" : ""}>${escapeHtml(u.name)}</option>`,
      )
      .join("");
    return `<tr class='${rowClass}'>
      <td class='p-2 font-semibold'>${slot}</td>
      <td class='p-2'><select id='slot-${slot}' class='border rounded p-1 w-full' onchange="queueSlotAssignment('${slot}', this.value)">${options}</select></td>
      <td class='p-2 text-xs text-amber-700'>${changed ? "ungespeichert" : ""}</td>
    </tr>`;
  }).join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <div class='flex items-center justify-between gap-2 mb-3'>
        <h3 class='font-semibold'>Personalverwaltung</h3>
        ${helpButton("planningPersonal")}
      </div>
      <div class='border rounded-lg bg-white p-3 mb-4'>
        <div class='font-semibold text-sm mb-3'>Mitarbeiter hinzufügen</div>
        <div class='grid sm:grid-cols-2 lg:grid-cols-3 gap-3'>
          <input id='newEmployeeName' class='border rounded p-2' placeholder='Neuer Name' />
          <select id='newEmployeeType' class='border rounded p-2'><option value='springer'>Springer</option><option value='core'>A/B/C</option></select>
          <select id='newEmployeeSlot' class='border rounded p-2'>
            <option value=''>Kein Slot</option>
            ${SLOT_CODES.map((slot) => `<option value='${slot}'>${slot}</option>`).join("")}
          </select>
          <select id='newEmployeeColor' class='border rounded p-2'>
            ${EMPLOYEE_COLOR_OPTIONS.map((color) => `<option value='${color.key}'>${color.label}</option>`).join("")}
          </select>
          <select id='newEmployeeRole' class='border rounded p-2'><option value='employee'>Mitarbeiter</option><option value='admin'>Admin</option></select>
          <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addEmployee()'>Mitarbeiter hinzufügen</button>
        </div>
      </div>
      <div class='grid lg:grid-cols-2 gap-3'>
        ${personnelCards}
      </div>
    </div>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='font-semibold mb-2'>Zuordnung</h3>
      <p class='text-sm text-slate-500 mb-2'>Admin kann festlegen, welcher Mitarbeiter aktuell A/B/C/D/E/F ist. Der Schichtplan passt sich nach dem zentralen Speichern an.</p>
      <div class='overflow-auto max-h-[55vh]'>
        <table class='w-full text-sm'><thead class='bg-white sticky top-0'><tr><th class='p-2 text-left'>Slot</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Status</th></tr></thead>
        <tbody>${slotAssignmentRows}</tbody></table>
      </div>
    </div>
    <div class='flex items-center justify-end gap-3 border rounded-lg p-3 bg-white'>
      ${hasPending ? `<span class='text-sm text-amber-700'>Es gibt ungespeicherte Änderungen</span>` : `<span class='text-sm text-slate-500'>Keine ungespeicherten Änderungen</span>`}
      <button class='px-4 py-2 rounded ${hasPending ? "bg-slate-900 text-white" : "bg-slate-200 text-slate-500 cursor-not-allowed"}' ${hasPending ? "" : "disabled"} onclick='saveAllPersonnelChanges()'>Alle Änderungen speichern</button>
    </div>
  </div>`;
}

function setProductionStatus(message = "", isError = false) {
  state.ui = state.ui || {};
  if (isError) {
    state.ui.productionActionError = message;
    state.ui.productionActionMessage = "";
    return;
  }
  state.ui.productionActionMessage = message;
  state.ui.productionActionError = "";
}

function formatProductionSupabaseError(
  error,
  fallback,
  duplicateMessage = "Dieser Abteilungscode ist bereits vorhanden.",
) {
  const message = error?.message || "";
  const code = error?.code || "";
  const lowerMessage = message.toLowerCase();
  if (
    code === "23505" ||
    lowerMessage.includes("duplicate") ||
    lowerMessage.includes("unique")
  ) {
    return duplicateMessage;
  }
  if (code === "23503" || lowerMessage.includes("foreign key")) {
    return `${fallback}: Datensatz ist noch verknüpft.`;
  }
  if (
    lowerMessage.includes("relation") ||
    lowerMessage.includes("does not exist") ||
    lowerMessage.includes("schema cache")
  ) {
    return `${fallback}: Supabase-Tabelle oder Spalte fehlt. Bitte Datenbankstruktur prüfen.`;
  }
  return `${fallback}: ${message || "Unbekannter Supabase-Fehler"}`;
}

function getProductionStatusBanner() {
  const tableError = state.ui?.productionDepartmentsError;
  const machinesError = state.ui?.productionMachinesError;
  const countsError = state.ui?.productionCountsError;
  const ordersError = state.ui?.productionOrdersError;
  const stationsError = state.ui?.productionStationsError;
  const orderEmployeesError = state.ui?.productionOrderEmployeesError;
  const stationCountsError = state.ui?.productionStationCountsError;
  const stationEventsError = state.ui?.productionStationEventsError;
  const orderHistoryError = state.ui?.productionOrderHistoryError;
  const qaCausesError = state.ui?.productionQaCausesError;
  const checklistTemplatesError = state.ui?.productionChecklistTemplatesError;
  const orderChecklistError = state.ui?.productionOrderChecklistError;
  const actionError = state.ui?.productionActionError;
  const actionMessage = state.ui?.productionActionMessage;
  const errors = [];
  if (tableError) errors.push(`departments: ${tableError}`);
  if (machinesError) errors.push(`production_machines: ${machinesError}`);
  if (countsError) errors.push(`production_counts: ${countsError}`);
  if (ordersError) errors.push(`production_orders: ${ordersError}`);
  if (stationsError) errors.push(`production_order_stations: ${stationsError}`);
  if (orderEmployeesError) errors.push(`production_order_employees: ${orderEmployeesError}`);
  if (stationCountsError) errors.push(`production_station_counts: ${stationCountsError}`);
  if (stationEventsError) errors.push(`production_station_events: ${stationEventsError}`);
  if (orderHistoryError) errors.push(`production_order_history: ${orderHistoryError}`);
  if (qaCausesError) errors.push(`production_qa_causes: ${qaCausesError}`);
  if (checklistTemplatesError) errors.push(`production_checklist_templates: ${checklistTemplatesError}`);
  if (orderChecklistError) errors.push(`production_order_checklist: ${orderChecklistError}`);
  if (actionError) errors.push(actionError);

  const errorBanner = errors.length
    ? `<div class='rounded border border-rose-200 bg-rose-50 p-3 text-sm text-rose-700'>${errors.map(escapeHtml).join(" | ")}</div>`
    : "";
  const messageBanner = actionMessage
    ? `<div class='rounded border border-emerald-200 bg-emerald-50 p-3 text-sm text-emerald-700'>${escapeHtml(actionMessage)}</div>`
    : "";

  return `${errorBanner}${messageBanner}`;
}

function getActiveDepartmentLeaderOptions(selectedId = "") {
  const activeEmployees = (state.employeesList || []).filter(
    (employee) => employee.is_active !== false,
  );
  const selectedEmployee = selectedId
    ? (state.employeesList || []).find((employee) => employee.id === selectedId)
    : null;
  const selectedIsActive =
    !!selectedEmployee && selectedEmployee.is_active !== false;
  const selectedInactiveOption =
    selectedEmployee && !selectedIsActive
      ? `<option value='${escapeHtml(selectedEmployee.id)}' selected>${escapeHtml(`${selectedEmployee.display_name || selectedEmployee.name || selectedEmployee.id} (inaktiv)`)}</option>`
      : "";
  const options = [
    `<option value='' ${!selectedId ? "selected" : ""}>Kein Abteilungsleiter</option>`,
    selectedInactiveOption,
    ...activeEmployees.map((employee) => {
      const label =
        employee.display_name ||
        [employee.first_name, employee.last_name].filter(Boolean).join(" ") ||
        employee.personnel_no ||
        employee.id;
      return `<option value='${escapeHtml(employee.id)}' ${selectedId === employee.id ? "selected" : ""}>${escapeHtml(label)}</option>`;
    }),
  ];

  return {
    activeEmployees,
    html: options.join(""),
  };
}

function getEmployeeDisplayNameById(employeeId) {
  if (!employeeId) return "-";
  const employee = (state.employeesList || []).find((entry) => entry.id === employeeId);
  return employee?.display_name || employee?.name || employeeId;
}

function getActiveDepartments() {
  return (state.departments || []).filter(
    (department) => department.active !== false,
  );
}

function getDepartmentById(departmentId) {
  if (!departmentId) return null;
  return (state.departments || []).find(
    (department) => department.id === departmentId,
  ) || null;
}

function getManagedDepartmentIds() {
  if (currentUser?.role === "admin") {
    return (state.departments || []).map((department) => department.id);
  }
  if (currentUser?.role !== "department_admin" || !currentEmployeeRecord?.id) {
    return [];
  }
  return (state.departments || [])
    .filter((department) => department.leader_employee_id === currentEmployeeRecord.id)
    .map((department) => department.id);
}

function getVisibleProductionDepartments() {
  if (currentUser?.role === "admin") return state.departments || [];
  const managedDepartmentIds = new Set(getManagedDepartmentIds());
  return (state.departments || []).filter((department) =>
    managedDepartmentIds.has(department.id),
  );
}

function getVisibleProductionMachines() {
  if (currentUser?.role === "admin") return state.productionMachines || [];
  const managedDepartmentIds = new Set(getManagedDepartmentIds());
  return (state.productionMachines || []).filter(
    (machine) => machine.department_id && managedDepartmentIds.has(machine.department_id),
  );
}

function getVisibleProductionCounts() {
  if (currentUser?.role === "admin") return state.productionCounts || [];
  const managedDepartmentIds = new Set(getManagedDepartmentIds());
  return (state.productionCounts || []).filter(
    (count) => count.department_id && managedDepartmentIds.has(count.department_id),
  );
}

function getVisibleProductionOrders() {
  if (currentUser?.role === "admin") return state.productionOrders || [];
  const managedDepartmentIds = new Set(getManagedDepartmentIds());
  return (state.productionOrders || []).filter(
    (order) => order.department_id && managedDepartmentIds.has(order.department_id),
  );
}

function getActiveProductionCountMachines() {
  return getVisibleProductionMachines().filter(
    (machine) => machine.active !== false,
  );
}

function getActiveProductionOrderMachines() {
  return getVisibleProductionMachines().filter(
    (machine) => machine.active !== false,
  );
}

function getProductionMachineById(machineId) {
  if (!machineId) return null;
  return (state.productionMachines || []).find(
    (machine) => machine.id === machineId,
  ) || null;
}

function getProductionCountById(countId) {
  if (!countId) return null;
  return (state.productionCounts || []).find((count) => count.id === countId) || null;
}

function getProductionOrderById(orderId) {
  if (!orderId) return null;
  return (state.productionOrders || []).find((order) => order.id === orderId) || null;
}

function canEditProductionOrder(order) {
  if (!canAccessProduction() || !order) return false;
  if (currentUser?.role === "admin") return true;
  if (currentUser?.role !== "department_admin") return false;
  return getManagedDepartmentIds().includes(order.department_id);
}

function canEditProductionCount(count) {
  if (!canAccessProduction() || !count) return false;
  if (currentUser?.role === "admin") return true;
  if (currentUser?.role !== "department_admin" || count.status === "completed") {
    return false;
  }
  return getManagedDepartmentIds().includes(count.department_id);
}

function renderProductionMachineOptions(selectedId = "") {
  const machines = getActiveProductionCountMachines();
  const selectedMachine = getProductionMachineById(selectedId);
  const selectedIsActive =
    !!selectedMachine && selectedMachine.active !== false;
  const selectedInactiveOption =
    selectedMachine && !selectedIsActive
      ? `<option value='${escapeHtml(selectedMachine.id)}' selected>${escapeHtml(`${selectedMachine.name || selectedMachine.machine_code || selectedMachine.id} (inaktiv)`)}</option>`
      : "";
  const options = [
    `<option value='' ${!selectedId ? "selected" : ""}>Maschine auswählen</option>`,
    selectedInactiveOption,
    ...machines.map((machine) => {
      const department = getDepartmentById(machine.department_id);
      const departmentLabel =
        department?.name || department?.code || (machine.department_id ? machine.department_id : "ohne Abteilung");
      const label = `${machine.name || machine.machine_code || machine.id} - ${departmentLabel}`;
      return `<option value='${escapeHtml(machine.id)}' ${selectedId === machine.id ? "selected" : ""}>${escapeHtml(label)}</option>`;
    }),
  ];
  return options.join("");
}

function renderProductionOrderMachineOptions(selectedId = "") {
  const machines = getActiveProductionOrderMachines();
  const selectedMachine = getProductionMachineById(selectedId);
  const selectedIsActive =
    !!selectedMachine && selectedMachine.active !== false;
  const selectedInactiveOption =
    selectedMachine && !selectedIsActive
      ? `<option value='${escapeHtml(selectedMachine.id)}' selected>${escapeHtml(`${selectedMachine.name || selectedMachine.machine_code || selectedMachine.id} (inaktiv)`)}</option>`
      : "";
  const options = [
    `<option value='' ${!selectedId ? "selected" : ""}>Maschine auswählen</option>`,
    selectedInactiveOption,
    ...machines.map((machine) => {
      const department = getDepartmentById(machine.department_id);
      const departmentLabel =
        department?.name || department?.code || (machine.department_id ? machine.department_id : "ohne Abteilung");
      const label = `${machine.name || machine.machine_code || machine.id} - ${departmentLabel}`;
      return `<option value='${escapeHtml(machine.id)}' ${selectedId === machine.id ? "selected" : ""}>${escapeHtml(label)}</option>`;
    }),
  ];
  return options.join("");
}

function renderProductionCountEmployeeOptions(selectedId = "") {
  const selectedEmployeeId = selectedId || currentEmployeeRecord?.id || "";
  const activeEmployees = (state.employeesList || []).filter(
    (employee) => employee.is_active !== false,
  );
  const selectedEmployee =
    selectedEmployeeId && !activeEmployees.some((employee) => employee.id === selectedEmployeeId)
      ? (state.employeesList || []).find((employee) => employee.id === selectedEmployeeId)
      : null;
  const selectedInactiveOption =
    selectedEmployee
      ? `<option value='${escapeHtml(selectedEmployee.id)}' selected>${escapeHtml(`${selectedEmployee.display_name || selectedEmployee.name || selectedEmployee.id} (inaktiv)`)}</option>`
      : "";
  const options = [
    selectedInactiveOption,
    ...activeEmployees.map((employee) => {
      const label =
        employee.display_name ||
        [employee.first_name, employee.last_name].filter(Boolean).join(" ") ||
        employee.personnel_no ||
        employee.id;
      return `<option value='${escapeHtml(employee.id)}' ${selectedEmployeeId === employee.id ? "selected" : ""}>${escapeHtml(label)}</option>`;
    }),
  ];
  return options.join("");
}

function getProductionCountDepartmentDisplay(machineId) {
  const machine = getProductionMachineById(machineId);
  const department = getDepartmentById(machine?.department_id);
  return department?.name || department?.code || (machine?.department_id ? machine.department_id : "-");
}

function getProductionOrderDepartmentDisplay(machineId, departmentId = "") {
  const machine = getProductionMachineById(machineId);
  const department = getDepartmentById(machine?.department_id || departmentId);
  return department?.name || department?.code || (machine?.department_id || departmentId ? machine?.department_id || departmentId : "-");
}

function renderProductionDepartmentAdminNotice() {
  if (currentUser?.role !== "department_admin") return "";
  const managedDepartmentIds = getManagedDepartmentIds();
  if (!managedDepartmentIds.length) {
    return `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Keine Abteilung als Abteilungsleiter zugeordnet.</div>`;
  }
  return `<div class='rounded border border-sky-200 bg-sky-50 p-3 text-sm text-sky-800'>Abteilungsleiter-Ansicht: nur eigene Abteilung</div>`;
}

function getDepartmentsLedByEmployee(employeeId) {
  if (!employeeId) return [];
  return (state.departments || []).filter(
    (department) => department.leader_employee_id === employeeId,
  );
}

function getSingleLedDepartmentId(employeeId) {
  const ledDepartments = getDepartmentsLedByEmployee(employeeId);
  return ledDepartments.length === 1 ? ledDepartments[0].id : "";
}

function getEmployeeEffectiveDepartmentId(employee) {
  if (employee?.department_id) return employee.department_id;
  return getSingleLedDepartmentId(employee?.id);
}

function renderDepartmentOptions(selectedId = "") {
  const activeDepartments = getActiveDepartments();
  const selectedDepartment = getDepartmentById(selectedId);
  const selectedIsActive =
    !!selectedDepartment && selectedDepartment.active !== false;
  const selectedInactiveOption =
    selectedDepartment && !selectedIsActive
      ? `<option value='${escapeHtml(selectedDepartment.id)}' selected>${escapeHtml(`${selectedDepartment.name || selectedDepartment.code || selectedDepartment.id} (inaktiv)`)}</option>`
      : "";
  const options = [
    `<option value='' ${!selectedId ? "selected" : ""}>Keine Abteilung</option>`,
    selectedInactiveOption,
    ...activeDepartments.map((department) => {
      const label = department.name || department.code || department.id;
      return `<option value='${escapeHtml(department.id)}' ${selectedId === department.id ? "selected" : ""}>${escapeHtml(label)}</option>`;
    }),
  ];
  return options.join("");
}

function getDepartmentDisplayForEmployee(employee) {
  const department = getDepartmentById(employee?.department_id);
  if (department) return department.name || department.code || department.id;

  if (!employee?.department_id) {
    const ledDepartments = getDepartmentsLedByEmployee(employee?.id);
    if (ledDepartments.length === 1) {
      return ledDepartments[0].name || ledDepartments[0].code || ledDepartments[0].id;
    }
    if (employee?.department) return employee.department;
  }

  return employee?.department || "-";
}

function getDepartmentLeaderHint(employee) {
  const ledDepartments = getDepartmentsLedByEmployee(employee?.id);
  if (!ledDepartments.length) return "";
  if (ledDepartments.length > 1) {
    return "Abteilungsleiter mehrerer Abteilungen - keine automatische Zuordnung";
  }
  const departmentName =
    ledDepartments[0].name || ledDepartments[0].code || ledDepartments[0].id;
  return `Abteilungsleiter: ${departmentName}`;
}

function renderProduction() {
  if (!canAccessProduction()) {
    return `<div class='bg-white rounded-xl shadow p-4'>
      <h2 class='text-lg font-semibold'>Produktion</h2>
      <p class='text-sm text-slate-600 mt-2'>Dieser Bereich ist für deine Rolle nicht verfügbar.</p>
    </div>`;
  }

  const subTab = PRODUCTION_SUBTABS.some((tab) => tab.id === state.productionSubTab)
    ? state.productionSubTab
    : "departments";
  const subTabButtons = PRODUCTION_SUBTABS.map((tab) => {
    const active = subTab === tab.id;
    return `<button class='px-3 py-2 rounded border ${active ? "humbel-subtab-active" : "humbel-subtab"}' onclick="setProductionSubTab('${tab.id}')">${tab.label}</button>`;
  }).join("");

  let content = "";
  if (subTab === "departments") content = renderProductionDepartmentsTab();
  if (subTab === "machines") content = renderProductionMachinesTab();
  if (subTab === "orders") content = renderProductionOrdersTab();
  if (subTab === "counts") content = renderProductionCountsTab();
  if (subTab === "setups") {
    content = renderModulePlaceholder(
      "Spannungen",
      "Spannungen werden später Maschinen und Aufträgen zugeordnet.",
    );
  }
  if (subTab === "scrap") {
    content = renderModulePlaceholder(
      "Ausschuss / Abklärung",
      "Ausschuss-Abklärung wird später mit Stückzahl und Protokoll verbunden.",
    );
  }
  if (subTab === "protocol") {
    content = renderProductionProtocolTab();
  }
  if (subTab === "settings") content = renderProductionSettingsTab();

  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-lg font-semibold'>Produktion</h2>
      <p class='text-sm text-slate-500 mt-1'>Grundstruktur für Abteilungen und spätere Maschinenverwaltung.</p>
    </div>
    <div class='flex gap-2 flex-wrap'>${subTabButtons}</div>
    ${getProductionStatusBanner()}
    ${renderProductionDepartmentAdminNotice()}
    ${content}
    ${renderProductionQaCauseModal()}
    ${renderProductionCompletionDialog()}
  </div>`;
}

function renderProductionDepartmentsTab() {
  const canEdit = canManageProduction();
  const leaderOptions = getActiveDepartmentLeaderOptions();
  const employeeInfo = !leaderOptions.activeEmployees.length
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Keine aktiven Mitarbeiter für die Auswahl als Abteilungsleiter vorhanden.</div>`
    : "";
  const createForm = canEdit
    ? `<div class='border rounded-lg p-3 bg-slate-50 space-y-3'>
        <h3 class='font-semibold'>Abteilung anlegen</h3>
        ${employeeInfo}
        <div class='grid sm:grid-cols-2 lg:grid-cols-5 gap-3'>
          <input id='productionNewDepartmentName' class='border rounded p-2 bg-white' placeholder='Name' />
          <input id='productionNewDepartmentCode' class='border rounded p-2 bg-white' placeholder='Code' />
          <select id='productionNewDepartmentLeader' class='border rounded p-2 bg-white'>${leaderOptions.html}</select>
          <label class='text-sm flex items-center gap-2 border rounded p-2 bg-white'>
            <input id='productionNewDepartmentActive' type='checkbox' checked />
            Aktiv
          </label>
          <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='createProductionDepartment()'>Anlegen</button>
        </div>
      </div>`
    : `<div class='rounded border border-slate-200 bg-slate-50 p-3 text-sm text-slate-600'>Du kannst die Produktionsstruktur sehen. Bearbeiten ist in dieser Version nur für Administratoren freigegeben.</div>`;

  const rows = getVisibleProductionDepartments()
    .map((department) => {
      const rowLeaderOptions = getActiveDepartmentLeaderOptions(
        department.leader_employee_id,
      );
      const active = department.active !== false;
      const statusClass = active
        ? "bg-emerald-100 text-emerald-800"
        : "bg-slate-200 text-slate-700";
      const leaderDisplay = getEmployeeDisplayNameById(
        department.leader_employee_id,
      );
      const actionButtons = canEdit
        ? `<button class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="saveProductionDepartment('${department.id}')">Speichern</button>
          ${
            active
              ? `<button class='px-3 py-2 rounded bg-rose-700 text-white text-sm ml-2' onclick="deactivateProductionDepartment('${department.id}')">Deaktivieren</button>`
              : `<button class='px-3 py-2 rounded bg-emerald-700 text-white text-sm ml-2' onclick="activateProductionDepartment('${department.id}')">Aktivieren</button>`
          }
          <button class='px-3 py-2 rounded bg-red-800 text-white text-sm ml-2' onclick="deleteProductionDepartment('${department.id}')">Löschen</button>`
        : "-";

      return `<tr class='border-b align-top ${active ? "" : "bg-slate-50 text-slate-500"}'>
        <td class='p-2'>
          ${
            canEdit
              ? `<input id='production-department-name-${department.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(department.name)}' />`
              : escapeHtml(department.name || "-")
          }
        </td>
        <td class='p-2'>
          ${
            canEdit
              ? `<input id='production-department-code-${department.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(department.code)}' />`
              : escapeHtml(department.code || "-")
          }
        </td>
        <td class='p-2'>
          ${
            canEdit
              ? `<select id='production-department-leader-${department.id}' class='border rounded p-2 w-full bg-white'>${rowLeaderOptions.html}</select>`
              : escapeHtml(leaderDisplay)
          }
        </td>
        <td class='p-2 whitespace-nowrap'>
          <span class='px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${active ? "Aktiv" : "Inaktiv"}</span>
        </td>
        <td class='p-2 whitespace-nowrap'>${actionButtons}</td>
      </tr>`;
    })
    .join("");

  return `<div class='space-y-4'>
    ${createForm}
    <div class='border rounded-lg bg-white overflow-auto'>
      <table class='w-full text-sm min-w-[850px]'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Name</th>
            <th class='p-2 text-left'>Code</th>
            <th class='p-2 text-left'>Abteilungsleiter</th>
            <th class='p-2 text-left'>Status</th>
            <th class='p-2 text-left'>Aktion</th>
          </tr>
        </thead>
        <tbody>${rows || "<tr><td class='p-3 text-slate-500' colspan='5'>Keine Abteilungen geladen.</td></tr>"}</tbody>
      </table>
    </div>
  </div>`;
}

function renderProductionMachinesTab() {
  const canEdit = canManageProduction();
  const activeDepartments = getActiveDepartments();
  const departmentInfo = !activeDepartments.length
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Keine aktiven Abteilungen für die Auswahl vorhanden. Maschinen können ohne Abteilung angelegt werden.</div>`
    : "";
  const createForm = canEdit
    ? `<div class='border rounded-lg p-3 bg-slate-50 space-y-3'>
        <h3 class='font-semibold'>Maschine anlegen</h3>
        ${departmentInfo}
        <div class='grid sm:grid-cols-2 lg:grid-cols-5 gap-3'>
          <input id='productionNewMachineName' class='border rounded p-2 bg-white' placeholder='Name' />
          <input id='productionNewMachineCode' class='border rounded p-2 bg-white' placeholder='Maschinencode' />
          <select id='productionNewMachineDepartment' class='border rounded p-2 bg-white'>${renderDepartmentOptions("")}</select>
          <label class='text-sm flex items-center gap-2 border rounded p-2 bg-white'>
            <input id='productionNewMachineActive' type='checkbox' checked />
            Aktiv
          </label>
          <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='createProductionMachine()'>Anlegen</button>
        </div>
      </div>`
    : `<div class='rounded border border-slate-200 bg-slate-50 p-3 text-sm text-slate-600'>Du kannst Maschinen sehen. Bearbeiten ist in dieser Version nur für Administratoren freigegeben.</div>`;

  const rows = getVisibleProductionMachines()
    .map((machine) => {
      const active = machine.active !== false;
      const statusClass = active
        ? "bg-emerald-100 text-emerald-800"
        : "bg-slate-200 text-slate-700";
      const department = getDepartmentById(machine.department_id);
      const departmentDisplay =
        department?.name || department?.code || (machine.department_id ? machine.department_id : "-");
      const actionButtons = canEdit
        ? `<button class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="saveProductionMachine('${machine.id}')">Speichern</button>
          ${
            active
              ? `<button class='px-3 py-2 rounded bg-rose-700 text-white text-sm ml-2' onclick="deactivateProductionMachine('${machine.id}')">Deaktivieren</button>`
              : `<button class='px-3 py-2 rounded bg-emerald-700 text-white text-sm ml-2' onclick="activateProductionMachine('${machine.id}')">Aktivieren</button>`
          }
          <button class='px-3 py-2 rounded bg-red-800 text-white text-sm ml-2' onclick="deleteProductionMachine('${machine.id}')">Löschen</button>`
        : "-";

      return `<tr class='border-b align-top ${active ? "" : "bg-slate-50 text-slate-500"}'>
        <td class='p-2'>
          ${
            canEdit
              ? `<input id='production-machine-name-${machine.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(machine.name)}' />`
              : escapeHtml(machine.name || "-")
          }
        </td>
        <td class='p-2'>
          ${
            canEdit
              ? `<input id='production-machine-code-${machine.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(machine.machine_code)}' />`
              : escapeHtml(machine.machine_code || "-")
          }
        </td>
        <td class='p-2'>
          ${
            canEdit
              ? `<select id='production-machine-department-${machine.id}' class='border rounded p-2 w-full bg-white'>${renderDepartmentOptions(machine.department_id)}</select>`
              : escapeHtml(departmentDisplay)
          }
        </td>
        <td class='p-2 whitespace-nowrap'>
          <span class='px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${active ? "Aktiv" : "Inaktiv"}</span>
        </td>
        <td class='p-2 whitespace-nowrap'>${actionButtons}</td>
      </tr>`;
    })
    .join("");

  return `<div class='space-y-4'>
    ${createForm}
    <div class='border rounded-lg bg-white overflow-auto'>
      <table class='w-full text-sm min-w-[850px]'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Name</th>
            <th class='p-2 text-left'>Maschinencode</th>
            <th class='p-2 text-left'>Abteilung</th>
            <th class='p-2 text-left'>Status</th>
            <th class='p-2 text-left'>Aktion</th>
          </tr>
        </thead>
        <tbody>${rows || "<tr><td class='p-3 text-slate-500' colspan='5'>Keine Maschinen geladen.</td></tr>"}</tbody>
      </table>
    </div>
  </div>`;
}

function renderProductionOrdersTab() {
  const canCreate = canAccessProduction();
  const machines = getActiveProductionOrderMachines();
  const machineInfo = !machines.length
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Keine aktiven Maschinen für Aufträge / BA verfügbar.</div>`
    : "";
  const createForm = canCreate
    ? `<div class='border rounded-lg p-3 bg-slate-50 space-y-3'>
        <h3 class='font-semibold'>Neuen Auftrag / BA anlegen</h3>
        ${machineInfo}
        <div class='grid md:grid-cols-2 lg:grid-cols-4 gap-3'>
          <label class='text-sm space-y-1'>
            <span class='font-medium'>Maschine</span>
            <select id='productionNewOrderMachine' class='border rounded p-2 bg-white w-full' onchange='updateProductionOrderDepartmentPreview("productionNewOrderMachine", "productionNewOrderDepartmentPreview")'>${renderProductionOrderMachineOptions("")}</select>
          </label>
          <label class='text-sm space-y-1'>
            <span class='font-medium'>Abteilung</span>
            <div id='productionNewOrderDepartmentPreview' class='border rounded p-2 bg-white text-slate-600 min-h-[42px]'>-</div>
          </label>
          <input id='productionNewOrderBaNumber' class='border rounded p-2 bg-white' placeholder='BA-Nummer' />
          <input id='productionNewOrderArticleNumber' class='border rounded p-2 bg-white' placeholder='Artikelnummer' />
          <input id='productionNewOrderBaQuantity' type='number' min='0' step='1' class='border rounded p-2 bg-white' placeholder='BA-Stückzahl' />
          <input id='productionNewOrderTargetQuantity' type='number' min='0' step='1' class='border rounded p-2 bg-white' placeholder='Zielstückzahl' />
          <input id='productionNewOrderPalletCount' type='number' min='0' step='1' class='border rounded p-2 bg-white' placeholder='Palettenanzahl' />
          <input id='productionNewOrderPiecesPerPallet' type='number' min='0' step='1' class='border rounded p-2 bg-white' placeholder='Stück pro Palette' />
          <label class='text-sm flex items-center gap-2 border rounded p-2 bg-white'>
            <input id='productionNewOrderUseChainLogic' type='checkbox' />
            Kettenlogik aktiv
          </label>
          <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='createProductionOrder()' ${machines.length ? "" : "disabled"}>Auftrag anlegen</button>
        </div>
      </div>`
    : `<div class='rounded border border-slate-200 bg-slate-50 p-3 text-sm text-slate-600'>Du kannst Aufträge sehen. Bearbeiten ist nur für freigegebene Produktionsrollen möglich.</div>`;

  const visibleOrders = getVisibleProductionOrders();
  const runningOrders = visibleOrders.filter((order) =>
    ["running", "paused"].includes(order.status),
  );
  const closedOrders = visibleOrders.filter((order) =>
    ["completed", "cancelled"].includes(order.status),
  );

  return `<div class='space-y-4'>
    ${createForm}
    ${renderProductionOrderTable("Laufende Aufträge", runningOrders, true)}
    ${renderProductionOrderTable("Abgeschlossene / abgebrochene Aufträge", closedOrders, false)}
  </div>`;
}

function renderProductionOrderTable(title, orders, allowActions) {
  const rows = orders.map((order) => renderProductionOrderRow(order, allowActions)).join("");
  return `<div class='border rounded-lg bg-white overflow-auto'>
    <div class='p-3 border-b bg-slate-50'>
      <h3 class='font-semibold'>${escapeHtml(title)}</h3>
    </div>
    <table class='w-full text-sm min-w-[1150px]'>
      <thead class='bg-slate-100 sticky top-0'>
        <tr>
          <th class='p-2 text-left'>Status</th>
          <th class='p-2 text-left'>Maschine</th>
          <th class='p-2 text-left'>Abteilung</th>
          <th class='p-2 text-left'>BA-Nummer</th>
          <th class='p-2 text-left'>Artikel</th>
          <th class='p-2 text-left'>BA-Stückzahl</th>
          <th class='p-2 text-left'>Ziel</th>
          <th class='p-2 text-left'>Paletten</th>
          <th class='p-2 text-left'>Stk./Palette</th>
          <th class='p-2 text-left'>Kette</th>
          <th class='p-2 text-left'>Gestartet</th>
          <th class='p-2 text-left'>Aktion</th>
        </tr>
      </thead>
      <tbody>${rows || `<tr><td class='p-3 text-slate-500' colspan='12'>Keine ${escapeHtml(title.toLowerCase())} geladen.</td></tr>`}</tbody>
    </table>
  </div>`;
}

function renderProductionOrderRow(order, allowActions) {
  const machine = getProductionMachineById(order.machine_id);
  const machineDisplay = machine?.name || machine?.machine_code || order.machine_id || "-";
  const departmentDisplay = getProductionOrderDepartmentDisplay(
    order.machine_id,
    order.department_id,
  );
  const canEdit = allowActions && canEditProductionOrder(order);
  const statusClass =
    order.status === "running"
      ? "bg-emerald-100 text-emerald-800"
      : order.status === "paused"
        ? "bg-amber-100 text-amber-800"
        : order.status === "completed"
          ? "bg-slate-200 text-slate-700"
          : "bg-rose-100 text-rose-700";
  const actionButtons = canEdit
    ? `<div class='flex gap-2 flex-wrap'>
        ${
          order.status === "running"
            ? `<button class='px-3 py-2 rounded bg-slate-700 text-white text-sm' onclick="pauseProductionOrder('${order.id}')">Pausieren</button>`
            : `<button class='px-3 py-2 rounded bg-emerald-700 text-white text-sm' onclick="resumeProductionOrder('${order.id}')">Fortsetzen</button>`
        }
        <button class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="completeProductionOrder('${order.id}')">Abschließen</button>
        <button class='px-3 py-2 rounded bg-rose-700 text-white text-sm' onclick="cancelProductionOrder('${order.id}')">Abbrechen</button>
      </div>`
    : "-";

  return `<tr class='border-b align-top ${order.status === "paused" ? "bg-amber-50" : ""}'>
    <td class='p-2 whitespace-nowrap'><span class='px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${escapeHtml(getProductionOrderStatusLabel(order.status))}</span></td>
    <td class='p-2'>${escapeHtml(machineDisplay)}</td>
    <td class='p-2'>${escapeHtml(departmentDisplay)}</td>
    <td class='p-2 font-semibold'>${escapeHtml(order.ba_number || "-")}</td>
    <td class='p-2'>${escapeHtml(order.article_number || "-")}</td>
    <td class='p-2'>${escapeHtml(order.ba_quantity)}</td>
    <td class='p-2'>${escapeHtml(order.target_quantity)}</td>
    <td class='p-2'>${escapeHtml(order.pallet_count)}</td>
    <td class='p-2'>${escapeHtml(order.pieces_per_pallet)}</td>
    <td class='p-2'>${order.use_chain_logic ? "Ja" : "Nein"}</td>
    <td class='p-2'>${escapeHtml(order.started_at ? new Date(order.started_at).toLocaleString() : "-")}</td>
    <td class='p-2 whitespace-nowrap'>${actionButtons}</td>
  </tr>`;
}

function getProductionOrderStatusLabel(status) {
  return {
    running: "Laufend",
    paused: "Pausiert",
    completed: "Abgeschlossen",
    cancelled: "Abgebrochen",
  }[status] || status || "-";
}

function getActiveProductionOrdersForMachine(machineId) {
  if (!machineId) return [];
  return getVisibleProductionOrders().filter(
    (order) =>
      order.machine_id === machineId &&
      ["running", "paused"].includes(order.status),
  );
}

function getProductionMachineDashboardStatus(machine) {
  const orders = getActiveProductionOrdersForMachine(machine.id);
  if (!orders.length) return "idle";
  if (orders.some((order) => order.status === "running")) return "active";
  return "paused";
}

function getProductionMachineDashboardStatusLabel(status) {
  return {
    idle: "Leerlauf",
    active: "Aktiv",
    paused: "Pausiert",
  }[status] || "Leerlauf";
}

function getProductionMachineDashboardStatusClass(status) {
  return {
    idle: "bg-slate-100 text-slate-700",
    active: "bg-emerald-100 text-emerald-800",
    paused: "bg-amber-100 text-amber-800",
  }[status] || "bg-slate-100 text-slate-700";
}

function getProductionOrderStations(orderId) {
  if (!orderId) return [];
  return (state.productionOrderStations || [])
    .filter((station) => station.order_id === orderId)
    .sort((a, b) => a.station_no - b.station_no);
}

function getProductionOrderStationById(stationId) {
  if (!stationId) return null;
  return (state.productionOrderStations || []).find((station) => station.id === stationId) || null;
}

function getProductionOrderEmployees(orderId, activeOnly = true) {
  if (!orderId) return [];
  return (state.productionOrderEmployees || [])
    .filter((entry) => entry.order_id === orderId)
    .filter((entry) => (activeOnly ? entry.active !== false : true))
    .sort((a, b) => a.sort_order - b.sort_order);
}

function getProductionOrderEmployeeById(entryId) {
  if (!entryId) return null;
  return (state.productionOrderEmployees || []).find((entry) => entry.id === entryId) || null;
}

function getProductionStationCount(stationId, orderEmployeeId) {
  if (!stationId || !orderEmployeeId) return null;
  return (state.productionStationCounts || []).find(
    (count) => count.station_id === stationId && count.order_employee_id === orderEmployeeId,
  ) || null;
}

function getProductionStationGoodQty(stationId, orderEmployeeId) {
  return getProductionStationCount(stationId, orderEmployeeId)?.good_qty || 0;
}

function getProductionStationGoodTotal(stationId) {
  if (!stationId) return 0;
  return (state.productionStationCounts || [])
    .filter((count) => count.station_id === stationId)
    .reduce((sum, count) => sum + Number(count.good_qty || 0), 0);
}

function getProductionOrderGoodTotal(orderId) {
  if (!orderId) return 0;
  return (state.productionStationCounts || [])
    .filter((count) => count.order_id === orderId)
    .reduce((sum, count) => sum + Number(count.good_qty || 0), 0);
}

function getProductionOrderScrapTotal(orderId) {
  return getProductionOrderStations(orderId).reduce(
    (sum, station) => sum + Number(station.scrap_total || 0),
    0,
  );
}

function getProductionOrderClarifyTotal(orderId) {
  return getProductionOrderStations(orderId).reduce(
    (sum, station) => sum + Number(station.clarify_total || 0),
    0,
  );
}

function getProductionOrderCompletedGoodTotal(orderId) {
  const stations = getProductionOrderStations(orderId);
  if (!stations.length) return 0;
  const lastStation = stations[stations.length - 1];
  return getProductionStationGoodTotal(lastStation.id);
}

function getProductionOrderChecklistItems(orderId) {
  if (!orderId) return [];
  return (state.productionOrderChecklist || [])
    .filter((item) => item.order_id === orderId)
    .sort((a, b) =>
      `${String(a.sort_order).padStart(6, "0")} ${a.item_label}`.localeCompare(
        `${String(b.sort_order).padStart(6, "0")} ${b.item_label}`,
        "de",
      ),
    );
}

function getProductionOrderChecklistItemById(itemId) {
  if (!itemId) return null;
  return (state.productionOrderChecklist || []).find((item) => item.id === itemId) || null;
}

function isProductionOrderChecklistComplete(orderId) {
  const items = getProductionOrderChecklistItems(orderId);
  return !!items.length && items.every((item) => item.checked === true);
}

function getAssignableProductionEmployees(orderId = "") {
  const assignedActiveIds = new Set(
    getProductionOrderEmployees(orderId, true)
      .map((entry) => entry.employee_id)
      .filter(Boolean),
  );
  return (state.employeesList || [])
    .filter((employee) => employee.is_active !== false)
    .filter((employee) => !["admin", "tool_scanner"].includes(employee.role))
    .filter((employee) => !assignedActiveIds.has(employee.id));
}

function getProductionOrderEmployeeRoleLabel(role) {
  return role === "worker" ? "Mitarbeiter" : role || "Mitarbeiter";
}

function getProductionEmployeeDisplayName(employee) {
  return (
    [employee?.first_name, employee?.last_name].filter(Boolean).join(" ").trim() ||
    employee?.display_name ||
    employee?.name ||
    employee?.personnel_no ||
    "Mitarbeiter"
  );
}

function calculatePreparedRemainingQuantity(order) {
  if (!order) return 0;
  const targetQuantity = Math.max(0, Math.trunc(Number(order.target_quantity || 0)));
  const finishedGoodQty =
    order.use_chain_logic !== false
      ? getProductionOrderCompletedGoodTotal(order.id)
      : getProductionOrderGoodTotal(order.id);
  return Math.max(0, targetQuantity - finishedGoodQty - getProductionOrderScrapTotal(order.id));
}

function getProductionStationTimeStatusLabel(status) {
  return status === "changed" ? "Zeit geändert" : "Zeit ok";
}

function logProductionStationSupabaseError(action, error, details = {}) {
  console.error("production_order_stations Supabase-Fehler:", {
    action,
    order_id: details.order_id || "",
    station_id: details.station_id || "",
    station_no: details.station_no || "",
    current_user_role: currentUser?.role || "",
    current_employee_role: currentEmployeeRecord?.role || "",
    error,
  });
}

function logProductionOrderChecklistSupabaseError(action, error, orderId = "") {
  console.error("production_order_checklist Supabase-Fehler:", {
    action,
    order_id: orderId || "",
    current_user_role: currentUser?.role || "",
    current_employee_role: currentEmployeeRecord?.role || "",
    error,
  });
}

async function refreshProductionOrderStationsFromSupabase() {
  const stations = await loadProductionOrderStationsFromSupabase();
  if (!Array.isArray(stations)) return false;
  applyProductionOrderStationsToState(stations);
  return true;
}

async function ensureProductionOrderStationsForOrder(orderId) {
  const order = getProductionOrderById(orderId);
  if (
    !order ||
    !order.id ||
    !["running", "paused"].includes(order.status) ||
    !canEditProductionOrder(order)
  ) return;
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    return;
  }
  if (state.ui?.productionStationsError) return;
  if (getProductionOrderStations(order.id).length) return;

  const { error } = await supabaseClient.from("production_order_stations").insert([
    {
      order_id: order.id,
      station_no: 1,
      name: "Spannung 1",
      lock_name: false,
      op_number: null,
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      updated_at: new Date().toISOString(),
    },
  ]);

  if (error) {
    logProductionStationSupabaseError("insert", error, {
      order_id: order.id,
      station_no: 1,
    });
    const duplicateCode =
      error.code === "23505" || String(error.message || "").toLowerCase().includes("duplicate");
    if (!duplicateCode) {
      setProductionStatus(
        formatProductionSupabaseError(
          error,
          "Erste Spannung konnte nicht angelegt werden",
        ),
        true,
      );
    }
    await refreshProductionOrderStationsFromSupabase();
    return;
  }

  await refreshProductionOrderStationsFromSupabase();
}

async function selectProductionCounterMachine(machineId) {
  state.ui = state.ui || {};
  state.ui.productionCounterSelectedMachineId = machineId || "";
  const activeOrders = getActiveProductionOrdersForMachine(machineId);
  state.ui.productionCounterActiveOrderId = activeOrders[0]?.id || "";
  persist();
  render();
  if (state.ui.productionCounterActiveOrderId) {
    await ensureProductionOrderStationsForOrder(state.ui.productionCounterActiveOrderId);
    await ensureProductionOrderChecklistForOrder(state.ui.productionCounterActiveOrderId);
    render();
  }
}

async function selectProductionCounterOrder(orderId) {
  state.ui = state.ui || {};
  const order = getProductionOrderById(orderId);
  if (!order || !["running", "paused"].includes(order.status)) return;
  state.ui.productionCounterSelectedMachineId = order.machine_id || "";
  state.ui.productionCounterActiveOrderId = order.id || "";
  persist();
  render();
  await ensureProductionOrderStationsForOrder(order.id);
  await ensureProductionOrderChecklistForOrder(order.id);
  render();
}

function resetProductionCounterSelection() {
  state.ui = state.ui || {};
  state.ui.productionCounterSelectedMachineId = "";
  state.ui.productionCounterActiveOrderId = "";
  persist();
  render();
}

function renderProductionCountsTab() {
  const machines = getVisibleProductionMachines().filter(
    (machine) => machine.active !== false,
  );
  const groupedMachines = new Map();
  machines.forEach((machine) => {
    const departmentKey = machine.department_id || "__none__";
    if (!groupedMachines.has(departmentKey)) groupedMachines.set(departmentKey, []);
    groupedMachines.get(departmentKey).push(machine);
  });

  const selectedMachineId = state.ui?.productionCounterSelectedMachineId || "";
  const selectedMachine = machines.find((machine) => machine.id === selectedMachineId);
  const selectedOrders = selectedMachine
    ? getActiveProductionOrdersForMachine(selectedMachine.id)
    : [];
  if (selectedMachine && selectedOrders.length) {
    return renderProductionMachineOrderPreview(selectedMachine, selectedOrders);
  }
  const firstMachineWithOrders = machines.find(
    (machine) => getActiveProductionOrdersForMachine(machine.id).length > 0,
  );
  const detailMachine = selectedMachine || firstMachineWithOrders || machines[0] || null;
  const detailOrders = detailMachine
    ? getActiveProductionOrdersForMachine(detailMachine.id)
    : [];
  const detailActive =
    selectedMachine && selectedOrders.length
      ? renderProductionCounterMachineDetail(selectedMachine, selectedOrders)
      : selectedMachine
        ? renderProductionCounterNoOrderNotice(selectedMachine)
        : detailMachine && detailOrders.length
          ? renderProductionCounterMachineDetail(detailMachine, detailOrders)
          : `<div class='border rounded-lg bg-slate-50 p-4 text-sm text-slate-600'>Bitte eine Maschinenkarte auswählen. Ohne laufenden Auftrag zuerst unter Aufträge / BA einen Auftrag anlegen.</div>`;

  const departmentSections = Array.from(groupedMachines.entries())
    .map(([departmentId, departmentMachines]) => {
      const department = getDepartmentById(departmentId === "__none__" ? "" : departmentId);
      const departmentName =
        department?.name || department?.code || (departmentId === "__none__" ? "Ohne Abteilung" : departmentId);
      const cards = departmentMachines
        .map((machine) => renderProductionCounterMachineCard(machine, selectedMachineId))
        .join("");
      return `<section class='space-y-3'>
        <div class='flex items-center justify-between gap-2 flex-wrap'>
          <h3 class='text-base font-semibold'>${escapeHtml(departmentName)}</h3>
          <span class='text-xs text-slate-500'>${departmentMachines.length} Maschine(n)</span>
        </div>
        <div class='grid md:grid-cols-2 xl:grid-cols-3 gap-3'>${cards}</div>
      </section>`;
    })
    .join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-4 bg-slate-50'>
      <h3 class='text-xl font-bold'>Zähler-Dashboard</h3>
      <p class='text-sm text-slate-600 mt-1'>Zählansicht wird schrittweise aus der alten Zählerapp migriert.</p>
      <p class='text-xs text-slate-500 mt-2'>Diese Version zeigt Maschinen und aktive Aufträge nur lesend. Keine Zählung, keine Spannungen, keine Buchung.</p>
    </div>
    <div class='grid xl:grid-cols-[minmax(0,2fr),minmax(360px,1fr)] gap-4 items-start'>
      <div class='space-y-5'>
        <div>
          <h3 class='font-semibold'>Maschinenübersicht</h3>
          <p class='text-sm text-slate-500 mt-1'>Maschinen gruppiert nach Abteilung mit laufenden und pausierten Aufträgen.</p>
        </div>
        ${departmentSections || "<div class='border rounded-lg bg-white p-4 text-sm text-slate-500'>Keine Maschinen sichtbar.</div>"}
      </div>
      <aside class='bg-white border rounded-lg p-4 space-y-3 sticky top-3'>
        <h3 class='font-semibold'>Maschinenvorschau</h3>
        ${detailActive}
      </aside>
    </div>
    <details class='border rounded-lg bg-slate-50 p-3 text-sm text-slate-600'>
      <summary class='font-semibold cursor-pointer'>Alt / einfacher Zähler zurückgestellt</summary>
      <p class='mt-2'>Der einfache production_counts-Zähler bleibt im Code erhalten, ist aber in dieser Ansicht nicht mehr die primäre Zähleroberfläche.</p>
    </details>
  </div>`;
}

function renderProductionCounterMachineCard(machine, selectedMachineId = "") {
  const orders = getActiveProductionOrdersForMachine(machine.id);
  const status = getProductionMachineDashboardStatus(machine);
  const departmentDisplay = getProductionOrderDepartmentDisplay(
    machine.id,
    machine.department_id,
  );
  const runningCount = orders.filter((order) => order.status === "running").length;
  const pausedCount = orders.filter((order) => order.status === "paused").length;
  const selectedClass = selectedMachineId === machine.id ? "ring-2 ring-[var(--humbel-blue)]" : "";
  const orderList = orders.length
    ? `<div class='space-y-2 mt-3'>
        ${orders.map((order) => renderProductionCounterOrderSummary(order)).join("")}
      </div>`
    : `<div class='mt-3 rounded border border-slate-200 bg-white p-3 text-sm text-slate-500'>Auftrag unter Aufträge / BA anlegen</div>`;
  return `<button type='button' class='text-left border rounded-lg bg-white p-4 shadow-sm hover:bg-slate-50 ${selectedClass}' onclick="selectProductionCounterMachine('${machine.id}')">
    <div class='flex items-start justify-between gap-3'>
      <div>
        <div class='font-semibold text-slate-900'>${escapeHtml(machine.name || machine.machine_code || "-")}</div>
        <div class='text-xs text-slate-500'>${escapeHtml(machine.machine_code || machine.id || "-")}</div>
        <div class='text-xs text-slate-500 mt-1'>${escapeHtml(departmentDisplay)}</div>
      </div>
      <span class='px-2 py-1 rounded-full text-xs font-semibold ${getProductionMachineDashboardStatusClass(status)}'>${getProductionMachineDashboardStatusLabel(status)}</span>
    </div>
    <div class='text-xs text-slate-600 mt-3'>${runningCount} laufend / ${pausedCount} pausiert</div>
    ${orderList}
  </button>`;
}

function renderProductionCounterOrderSummary(order) {
  return `<div class='rounded border border-slate-200 bg-slate-50 p-2 text-xs'>
    <div class='flex items-center justify-between gap-2'>
      <span class='font-semibold'>${escapeHtml(order.ba_number || "-")}</span>
      <span>${escapeHtml(getProductionOrderStatusLabel(order.status))}</span>
    </div>
    <div class='text-slate-500 mt-1'>Artikel: ${escapeHtml(order.article_number || "-")}</div>
    <div class='text-slate-500'>BA ${escapeHtml(order.ba_quantity)} / Ziel ${escapeHtml(order.target_quantity)}</div>
  </div>`;
}

function renderProductionMachineOrderPreview(machine, orders) {
  const departmentDisplay = getProductionOrderDepartmentDisplay(
    machine.id,
    machine.department_id,
  );
  const activeOrderId = state.ui?.productionCounterActiveOrderId || orders[0]?.id || "";
  const activeOrder =
    orders.find((order) => order.id === activeOrderId) || orders[0] || null;
  const baTabs = orders
    .map((order) => {
      const active = activeOrder?.id === order.id;
      return `<button type='button' class='px-3 py-2 rounded border ${active ? "humbel-subtab-active" : "humbel-subtab"}' onclick="selectProductionCounterOrder('${order.id}')">
        BA ${escapeHtml(order.ba_number || "-")}
      </button>`;
    })
    .join("");
  const diff = activeOrder
    ? Number(activeOrder.target_quantity || 0) - Number(activeOrder.ba_quantity || 0)
    : 0;
  const remaining = activeOrder
    ? calculatePreparedRemainingQuantity(activeOrder)
    : 0;
  const orderGoodTotal = activeOrder ? getProductionOrderGoodTotal(activeOrder.id) : 0;
  const orderScrapTotal = activeOrder ? getProductionOrderScrapTotal(activeOrder.id) : 0;
  const orderClarifyTotal = activeOrder ? getProductionOrderClarifyTotal(activeOrder.id) : 0;
  const metrics = activeOrder
    ? `<div class='grid sm:grid-cols-2 xl:grid-cols-4 gap-3'>
        ${renderProductionPreviewMetric("BA-Stückzahl", activeOrder.ba_quantity)}
        ${renderProductionPreviewMetric("Zielstückzahl", activeOrder.target_quantity)}
        ${renderProductionPreviewMetric("Differenz", diff)}
        ${renderProductionPreviewMetric("Gutteile Auftrag", orderGoodTotal)}
        ${renderProductionPreviewMetric("Ausschuss Auftrag", orderScrapTotal)}
        ${renderProductionPreviewMetric("In Abklärung Auftrag", orderClarifyTotal)}
        ${renderProductionPreviewMetric("Restmenge vorbereitet", remaining)}
      </div>`
    : "";
  const details = activeOrder
    ? `<div class='border rounded-lg bg-white p-4 space-y-4'>
        <div class='flex items-start justify-between gap-3 flex-wrap'>
          <div>
            <h3 class='text-xl font-bold'>BA ${escapeHtml(activeOrder.ba_number || "-")}</h3>
            <p class='text-sm text-slate-500 mt-1'>Artikel ${escapeHtml(activeOrder.article_number || "-")}</p>
          </div>
          <span class='px-2 py-1 rounded-full text-xs font-semibold ${activeOrder.status === "running" ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"}'>${escapeHtml(getProductionOrderStatusLabel(activeOrder.status))}</span>
        </div>
        ${metrics}
        ${renderProductionOrderEmployeesSection(activeOrder)}
        ${renderProductionOrderStationsSection(activeOrder)}
        ${renderProductionOrderChecklistSection(activeOrder, false)}
      </div>`
    : `<div class='border rounded-lg bg-slate-50 p-4 text-sm text-slate-600'>Kein aktiver BA ausgewählt.</div>`;

  return `<div class='space-y-4'>
    <button type='button' class='px-3 py-2 rounded bg-slate-200 text-slate-800' onclick='resetProductionCounterSelection()'>Zurück zur Maschinenübersicht</button>
    <div class='border rounded-lg p-4 bg-slate-50'>
      <div class='flex items-start justify-between gap-3 flex-wrap'>
        <div>
          <h3 class='text-xl font-bold'>Produktionsvorschau</h3>
          <p class='text-sm text-slate-600 mt-1'>${escapeHtml(machine.name || machine.machine_code || "-")} · ${escapeHtml(machine.machine_code || "-")}</p>
          <p class='text-sm text-slate-500 mt-1'>Abteilung: ${escapeHtml(departmentDisplay)}</p>
        </div>
        <span class='px-2 py-1 rounded-full text-xs font-semibold ${getProductionMachineDashboardStatusClass(getProductionMachineDashboardStatus(machine))}'>${getProductionMachineDashboardStatusLabel(getProductionMachineDashboardStatus(machine))}</span>
      </div>
    </div>
    <div class='bg-white border rounded-lg p-4 space-y-4'>
      <div>
        <h3 class='font-semibold'>Aktive BA</h3>
        <p class='text-sm text-slate-500 mt-1'>Nur laufende und pausierte Aufträge dieser Maschine.</p>
      </div>
      <div class='flex gap-2 flex-wrap'>${baTabs}</div>
      ${details}
    </div>
  </div>`;
}

function renderProductionOrderStationsSection(order) {
  const stations = getProductionOrderStations(order.id);
  const canEdit = canEditProductionOrder(order);
  const addButton = canEdit
    ? `<button type='button' class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="createProductionOrderStation('${order.id}')">Spannung hinzufügen</button>`
    : "";
  const stationCards = stations.length
    ? stations.map((station) => renderProductionOrderStationCard(station, canEdit)).join("")
    : `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Noch keine Spannung geladen. Beim Öffnen des BA wird Spannung 1 automatisch vorbereitet.</div>`;
  return `<section class='border rounded-lg bg-slate-50 p-4 space-y-4'>
    <div class='flex items-start justify-between gap-3 flex-wrap'>
      <div>
        <h4 class='font-semibold'>Spannungen</h4>
        <p class='text-sm text-slate-600 mt-1'>Gutteile werden je Mitarbeiter gezählt. Ausschuss und Abklärung werden je Spannung erfasst.</p>
      </div>
      ${addButton}
    </div>
    <div class='grid lg:grid-cols-2 gap-3'>${stationCards}</div>
  </section>`;
}

function renderProductionOrderChecklistSection(order, expanded = false) {
  const templates = state.productionChecklistTemplates || [];
  const checklistItems = getProductionOrderChecklistItems(order.id);
  const canEdit =
    canEditProductionOrder(order) &&
    ["running", "paused"].includes(order.status);
  const stations = getProductionOrderStations(order.id);
  const remaining = calculatePreparedRemainingQuantity(order);
  const checklistComplete = isProductionOrderChecklistComplete(order.id);
  const completing = state.ui?.productionOrderCompletingId === order.id;
  const preparing = state.ui?.productionChecklistPreparingOrderId === order.id;
  const needsPreparation = productionOrderChecklistNeedsPreparation(order);
  const missingTemplatesNotice = !templates.length
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Keine Abschluss-Checkliste eingerichtet.</div>`
    : "";
  const missingItemsNotice = needsPreparation
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800 flex items-center justify-between gap-3 flex-wrap'>
        <span>${preparing ? "Abschluss-Checkliste wird vorbereitet." : "Abschluss-Checkliste ist noch nicht vollständig vorbereitet."}</span>
        <button type='button' class='px-3 py-2 rounded bg-slate-900 text-white text-sm disabled:opacity-50' onclick="prepareProductionOrderChecklist('${order.id}')" ${canEdit && !preparing ? "" : "disabled"}>${preparing ? "Bitte warten..." : "Checkliste vorbereiten"}</button>
      </div>`
    : "";
  const itemsHtml = checklistItems.length
    ? `<div class='space-y-2'>${checklistItems.map((item) => renderProductionOrderChecklistItem(item, canEdit)).join("")}</div>`
    : "";
  const blockers = [];
  if (!templates.length) {
    blockers.push("Keine Abschluss-Checkliste eingerichtet.");
  } else if (needsPreparation) {
    blockers.push("Abschluss-Checkliste muss vorbereitet sein.");
  } else if (!checklistComplete) {
    blockers.push("Alle Checklistenpunkte müssen erledigt sein.");
  }
  if (!stations.length) blockers.push("Mindestens eine Spannung muss vorhanden sein.");
  if (Number(order.target_quantity || 0) > 0 && remaining > 0) {
    blockers.push("Auftrag hat noch Restmenge. Abschluss nicht möglich.");
  }
  const blockerHtml = blockers.length
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>${blockers.map(escapeHtml).join(" ")}</div>`
    : "";
  const completeButton =
    ["running", "paused"].includes(order.status)
      ? expanded
        ? `<button type='button' class='px-3 py-2 rounded bg-emerald-700 text-white text-sm disabled:opacity-50' onclick="completeProductionOrderWithChecklist('${order.id}')" ${canEdit && !completing && !blockers.length ? "" : "disabled"}>${completing ? "Fertigmeldung läuft..." : "Endgültig fertig melden"}</button>`
        : `<button type='button' class='px-3 py-2 rounded bg-emerald-700 text-white text-sm' onclick="openProductionCompletionDialog('${order.id}')">Auftrag fertig melden</button>`
      : "";
  const completedNotice =
    ["completed", "cancelled"].includes(order.status)
      ? `<div class='rounded border border-slate-200 bg-slate-50 p-3 text-sm text-slate-600'>Auftrag ist ${escapeHtml(getProductionOrderStatusLabel(order.status).toLowerCase())}; Checkliste ist nur lesbar.</div>`
      : "";

  if (!expanded) {
    const checklistStatus = !templates.length
      ? "Keine Checkliste eingerichtet"
      : checklistComplete
        ? "Checkliste erledigt"
        : "Checkliste offen";
    const canCompleteLabel = blockers.length ? "Abschluss aktuell nicht möglich" : "Abschluss möglich";
    const protocolButton = `<button type='button' class='px-3 py-2 rounded bg-slate-200 text-slate-800 text-sm' onclick="openProductionProtocolForOrder('${order.id}')">Protokoll anzeigen</button>`;
    return `<section class='border rounded-lg bg-slate-50 p-4 space-y-3'>
      <div class='flex items-start justify-between gap-3 flex-wrap'>
        <div>
          <h4 class='font-semibold'>Auftrag abschließen</h4>
          <p class='text-sm text-slate-600 mt-1'>Checkliste wird beim Fertigmelden geöffnet.</p>
        </div>
        <div class='flex gap-2 flex-wrap'>${protocolButton}${completeButton}</div>
      </div>
      <div class='grid sm:grid-cols-3 gap-2 text-sm'>
        <div class='rounded border bg-white p-2'><div class='text-xs text-slate-500'>Checkliste</div><div class='font-semibold'>${escapeHtml(checklistStatus)}</div></div>
        <div class='rounded border bg-white p-2'><div class='text-xs text-slate-500'>Restmenge</div><div class='font-semibold'>${escapeHtml(remaining)}</div></div>
        <div class='rounded border bg-white p-2'><div class='text-xs text-slate-500'>Status</div><div class='font-semibold'>${escapeHtml(canCompleteLabel)}</div></div>
      </div>
      ${blockerHtml}
      ${completedNotice}
    </section>`;
  }

  return `<section class='border rounded-lg bg-slate-50 p-4 space-y-4'>
    <div class='flex items-start justify-between gap-3 flex-wrap'>
      <div>
        <h4 class='font-semibold'>Abschluss-Checkliste</h4>
        <p class='text-sm text-slate-600 mt-1'>Pflichtpunkte je Auftrag vor der Fertigmeldung.</p>
      </div>
      ${completeButton}
    </div>
    ${missingTemplatesNotice}
    ${missingItemsNotice}
    ${itemsHtml}
    ${blockerHtml}
    ${completedNotice}
  </section>`;
}

function renderProductionOrderChecklistItem(item, canEdit) {
  const saving = state.ui?.productionChecklistSavingKey === item.id;
  const checkedAt = item.checked_at
    ? new Date(item.checked_at).toLocaleString("de-DE")
    : "";
  const checkedBy = item.checked_by_employee_id
    ? getEmployeeDisplayNameById(item.checked_by_employee_id)
    : "";
  const statusClass = item.checked
    ? "bg-emerald-100 text-emerald-800"
    : "bg-slate-100 text-slate-700";
  const statusLabel = item.checked ? "erledigt" : "offen";
  const detail = item.checked
    ? [checkedAt, checkedBy && `von ${checkedBy}`].filter(Boolean).join(" · ")
    : "";
  return `<label class='flex items-start gap-3 rounded border bg-white p-3 text-sm'>
    <input type='checkbox' class='mt-1 h-4 w-4' onchange="toggleProductionOrderChecklistItem('${item.id}', this.checked)" ${item.checked ? "checked" : ""} ${canEdit && !saving ? "" : "disabled"} />
    <span class='flex-1'>
      <span class='font-semibold block'>${escapeHtml(item.item_label)}</span>
      ${detail ? `<span class='text-xs text-slate-500 block mt-1'>${escapeHtml(detail)}</span>` : ""}
    </span>
    <span class='px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${escapeHtml(statusLabel)}</span>
  </label>`;
}

function renderProductionOrderEmployeesSection(order) {
  const canEdit = canEditProductionOrder(order);
  const activeEmployees = getProductionOrderEmployees(order.id, true);
  const inactiveEmployees = getProductionOrderEmployees(order.id, false)
    .filter((entry) => entry.active === false);
  const assignableEmployees = getAssignableProductionEmployees(order.id);
  const employeeOptions = assignableEmployees
    .map((employee) => {
      const label = `${getProductionEmployeeDisplayName(employee)}${employee.personnel_no ? ` (${employee.personnel_no})` : ""}`;
      return `<option value='${escapeHtml(employee.id)}'>${escapeHtml(label)}</option>`;
    })
    .join("");
  const addControl = canEdit
    ? `<div class='flex gap-2 flex-wrap items-end'>
        <label class='block text-sm min-w-[240px] flex-1'>
          Mitarbeiter hinzufügen
          <select id='productionOrderEmployeeSelect-${order.id}' class='border rounded p-2 w-full mt-1 bg-white' ${assignableEmployees.length ? "" : "disabled"}>
            <option value=''>Bitte auswählen</option>
            ${employeeOptions}
          </select>
        </label>
        <button type='button' class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="addProductionOrderEmployee('${order.id}')" ${assignableEmployees.length ? "" : "disabled"}>Hinzufügen</button>
      </div>`
    : "";
  const activeList = activeEmployees.length
    ? `<div class='grid sm:grid-cols-2 gap-2'>${activeEmployees.map((entry) => renderProductionOrderEmployeePill(entry, canEdit, true)).join("")}</div>`
    : `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Noch kein Mitarbeiter dem Auftrag zugeordnet.</div>`;
  const inactiveList = inactiveEmployees.length
    ? `<details class='text-sm'>
        <summary class='cursor-pointer font-semibold text-slate-600'>Früher zugeordnet</summary>
        <div class='grid sm:grid-cols-2 gap-2 mt-2'>${inactiveEmployees.map((entry) => renderProductionOrderEmployeePill(entry, canEdit, false)).join("")}</div>
      </details>`
    : "";
  const noEmployeesNotice =
    canEdit && !assignableEmployees.length
      ? `<div class='text-xs text-slate-500'>Keine weiteren aktiven Mitarbeiter verfügbar.</div>`
      : "";

  return `<section class='border rounded-lg bg-slate-50 p-4 space-y-4'>
    <div>
      <h4 class='font-semibold'>Mitarbeiter am Auftrag</h4>
      <p class='text-sm text-slate-600 mt-1'>Aktive Mitarbeiter können je Spannung Gutteile zählen.</p>
    </div>
    ${addControl}
    ${noEmployeesNotice}
    ${activeList}
    ${inactiveList}
  </section>`;
}

function renderProductionOrderEmployeePill(entry, canEdit, isActive) {
  const action = canEdit
    ? isActive
      ? `<button type='button' class='px-2 py-1 rounded bg-slate-200 text-slate-800 text-xs' onclick="deactivateProductionOrderEmployee('${entry.id}')">Entfernen</button>`
      : `<button type='button' class='px-2 py-1 rounded bg-emerald-700 text-white text-xs' onclick="reactivateProductionOrderEmployee('${entry.id}')">Reaktivieren</button>`
    : "";
  return `<div class='rounded border bg-white p-3 flex items-start justify-between gap-3'>
    <div>
      <div class='font-semibold text-sm'>${escapeHtml(entry.employee_name || "Mitarbeiter")}</div>
      <div class='text-xs text-slate-500'>${entry.personnel_no ? `PN ${escapeHtml(entry.personnel_no)} · ` : ""}${escapeHtml(getProductionOrderEmployeeRoleLabel(entry.role))}</div>
    </div>
    ${action}
  </div>`;
}

function renderProductionOrderStationCard(station, canEdit) {
  const readonly = canEdit ? "" : "disabled";
  const orderEmployees = getProductionOrderEmployees(station.order_id, true);
  const stationGoodTotal = getProductionStationGoodTotal(station.id);
  const stationScrapClarifyControls = renderProductionStationScrapClarifyControls(
    station,
    canEdit,
  );
  const employeesList = orderEmployees.length
    ? `<div class='space-y-2'>${orderEmployees.map((entry) => renderProductionStationEmployeeCounter(station, entry, canEdit)).join("")}</div>`
    : `<div class='rounded border border-amber-200 bg-amber-50 p-2 text-xs text-amber-800'>Noch kein Mitarbeiter dem Auftrag zugeordnet.</div>`;
  const actualMinutesValue =
    station.actual_time_minutes === null || station.actual_time_minutes === undefined
      ? ""
      : station.actual_time_minutes;
  const actionButtons = canEdit
    ? `<div class='flex gap-2 flex-wrap'>
        <button type='button' class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="saveProductionOrderStation('${station.id}')">Speichern</button>
        <button type='button' class='px-3 py-2 rounded bg-rose-700 text-white text-sm' onclick="deleteProductionOrderStation('${station.id}')">Löschen</button>
      </div>`
    : `<span class='text-sm text-slate-500'>Nur lesbar</span>`;
  return `<article class='border rounded-lg bg-white p-4 space-y-3'>
    <div class='flex items-start justify-between gap-3'>
      <div>
        <h5 class='font-semibold'>Spannung ${escapeHtml(station.station_no)}</h5>
        <p class='text-xs text-slate-500'>${escapeHtml(getProductionStationTimeStatusLabel(station.time_status))}</p>
      </div>
      <span class='px-2 py-1 rounded-full text-xs font-semibold bg-slate-100 text-slate-700'>OP ${escapeHtml(station.op_number || "-")}</span>
    </div>
    <label class='block text-sm'>
      Name
      <input id='productionStationName-${station.id}' class='border rounded p-2 w-full mt-1 bg-white' value="${escapeHtml(station.name || "")}" ${readonly} />
    </label>
    <div class='grid sm:grid-cols-2 gap-3'>
      <label class='block text-sm'>
        OP-Nummer
        <input id='productionStationOp-${station.id}' class='border rounded p-2 w-full mt-1 bg-white' value="${escapeHtml(station.op_number || "")}" ${readonly} />
      </label>
      <label class='block text-sm'>
        Zeitstatus
        <select id='productionStationTimeStatus-${station.id}' class='border rounded p-2 w-full mt-1 bg-white' ${readonly}>
          <option value='ok' ${station.time_status === "ok" ? "selected" : ""}>Zeit ok</option>
          <option value='changed' ${station.time_status === "changed" ? "selected" : ""}>Zeit geändert</option>
        </select>
      </label>
    </div>
    <label class='block text-sm'>
      Ist-Zeit Minuten
      <input id='productionStationActualMinutes-${station.id}' type='number' min='0' step='1' class='border rounded p-2 w-full mt-1 bg-white' value="${escapeHtml(actualMinutesValue)}" ${readonly} />
    </label>
    <div class='grid grid-cols-2 gap-2 text-sm'>
      <div class='rounded border bg-slate-50 p-2'>
        <div class='text-xs text-slate-500'>Ausschuss gesamt</div>
        <div class='font-semibold'>${escapeHtml(station.scrap_total)}</div>
      </div>
      <div class='rounded border bg-slate-50 p-2'>
        <div class='text-xs text-slate-500'>In Abklärung gesamt</div>
        <div class='font-semibold'>${escapeHtml(station.clarify_total)}</div>
      </div>
      <div class='rounded border bg-slate-50 p-2'>
        <div class='text-xs text-slate-500'>Ausschuss Lebenslauf</div>
        <div class='font-semibold'>${escapeHtml(station.scrap_lifetime)}</div>
      </div>
      <div class='rounded border bg-slate-50 p-2'>
        <div class='text-xs text-slate-500'>In Abklärung Lebenslauf</div>
        <div class='font-semibold'>${escapeHtml(station.clarify_lifetime)}</div>
      </div>
    </div>
    ${stationScrapClarifyControls}
    <div class='space-y-2'>
      <div class='flex items-center justify-between gap-2'>
        <div class='text-sm font-semibold'>Mitarbeiter</div>
        <div class='text-xs font-semibold text-slate-600'>Summe Gutteile: ${escapeHtml(stationGoodTotal)}</div>
      </div>
      ${employeesList}
    </div>
    ${actionButtons}
  </article>`;
}

function renderProductionStationScrapClarifyControls(station, canEdit) {
  return `<div class='rounded border border-slate-200 bg-slate-50 p-3 space-y-3'>
    <div>
      <div class='text-sm font-semibold'>Ausschuss / In Abklärung</div>
      <div class='text-xs text-slate-500 mt-1'>Ursache wird beim +1 erfasst.</div>
    </div>
    <div class='grid sm:grid-cols-2 gap-3'>
      ${renderProductionStationAmountControl(station, "scrap", "Ausschuss", station.scrap_total, canEdit)}
      ${renderProductionStationAmountControl(station, "clarify", "In Abklärung", station.clarify_total, canEdit)}
    </div>
  </div>`;
}

function renderProductionStationAmountControl(station, type, label, value, canEdit) {
  const savingKey = `${station.id}:${type}`;
  const saving = state.ui?.productionStationAmountSavingKey === savingKey;
  const disabled = !canEdit || saving ? "disabled" : "";
  return `<div class='rounded bg-white border p-3'>
    <div class='text-xs text-slate-500'>${escapeHtml(label)}</div>
    <div class='flex items-center gap-2 mt-2'>
      <button type='button' aria-label='${escapeHtml(label)} -1' class='px-3 py-2 rounded bg-slate-200 text-slate-900 font-semibold disabled:opacity-50' onclick="adjustProductionStationAmount('${station.id}', '${type}', -1)" ${disabled}>-1</button>
      <div class='min-w-[72px] rounded bg-slate-50 border px-3 py-2 text-center text-lg font-bold'>${escapeHtml(value)}</div>
      <button type='button' aria-label='${escapeHtml(label)} +1' class='px-3 py-2 rounded ${type === "scrap" ? "bg-rose-700" : "bg-amber-600"} text-white font-semibold disabled:opacity-50' onclick="adjustProductionStationAmount('${station.id}', '${type}', 1)" ${disabled}>+1</button>
    </div>
  </div>`;
}

function renderProductionStationEmployeeCounter(station, entry, canEdit) {
  const goodQty = getProductionStationGoodQty(station.id, entry.id);
  const savingKey = `${station.id}:${entry.id}`;
  const saving = state.ui?.productionGoodQtySavingKey === savingKey;
  const disabled = !canEdit || saving ? "disabled" : "";
  return `<div class='rounded border border-slate-200 bg-slate-50 p-3 text-sm'>
    <div class='flex items-center justify-between gap-3 flex-wrap'>
      <div>
        <div class='font-semibold'>${escapeHtml(entry.employee_name || "Mitarbeiter")}</div>
        <div class='text-xs text-slate-500'>${entry.personnel_no ? `PN ${escapeHtml(entry.personnel_no)} · ` : ""}${escapeHtml(getProductionOrderEmployeeRoleLabel(entry.role))}</div>
      </div>
      <div class='flex items-center gap-2'>
        <button type='button' aria-label='Gutteile -1 ${escapeHtml(entry.employee_name || "Mitarbeiter")}' class='px-3 py-2 rounded bg-slate-200 text-slate-900 font-semibold disabled:opacity-50' onclick="adjustProductionStationGoodQty('${station.id}', '${entry.id}', -1)" ${disabled}>-1</button>
        <div class='min-w-[72px] rounded bg-white border px-3 py-2 text-center'>
          <div class='text-xs text-slate-500'>Gutteile</div>
          <div class='text-lg font-bold'>${escapeHtml(goodQty)}</div>
        </div>
        <button type='button' aria-label='Gutteile +1 ${escapeHtml(entry.employee_name || "Mitarbeiter")}' class='px-3 py-2 rounded bg-emerald-700 text-white font-semibold disabled:opacity-50' onclick="adjustProductionStationGoodQty('${station.id}', '${entry.id}', 1)" ${disabled}>+1</button>
      </div>
    </div>
  </div>`;
}

function renderProductionPreviewMetric(label, value) {
  return `<div class='border rounded-lg bg-slate-50 p-3'>
    <div class='text-xs text-slate-500'>${escapeHtml(label)}</div>
    <div class='text-2xl font-bold mt-1'>${escapeHtml(value)}</div>
  </div>`;
}

function renderProductionCounterMachineDetail(machine, orders) {
  const departmentDisplay = getProductionOrderDepartmentDisplay(
    machine.id,
    machine.department_id,
  );
  const rows = orders
    .map(
      (order) => `<tr class='border-b'>
        <td class='p-2 font-semibold'>${escapeHtml(order.ba_number || "-")}</td>
        <td class='p-2'>${escapeHtml(order.article_number || "-")}</td>
        <td class='p-2'>${escapeHtml(getProductionOrderStatusLabel(order.status))}</td>
        <td class='p-2'>${escapeHtml(order.ba_quantity)}</td>
        <td class='p-2'>${escapeHtml(order.target_quantity)}</td>
      </tr>`,
    )
    .join("");
  return `<div class='space-y-3'>
    <div>
      <div class='font-semibold'>${escapeHtml(machine.name || machine.machine_code || "-")}</div>
      <div class='text-xs text-slate-500'>${escapeHtml(machine.machine_code || "-")} · ${escapeHtml(departmentDisplay)}</div>
    </div>
    <div class='overflow-auto border rounded'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2 text-left'>BA</th>
            <th class='p-2 text-left'>Artikel</th>
            <th class='p-2 text-left'>Status</th>
            <th class='p-2 text-left'>BA-Stückzahl</th>
            <th class='p-2 text-left'>Ziel</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <button type='button' class='px-3 py-2 rounded bg-slate-200 text-slate-800 cursor-not-allowed' disabled>Produktionsansicht öffnen</button>
    <p class='text-xs text-slate-500'>Produktionsansicht, Spannungen und echte Zählung folgen später.</p>
  </div>`;
}

function renderProductionCounterNoOrderNotice(machine) {
  return `<div class='space-y-3'>
    <div>
      <div class='font-semibold'>${escapeHtml(machine.name || machine.machine_code || "-")}</div>
      <div class='text-xs text-slate-500'>${escapeHtml(machine.machine_code || "-")}</div>
    </div>
    <div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Für diese Maschine ist kein laufender oder pausierter Auftrag vorhanden. Auftrag zuerst unter Aufträge / BA anlegen.</div>
  </div>`;
}

function renderLegacyProductionCountsTab() {
  const machines = getActiveProductionCountMachines();
  const machineInfo = !machines.length
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Keine aktiven Maschinen für den Stückzahl-Zähler verfügbar.</div>`
    : "";
  const currentEmployeeName =
    currentEmployeeRecord?.display_name ||
    currentEmployeeRecord?.name ||
    currentUser?.name ||
    "-";
  const newForm = `<div class='border rounded-lg p-3 bg-slate-50 space-y-3'>
      <h3 class='font-semibold'>Neuen Zähler starten</h3>
      ${machineInfo}
      <div class='grid md:grid-cols-2 lg:grid-cols-3 gap-3'>
        <label class='text-sm space-y-1'>
          <span class='font-medium'>Maschine</span>
          <select id='productionNewCountMachine' class='border rounded p-2 bg-white w-full' onchange='updateProductionCountDepartmentPreview("productionNewCountMachine", "productionNewCountDepartmentPreview")'>${renderProductionMachineOptions("")}</select>
        </label>
        <label class='text-sm space-y-1'>
          <span class='font-medium'>Abteilung</span>
          <div id='productionNewCountDepartmentPreview' class='border rounded p-2 bg-white text-slate-600 min-h-[42px]'>-</div>
        </label>
        <label class='text-sm space-y-1'>
          <span class='font-medium'>Mitarbeiter</span>
          ${
            currentUser?.role === "admin"
              ? `<select id='productionNewCountEmployee' class='border rounded p-2 bg-white w-full'>${renderProductionCountEmployeeOptions(currentEmployeeRecord?.id || "")}</select>`
              : `<div class='border rounded p-2 bg-white text-slate-600 min-h-[42px]'>${escapeHtml(currentEmployeeName)}</div>`
          }
        </label>
        <input id='productionNewCountOrder' class='border rounded p-2 bg-white' placeholder='Auftrag' />
        <input id='productionNewCountArticle' class='border rounded p-2 bg-white' placeholder='Artikel' />
        <textarea id='productionNewCountNote' class='border rounded p-2 bg-white md:col-span-2 lg:col-span-1' rows='2' placeholder='Notiz'></textarea>
      </div>
      <div class='grid sm:grid-cols-2 gap-3'>
        ${renderProductionCountStepper("productionNewCountGood", "Gutteile", 0)}
        ${renderProductionCountStepper("productionNewCountScrap", "Ausschuss", 0)}
      </div>
      <div class='flex gap-2 flex-wrap'>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='createProductionCount()' ${machines.length ? "" : "disabled"}>Speichern</button>
        <button class='px-3 py-2 rounded bg-slate-200 text-slate-800' onclick='resetNewProductionCountForm()'>Zurücksetzen</button>
      </div>
    </div>`;

  const rows = getVisibleProductionCounts()
    .map((count) => renderProductionCountRow(count))
    .join("");

  return `<div class='space-y-4'>
    ${newForm}
    <div class='border rounded-lg bg-white overflow-auto'>
      <table class='w-full text-sm min-w-[1050px]'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Status</th>
            <th class='p-2 text-left'>Maschine</th>
            <th class='p-2 text-left'>Abteilung</th>
            <th class='p-2 text-left'>Auftrag</th>
            <th class='p-2 text-left'>Artikel</th>
            <th class='p-2 text-left'>Mitarbeiter</th>
            <th class='p-2 text-left'>Gutteile</th>
            <th class='p-2 text-left'>Ausschuss</th>
            <th class='p-2 text-left'>Notiz</th>
            <th class='p-2 text-left'>Aktion</th>
          </tr>
        </thead>
        <tbody>${rows || "<tr><td class='p-3 text-slate-500' colspan='10'>Keine Stückzahl-Zähler geladen.</td></tr>"}</tbody>
      </table>
    </div>
  </div>`;
}

function renderProductionCountStepper(inputId, label, value) {
  const safeValue = Math.max(0, Number(value || 0));
  return `<div class='border rounded-lg bg-white p-3 space-y-2'>
    <div class='text-sm font-medium'>${escapeHtml(label)}</div>
    <div class='flex items-center gap-2'>
      <button class='px-3 py-2 rounded bg-slate-200 text-slate-800' onclick="adjustProductionCountField('${inputId}', -1)">-1</button>
      <input id='${inputId}' type='number' min='0' step='1' class='border rounded p-2 w-24 text-center bg-white' value='${safeValue}' />
      <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick="adjustProductionCountField('${inputId}', 1)">+1</button>
    </div>
  </div>`;
}

function renderProductionCountRow(count) {
  const canEdit = canEditProductionCount(count);
  const isRunning = count.status !== "completed";
  const editable = canEdit && isRunning;
  const machine = getProductionMachineById(count.machine_id);
  const machineDisplay = machine?.name || machine?.machine_code || count.machine_id || "-";
  const departmentDisplay = getProductionCountDepartmentDisplay(count.machine_id);
  const employeeDisplay = getEmployeeDisplayNameById(count.employee_id);
  const statusClass =
    count.status === "completed"
      ? "bg-slate-200 text-slate-700"
      : "bg-emerald-100 text-emerald-800";
  const actionButtons = editable
    ? `<button class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="saveProductionCount('${count.id}')">Speichern</button>
      <button class='px-3 py-2 rounded bg-emerald-700 text-white text-sm ml-2' onclick="completeProductionCount('${count.id}')">Abschließen</button>
      <button class='px-3 py-2 rounded bg-slate-200 text-slate-800 text-sm ml-2' onclick="resetProductionCount('${count.id}')">Zurücksetzen</button>`
    : "-";

  return `<tr class='border-b align-top ${isRunning ? "" : "bg-slate-50 text-slate-500"}'>
    <td class='p-2 whitespace-nowrap'><span class='px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${count.status === "completed" ? "Abgeschlossen" : "Laufend"}</span></td>
    <td class='p-2'>
      ${
        editable
          ? `<select id='productionCountMachine-${count.id}' class='border rounded p-2 w-full bg-white' onchange='updateProductionCountDepartmentPreview("productionCountMachine-${count.id}", "productionCountDepartment-${count.id}")'>${renderProductionMachineOptions(count.machine_id)}</select>`
          : escapeHtml(machineDisplay)
      }
    </td>
    <td class='p-2'><div id='productionCountDepartment-${count.id}'>${escapeHtml(departmentDisplay)}</div></td>
    <td class='p-2'>${editable ? `<input id='productionCountOrder-${count.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(count.order_no)}' />` : escapeHtml(count.order_no || "-")}</td>
    <td class='p-2'>${editable ? `<input id='productionCountArticle-${count.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(count.article_no)}' />` : escapeHtml(count.article_no || "-")}</td>
    <td class='p-2'>
      ${
        editable && currentUser?.role === "admin"
          ? `<select id='productionCountEmployee-${count.id}' class='border rounded p-2 w-full bg-white'>${renderProductionCountEmployeeOptions(count.employee_id)}</select>`
          : escapeHtml(employeeDisplay)
      }
    </td>
    <td class='p-2'>${editable ? renderProductionCountInlineStepper(`productionCountGood-${count.id}`, count.good_qty) : escapeHtml(count.good_qty)}</td>
    <td class='p-2'>${editable ? renderProductionCountInlineStepper(`productionCountScrap-${count.id}`, count.scrap_qty) : escapeHtml(count.scrap_qty)}</td>
    <td class='p-2'>${editable ? `<textarea id='productionCountNote-${count.id}' class='border rounded p-2 w-full bg-white' rows='2'>${escapeHtml(count.note)}</textarea>` : escapeHtml(count.note || "-")}</td>
    <td class='p-2 whitespace-nowrap'>${actionButtons}</td>
  </tr>`;
}

function renderProductionCountInlineStepper(inputId, value) {
  const safeValue = Math.max(0, Number(value || 0));
  return `<div class='flex items-center gap-1'>
    <button class='px-2 py-1 rounded bg-slate-200 text-slate-800' onclick="adjustProductionCountField('${inputId}', -1)">-1</button>
    <input id='${inputId}' type='number' min='0' step='1' class='border rounded p-1 w-20 text-center bg-white' value='${safeValue}' />
    <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="adjustProductionCountField('${inputId}', 1)">+1</button>
  </div>`;
}

function getProductionEventTypeLabel(type) {
  return {
    good: "Gutteil",
    good_correction: "Gutteil-Korrektur",
    scrap: "Ausschuss",
    scrap_correction: "Ausschuss-Korrektur",
    clarify: "In Abklärung",
    clarify_correction: "Abklär-Korrektur",
    station_created: "Spannung angelegt",
    station_deleted: "Spannung gelöscht",
    station_renamed: "Spannung umbenannt",
    order_created: "Auftrag angelegt",
    order_completed: "Auftrag abgeschlossen",
    order_cancelled: "Auftrag abgebrochen",
    checklist: "Checkliste",
    time_changed: "Zeit geändert",
    note: "Notiz",
  }[type] || type || "Ereignis";
}

function getProductionProtocolEmployeeName(entry) {
  if (entry.employee_id) return getEmployeeDisplayNameById(entry.employee_id);
  if (entry.order_employee_id) {
    const orderEmployee = getProductionOrderEmployeeById(entry.order_employee_id);
    if (orderEmployee?.employee_id) return getEmployeeDisplayNameById(orderEmployee.employee_id);
    return orderEmployee?.employee_name || "-";
  }
  return "-";
}

function getProductionProtocolStationName(stationId) {
  const station = getProductionOrderStationById(stationId);
  if (!station) return stationId ? "Spannung" : "-";
  return `Spannung ${station.station_no}${station.op_number ? ` · OP ${station.op_number}` : ""}`;
}

function summarizeProductionHistoryPayload(entry) {
  const payload = entry.payload || {};
  if (!payload || typeof payload !== "object") return "";
  if (entry.history_type === "order_completed") {
    return [
      payload.ba_number && `BA ${payload.ba_number}`,
      `Gutteile ${Number(payload.good_total || 0)}`,
      `Ausschuss ${Number(payload.scrap_total || 0)}`,
      `Abklärung ${Number(payload.clarify_total || 0)}`,
      `Restmenge ${Number(payload.remaining || 0)}`,
    ].filter(Boolean).join(" · ");
  }
  if (entry.history_type === "checklist") {
    return `${payload.item_label || payload.item_key || "Checkliste"}: ${payload.checked ? "erledigt" : "offen"}`;
  }
  return Object.entries(payload)
    .slice(0, 5)
    .map(([key, value]) => `${key}: ${typeof value === "object" ? JSON.stringify(value) : value}`)
    .join(" · ");
}

function getProductionProtocolEntries(orderId) {
  if (!orderId) return [];
  const stationEvents = (state.productionStationEvents || [])
    .filter((event) => event.order_id === orderId)
    .map((event) => ({ ...event, source: "station_event" }));
  const orderHistory = (state.productionOrderHistory || [])
    .filter((entry) => entry.order_id === orderId)
    .map((entry) => ({ ...entry, source: "order_history", event_type: entry.history_type }));
  const direction = state.ui?.productionProtocolSort === "asc" ? 1 : -1;
  return [...stationEvents, ...orderHistory].sort((a, b) => {
    const aTime = new Date(a.created_at || 0).getTime();
    const bTime = new Date(b.created_at || 0).getTime();
    return (aTime - bTime) * direction;
  });
}

function renderProductionProtocolTab() {
  const orders = getVisibleProductionOrders();
  const selectedOrderId =
    state.ui?.productionProtocolSelectedOrderId &&
    orders.some((order) => order.id === state.ui.productionProtocolSelectedOrderId)
      ? state.ui.productionProtocolSelectedOrderId
      : orders[0]?.id || "";
  const order = getProductionOrderById(selectedOrderId);
  const orderOptions = orders
    .map((entry) => {
      const machine = getProductionMachineById(entry.machine_id);
      const label = `${entry.ba_number || "-"} · ${machine?.name || machine?.machine_code || "Maschine"} · ${getProductionOrderStatusLabel(entry.status)}`;
      return `<option value='${escapeHtml(entry.id)}' ${entry.id === selectedOrderId ? "selected" : ""}>${escapeHtml(label)}</option>`;
    })
    .join("");
  const entries = getProductionProtocolEntries(selectedOrderId);
  const rows = entries.length
    ? `<div class='space-y-3'>${entries.map(renderProductionProtocolEntry).join("")}</div>`
    : `<div class='rounded border border-slate-200 bg-slate-50 p-4 text-sm text-slate-600'>Noch keine Protokolleinträge vorhanden.</div>`;
  const sortLabel = state.ui?.productionProtocolSort === "asc" ? "Älteste zuerst" : "Neueste zuerst";

  return `<div class='space-y-4'>
    <div class='border rounded-lg bg-slate-50 p-4'>
      <h3 class='text-xl font-bold'>Produktionsprotokoll</h3>
      <p class='text-sm text-slate-600 mt-1'>Ereignisse und Auftragshistorie je BA.</p>
    </div>
    <div class='border rounded-lg bg-white p-4 space-y-3'>
      <div class='grid md:grid-cols-[minmax(0,1fr),auto] gap-3 items-end'>
        <label class='block text-sm'>
          Auftrag / BA
          <select class='border rounded p-2 w-full mt-1 bg-white' onchange="selectProductionProtocolOrder(this.value)">
            ${orderOptions || "<option value=''>Keine Aufträge sichtbar</option>"}
          </select>
        </label>
        <button type='button' class='px-3 py-2 rounded bg-slate-200 text-slate-800 text-sm' onclick='toggleProductionProtocolSort()'>${escapeHtml(sortLabel)}</button>
      </div>
      ${order ? `<div class='rounded border bg-slate-50 p-3 text-sm text-slate-700'>BA ${escapeHtml(order.ba_number || "-")} · ${escapeHtml(getProductionOrderStatusLabel(order.status))}</div>` : ""}
      ${rows}
    </div>
  </div>`;
}

function renderProductionProtocolEntry(entry) {
  const isHistory = entry.source === "order_history";
  const label = getProductionEventTypeLabel(isHistory ? entry.history_type : entry.event_type);
  const timestamp = entry.created_at ? new Date(entry.created_at).toLocaleString("de-DE") : "-";
  const station = entry.station_id ? getProductionProtocolStationName(entry.station_id) : "";
  const employee = getProductionProtocolEmployeeName(entry);
  const cause = entry.qa_cause_id ? getProductionQaCauseById(entry.qa_cause_id) : null;
  const causeText = cause ? `${cause.group_label}: ${cause.reason_label}` : "";
  const payload = isHistory ? summarizeProductionHistoryPayload(entry) : "";
  const details = [
    entry.qty ? `Menge ${entry.qty}` : "",
    station,
    employee !== "-" ? `Mitarbeiter ${employee}` : "",
    causeText,
    payload,
    entry.note,
  ].filter(Boolean);
  return `<article class='rounded border bg-white p-3 text-sm'>
    <div class='flex items-start justify-between gap-3 flex-wrap'>
      <div>
        <div class='font-semibold'>${escapeHtml(label)}</div>
        <div class='text-xs text-slate-500 mt-1'>${escapeHtml(timestamp)} · ${isHistory ? "Auftragshistorie" : "Stationsereignis"}</div>
      </div>
      ${entry.qty ? `<span class='px-2 py-1 rounded-full bg-slate-100 text-slate-700 text-xs font-semibold'>${escapeHtml(entry.qty)}</span>` : ""}
    </div>
    ${details.length ? `<div class='mt-2 text-slate-700 space-y-1'>${details.map((detail) => `<div>${escapeHtml(detail)}</div>`).join("")}</div>` : ""}
  </article>`;
}

function renderProductionSettingsTab() {
  return `<div class='border rounded-lg p-3 bg-slate-50'>
    <h3 class='font-semibold mb-2'>Einstellungen</h3>
    <p class='text-sm text-slate-600'>Produktionseinstellungen folgen später.</p>
  </div>`;
}

function readProductionDepartmentForm(id) {
  return {
    name:
      document.getElementById(`production-department-name-${id}`)?.value?.trim() ||
      "",
    code:
      document.getElementById(`production-department-code-${id}`)?.value?.trim() ||
      "",
    leaderEmployeeId:
      document.getElementById(`production-department-leader-${id}`)?.value || "",
  };
}

function validateProductionDepartmentInput(values) {
  if (!values.name) return "Bitte Name der Abteilung ausfüllen.";
  if (!values.code) return "Bitte Code der Abteilung ausfüllen.";
  return "";
}

function readProductionMachineForm(id) {
  return {
    name:
      document.getElementById(`production-machine-name-${id}`)?.value?.trim() ||
      "",
    machineCode:
      document.getElementById(`production-machine-code-${id}`)?.value?.trim() ||
      "",
    departmentId:
      document.getElementById(`production-machine-department-${id}`)?.value || "",
  };
}

function validateProductionMachineInput(values) {
  if (!values.name) return "Bitte Name der Maschine ausfüllen.";
  if (!values.machineCode) return "Bitte Maschinencode ausfüllen.";
  return "";
}

function readNonNegativeIntegerInput(id) {
  const value = Number(document.getElementById(id)?.value || 0);
  return Number.isFinite(value) ? Math.max(0, Math.trunc(value)) : 0;
}

function readProductionOrderForm() {
  const machineId = document.getElementById("productionNewOrderMachine")?.value || "";
  const machine = getProductionMachineById(machineId);
  return {
    machineId,
    departmentId: machine?.department_id || "",
    baNumber:
      document.getElementById("productionNewOrderBaNumber")?.value?.trim() || "",
    articleNumber:
      document.getElementById("productionNewOrderArticleNumber")?.value?.trim() || "",
    baQuantity: readNonNegativeIntegerInput("productionNewOrderBaQuantity"),
    targetQuantity: readNonNegativeIntegerInput("productionNewOrderTargetQuantity"),
    palletCount: readNonNegativeIntegerInput("productionNewOrderPalletCount"),
    piecesPerPallet: readNonNegativeIntegerInput("productionNewOrderPiecesPerPallet"),
    useChainLogic:
      document.getElementById("productionNewOrderUseChainLogic")?.checked === true,
  };
}

function validateProductionOrderInput(values) {
  if (!values.machineId) return "Bitte Maschine auswählen.";
  if (!getActiveProductionOrderMachines().some((machine) => machine.id === values.machineId)) {
    return "Diese Maschine ist für Aufträge / BA nicht verfügbar.";
  }
  if (!values.baNumber) return "Bitte BA-Nummer ausfüllen.";
  if (
    values.baQuantity < 0 ||
    values.targetQuantity < 0 ||
    values.palletCount < 0 ||
    values.piecesPerPallet < 0
  ) {
    return "Zahlenfelder dürfen nicht negativ sein.";
  }
  return "";
}

function readProductionOrderStationValues(stationId) {
  const actualMinutesRaw =
    document.getElementById(`productionStationActualMinutes-${stationId}`)?.value?.trim() || "";
  return {
    name:
      document.getElementById(`productionStationName-${stationId}`)?.value?.trim() || "",
    opNumber:
      document.getElementById(`productionStationOp-${stationId}`)?.value?.trim() || "",
    timeStatus:
      document.getElementById(`productionStationTimeStatus-${stationId}`)?.value === "changed"
        ? "changed"
        : "ok",
    actualMinutesRaw,
    actualMinutes:
      actualMinutesRaw === "" ? null : Math.trunc(Number(actualMinutesRaw)),
  };
}

function validateProductionOrderStationInput(values) {
  if (!values.name) return "Bitte Namen der Spannung ausfüllen.";
  if (values.actualMinutesRaw !== "" && !/^\d+$/.test(values.actualMinutesRaw)) {
    return "Ist-Zeit muss eine ganze Zahl ab 0 sein.";
  }
  return "";
}

function getAllowedProductionCountMachine(machineId) {
  if (!machineId) return null;
  return getActiveProductionCountMachines().find(
    (machine) => machine.id === machineId,
  ) || null;
}

function readProductionCountValues(prefix, id = "") {
  const suffix = id ? `-${id}` : "";
  const machineId =
    document.getElementById(`${prefix}Machine${suffix}`)?.value || "";
  const machine = getAllowedProductionCountMachine(machineId);
  const existingCount = id ? getProductionCountById(id) : null;
  const employeeId =
    currentUser?.role === "admin"
      ? document.getElementById(`${prefix}Employee${suffix}`)?.value || currentEmployeeRecord?.id || ""
      : existingCount?.employee_id || currentEmployeeRecord?.id || "";
  return {
    machineId,
    departmentId: machine?.department_id || "",
    employeeId,
    orderNo:
      document.getElementById(`${prefix}Order${suffix}`)?.value?.trim() || "",
    articleNo:
      document.getElementById(`${prefix}Article${suffix}`)?.value?.trim() || "",
    goodQty: Math.max(
      0,
      Number(document.getElementById(`${prefix}Good${suffix}`)?.value || 0),
    ),
    scrapQty: Math.max(
      0,
      Number(document.getElementById(`${prefix}Scrap${suffix}`)?.value || 0),
    ),
    note:
      document.getElementById(`${prefix}Note${suffix}`)?.value?.trim() || "",
  };
}

function validateProductionCountInput(values) {
  if (!values.machineId) return "Bitte Maschine auswählen.";
  if (!getAllowedProductionCountMachine(values.machineId)) {
    return "Diese Maschine ist für den Stückzahl-Zähler nicht verfügbar.";
  }
  if (!values.employeeId) return "Bitte Mitarbeiter auswählen.";
  if (values.goodQty < 0 || values.scrapQty < 0) {
    return "Stückzahlen dürfen nicht negativ sein.";
  }
  return "";
}

async function findDepartmentByCode(code, exceptId = null) {
  const { data, error } = await supabaseClient
    .from("departments")
    .select("id")
    .eq("code", code)
    .limit(1);

  if (error) return { error, exists: false };

  const existing = (data || []).find((row) => row.id !== exceptId);
  return { error: null, exists: !!existing };
}

async function findProductionMachineByCode(machineCode, exceptId = null) {
  const { data, error } = await supabaseClient
    .from("production_machines")
    .select("id")
    .eq("machine_code", machineCode)
    .limit(1);

  if (error) return { error, exists: false };

  const existing = (data || []).find((row) => row.id !== exceptId);
  return { error: null, exists: !!existing };
}

async function refreshDepartmentsFromSupabase() {
  const departments = await loadDepartmentsFromSupabase();
  applyDepartmentsToState(departments);
}

async function refreshProductionMachinesFromSupabase() {
  const machines = await loadProductionMachinesFromSupabase();
  applyProductionMachinesToState(machines);
}

async function refreshProductionCountsFromSupabase() {
  const counts = await loadProductionCountsFromSupabase();
  applyProductionCountsToState(counts);
}

async function refreshProductionOrdersFromSupabase() {
  const orders = await loadProductionOrdersFromSupabase();
  applyProductionOrdersToState(orders);
}

async function refreshProductionOrderEmployeesFromSupabase() {
  const employees = await loadProductionOrderEmployeesFromSupabase();
  applyProductionOrderEmployeesToState(employees);
}

async function refreshProductionStationCountsFromSupabase() {
  const counts = await loadProductionStationCountsFromSupabase();
  applyProductionStationCountsToState(counts);
}

async function refreshProductionStationEventsFromSupabase() {
  const events = await loadProductionStationEventsFromSupabase();
  if (!Array.isArray(events)) return false;
  applyProductionStationEventsToState(events);
  return true;
}

async function refreshProductionOrderHistoryFromSupabase() {
  const history = await loadProductionOrderHistoryFromSupabase();
  if (!Array.isArray(history)) return false;
  applyProductionOrderHistoryToState(history);
  return true;
}

async function refreshProductionOrderChecklistFromSupabase() {
  const checklist = await loadProductionOrderChecklistFromSupabase();
  if (!Array.isArray(checklist)) return false;
  applyProductionOrderChecklistToState(checklist);
  return true;
}

function isProductionDuplicateError(error) {
  const message = String(error?.message || "").toLowerCase();
  return error?.code === "23505" || message.includes("duplicate") || message.includes("unique");
}

async function ensureProductionOrderChecklistForOrder(orderId) {
  const order = getProductionOrderById(orderId);
  if (!order || !order.id) return false;
  const templates = state.productionChecklistTemplates || [];
  if (!templates.length) return false;
  state.ui = state.ui || {};
  if (state.ui.productionChecklistPreparingOrderId === order.id) return false;
  const existingKeys = new Set(
    getProductionOrderChecklistItems(order.id).map((item) => item.item_key),
  );
  const missingTemplates = templates.filter((template) => !existingKeys.has(template.item_key));
  if (!missingTemplates.length) return true;
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    return false;
  }

  state.ui.productionChecklistPreparingOrderId = order.id;
  const now = new Date().toISOString();
  const payload = missingTemplates.map((template) => ({
    order_id: order.id,
    item_key: template.item_key,
    item_label: template.item_label,
    checked: false,
    checked_at: null,
    checked_by_employee_id: null,
    sort_order: template.sort_order,
    updated_at: now,
  }));
  const { error } = await supabaseClient.from("production_order_checklist").insert(payload);
  if (error && !isProductionDuplicateError(error)) {
    logProductionOrderChecklistSupabaseError("insert", error, order.id);
    setProductionStatus(
      formatProductionSupabaseError(error, "Abschluss-Checkliste konnte nicht vorbereitet werden"),
      true,
    );
    await refreshProductionOrderChecklistFromSupabase();
    state.ui.productionChecklistPreparingOrderId = "";
    return false;
  }
  if (error && isProductionDuplicateError(error)) {
    logProductionOrderChecklistSupabaseError("insert", error, order.id);
  }
  const refreshed = await refreshProductionOrderChecklistFromSupabase();
  state.ui.productionChecklistPreparingOrderId = "";
  return refreshed;
}

async function writeProductionOrderHistory(order, historyType, payload = {}, qty = 0, note = null) {
  return supabaseClient.from("production_order_history").insert([
    {
      order_id: order.id,
      station_id: null,
      employee_id: currentEmployeeRecord?.id || null,
      history_type: historyType,
      qty,
      payload,
      note,
    },
  ]);
}

function selectProductionProtocolOrder(orderId) {
  state.ui = state.ui || {};
  state.ui.productionProtocolSelectedOrderId = orderId || "";
  render();
}

function toggleProductionProtocolSort() {
  state.ui = state.ui || {};
  state.ui.productionProtocolSort = state.ui.productionProtocolSort === "asc" ? "desc" : "asc";
  render();
}

function openProductionProtocolForOrder(orderId) {
  state.ui = state.ui || {};
  state.productionSubTab = "protocol";
  state.ui.productionProtocolSelectedOrderId = orderId || "";
  persist();
  render();
}

async function writeProductionStationGoodEvent(station, orderEmployee, delta) {
  return supabaseClient.from("production_station_events").insert([
    {
      order_id: station.order_id,
      station_id: station.id,
      order_employee_id: orderEmployee.id,
      employee_id: orderEmployee.employee_id || null,
      event_type: delta > 0 ? "good" : "good_correction",
      qty: delta,
      qa_cause_id: null,
      note: null,
    },
  ]);
}

async function writeProductionStationAmountEvent(station, type, delta, qaCauseId = null, note = "") {
  const eventType =
    type === "scrap"
      ? delta > 0 ? "scrap" : "scrap_correction"
      : delta > 0 ? "clarify" : "clarify_correction";
  return supabaseClient.from("production_station_events").insert([
    {
      order_id: station.order_id,
      station_id: station.id,
      order_employee_id: null,
      employee_id: currentEmployeeRecord?.id || null,
      event_type: eventType,
      qty: delta,
      qa_cause_id: qaCauseId || null,
      note: note || null,
    },
  ]);
}

function getProductionQaCauseById(causeId) {
  if (!causeId) return null;
  return (state.productionQaCauses || []).find((cause) => cause.id === causeId) || null;
}

function getGroupedProductionQaCauses() {
  const groups = new Map();
  (state.productionQaCauses || []).forEach((cause) => {
    const groupLabel = cause.group_label || "Ohne Gruppe";
    if (!groups.has(groupLabel)) groups.set(groupLabel, []);
    groups.get(groupLabel).push(cause);
  });
  return Array.from(groups.entries());
}

function openProductionQaCauseModal(stationId, type) {
  const station = getProductionOrderStationById(stationId);
  const order = getProductionOrderById(station?.order_id);
  if (!station || !["scrap", "clarify"].includes(type) || !canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diese Menge nicht ändern.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  const causes = state.productionQaCauses || [];
  if (!causes.length) {
    setProductionStatus("Keine aktiven 6M-Ursachen geladen. +1 kann nicht ohne Ursache gebucht werden.", true);
    render();
    return;
  }
  state.ui = state.ui || {};
  state.ui.productionQaCauseModal = {
    stationId,
    type,
    selectedCauseId: causes[0]?.id || "",
    note: "",
  };
  setProductionStatus("");
  render();
}

function closeProductionQaCauseModal(message = "") {
  state.ui = state.ui || {};
  state.ui.productionQaCauseModal = null;
  if (message) setProductionStatus(message);
  render();
}

function setProductionQaCauseSelection(causeId) {
  state.ui = state.ui || {};
  if (!state.ui.productionQaCauseModal) return;
  state.ui.productionQaCauseModal.selectedCauseId = causeId || "";
  render();
}

function renderProductionQaCauseModal() {
  const modal = state.ui?.productionQaCauseModal;
  if (!modal) return "";
  const station = getProductionOrderStationById(modal.stationId);
  if (!station) return "";
  const label = modal.type === "scrap" ? "Ausschuss" : "In Abklärung";
  const groupedCauses = getGroupedProductionQaCauses();
  const causeButtons = groupedCauses
    .map(([groupLabel, causes]) => {
      const buttons = causes
        .map((cause) => {
          const active = modal.selectedCauseId === cause.id;
          return `<button type='button' class='px-3 py-2 rounded border text-sm text-left ${active ? "bg-slate-900 text-white" : "bg-white text-slate-800"}' onclick="setProductionQaCauseSelection('${cause.id}')">${escapeHtml(cause.reason_label)}</button>`;
        })
        .join("");
      return `<div class='space-y-2'>
        <div class='text-sm font-semibold'>${escapeHtml(groupLabel)}</div>
        <div class='grid sm:grid-cols-2 gap-2'>${buttons}</div>
      </div>`;
    })
    .join("");
  const error = state.ui?.productionQaCauseModalError
    ? `<div class='rounded border border-rose-200 bg-rose-50 p-2 text-sm text-rose-700'>${escapeHtml(state.ui.productionQaCauseModalError)}</div>`
    : "";
  return `<div class='fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4'>
    <div class='bg-white rounded-lg shadow-xl max-w-3xl w-full max-h-[90vh] overflow-auto p-4 space-y-4'>
      <div class='flex items-start justify-between gap-3'>
        <div>
          <h3 class='text-lg font-semibold'>6M-Ursache für ${escapeHtml(label)}</h3>
          <p class='text-sm text-slate-500 mt-1'>Spannung ${escapeHtml(station.station_no)} · Ursache wählen und optional Notiz ergänzen.</p>
        </div>
        <button type='button' class='px-3 py-2 rounded bg-slate-200 text-slate-800' onclick='closeProductionQaCauseModal("Ursachenerfassung abgebrochen.")'>Abbrechen</button>
      </div>
      ${error}
      <div class='space-y-4'>${causeButtons}</div>
      <label class='block text-sm'>
        Notiz
        <textarea id='productionQaCauseNote' class='border rounded p-2 w-full mt-1 bg-white' rows='3' placeholder='Optional'></textarea>
      </label>
      <div class='flex justify-end gap-2'>
        <button type='button' class='px-3 py-2 rounded bg-slate-200 text-slate-800' onclick='closeProductionQaCauseModal("Ursachenerfassung abgebrochen.")'>Abbrechen</button>
        <button type='button' class='px-3 py-2 rounded bg-slate-900 text-white' onclick='confirmProductionQaCauseModal()'>Ursache speichern</button>
      </div>
    </div>
  </div>`;
}

async function confirmProductionQaCauseModal() {
  const modal = state.ui?.productionQaCauseModal;
  if (!modal) return;
  const cause = getProductionQaCauseById(modal.selectedCauseId);
  if (!cause) {
    state.ui.productionQaCauseModalError = "Bitte eine 6M-Ursache auswählen.";
    render();
    return;
  }
  const note = document.getElementById("productionQaCauseNote")?.value?.trim() || "";
  state.ui.productionQaCauseModal = null;
  state.ui.productionQaCauseModalError = "";
  await adjustProductionStationAmount(modal.stationId, modal.type, 1, {
    qaCauseId: cause.id,
    note,
  });
}

async function adjustProductionStationAmount(stationId, type, delta, options = {}) {
  const station = getProductionOrderStationById(stationId);
  const order = getProductionOrderById(station?.order_id);
  if (!station || !["scrap", "clarify"].includes(type) || !canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diese Menge nicht ändern.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  delta = delta > 0 ? 1 : -1;
  if (delta > 0 && !options.qaCauseId) {
    openProductionQaCauseModal(stationId, type);
    return;
  }
  const totalField = type === "scrap" ? "scrap_total" : "clarify_total";
  const lifetimeField = type === "scrap" ? "scrap_lifetime" : "clarify_lifetime";
  const currentTotal = Math.max(0, Number(station[totalField] || 0));
  const currentLifetime = Math.max(0, Number(station[lifetimeField] || 0));
  if (delta < 0 && currentTotal <= 0) {
    setProductionStatus(
      type === "scrap"
        ? "Ausschuss kann nicht unter 0 fallen."
        : "Abklärmenge kann nicht unter 0 fallen.",
      true,
    );
    render();
    return;
  }

  state.ui = state.ui || {};
  const savingKey = `${station.id}:${type}`;
  if (state.ui.productionStationAmountSavingKey === savingKey) return;
  state.ui.productionStationAmountSavingKey = savingKey;
  render();

  const payload = {
    [totalField]: Math.max(0, currentTotal + delta),
    updated_at: new Date().toISOString(),
  };
  if (delta > 0) payload[lifetimeField] = currentLifetime + 1;

  const { error } = await supabaseClient
    .from("production_order_stations")
    .update(payload)
    .eq("id", station.id);

  if (error) {
    logProductionStationSupabaseError("update", error, {
      order_id: station.order_id,
      station_id: station.id,
      station_no: station.station_no,
    });
    state.ui.productionStationAmountSavingKey = "";
    setProductionStatus(
      formatProductionSupabaseError(error, "Menge konnte nicht gespeichert werden"),
      true,
    );
    render();
    return;
  }

  const { error: eventError } = await writeProductionStationAmountEvent(
    station,
    type,
    delta,
    options.qaCauseId || null,
    options.note || "",
  );
  await refreshProductionOrderStationsFromSupabase();
  await refreshProductionStationEventsFromSupabase();
  state.ui.productionStationAmountSavingKey = "";
  if (eventError) {
    console.error("Fehler beim Schreiben des Mengen-Protokolls:", eventError);
    setProductionStatus("Menge gespeichert, Ursache/Protokoll konnte nicht geschrieben werden.", true);
    render();
    return;
  }

  setProductionStatus(
    type === "scrap"
      ? delta > 0 ? "Ausschuss wurde gezählt." : "Ausschuss-Korrektur wurde gespeichert."
      : delta > 0 ? "Abklärmenge wurde gezählt." : "Abklär-Korrektur wurde gespeichert.",
  );
  render();
}

async function adjustProductionStationGoodQty(stationId, orderEmployeeId, delta, retried = false) {
  const station = getProductionOrderStationById(stationId);
  const order = getProductionOrderById(station?.order_id);
  const orderEmployee = getProductionOrderEmployeeById(orderEmployeeId);
  if (!station || !orderEmployee || orderEmployee.active === false || !canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diese Gutteile nicht zählen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  delta = delta > 0 ? 1 : -1;
  const savingKey = `${station.id}:${orderEmployee.id}`;
  if (state.ui?.productionGoodQtySavingKey === savingKey) return;

  const currentCount = getProductionStationCount(station.id, orderEmployee.id);
  const currentQty = currentCount?.good_qty || 0;
  if (delta < 0 && currentQty <= 0) {
    setProductionStatus("Gutmenge kann nicht unter 0 fallen.", true);
    render();
    return;
  }

  state.ui = state.ui || {};
  state.ui.productionGoodQtySavingKey = savingKey;
  render();

  const nextQty = Math.max(0, currentQty + delta);
  let countError = null;
  if (currentCount) {
    const { error } = await supabaseClient
      .from("production_station_counts")
      .update({ good_qty: nextQty, updated_at: new Date().toISOString() })
      .eq("id", currentCount.id);
    countError = error;
  } else {
    const { error } = await supabaseClient.from("production_station_counts").insert([
      {
        order_id: station.order_id,
        station_id: station.id,
        order_employee_id: orderEmployee.id,
        good_qty: nextQty,
        updated_at: new Date().toISOString(),
      },
    ]);
    countError = error;
  }

  if (countError) {
    console.error("Fehler beim Speichern der Gutteile:", countError);
    state.ui.productionGoodQtySavingKey = "";
    if (!retried && isProductionDuplicateError(countError)) {
      await refreshProductionStationCountsFromSupabase();
      await adjustProductionStationGoodQty(station.id, orderEmployee.id, delta, true);
      return;
    }
    setProductionStatus(
      formatProductionSupabaseError(
        countError,
        "Gutteilmenge konnte nicht gespeichert werden",
        "Gutteilmenge konnte nicht gespeichert werden: Zähler existiert bereits.",
      ),
      true,
    );
    render();
    return;
  }

  const { error: eventError } = await writeProductionStationGoodEvent(
    station,
    orderEmployee,
    delta,
  );
  await refreshProductionStationCountsFromSupabase();
  await refreshProductionStationEventsFromSupabase();
  state.ui.productionGoodQtySavingKey = "";
  if (eventError) {
    console.error("Fehler beim Schreiben des Gutteil-Protokolls:", eventError);
    setProductionStatus("Gutteil gespeichert, Protokolleintrag konnte nicht geschrieben werden.", true);
    render();
    return;
  }

  setProductionStatus(delta > 0 ? "Gutteil wurde gezählt." : "Gutteil-Korrektur wurde gespeichert.");
  render();
}

async function setProductionOrderEmployeeActive(entryId, active) {
  const entry = getProductionOrderEmployeeById(entryId);
  const order = getProductionOrderById(entry?.order_id);
  if (!entry || !canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diese Mitarbeiterzuordnung nicht ändern.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_order_employees")
    .update({ active, updated_at: new Date().toISOString() })
    .eq("id", entry.id);

  if (error) {
    console.error("Fehler beim Ändern der Auftragsmitarbeiter:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Mitarbeiterzuordnung konnte nicht geändert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionOrderEmployeesFromSupabase();
  setProductionStatus(active ? "Mitarbeiter wurde reaktiviert." : "Mitarbeiter wurde vom Auftrag entfernt.");
  render();
}

function deactivateProductionOrderEmployee(entryId) {
  setProductionOrderEmployeeActive(entryId, false);
}

function reactivateProductionOrderEmployee(entryId) {
  setProductionOrderEmployeeActive(entryId, true);
}

async function addProductionOrderEmployee(orderId) {
  const order = getProductionOrderById(orderId);
  if (!canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diesem Auftrag keine Mitarbeiter zuordnen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const employeeId = document.getElementById(`productionOrderEmployeeSelect-${order.id}`)?.value || "";
  const employee = getAssignableProductionEmployees("").find((entry) => entry.id === employeeId);
  if (!employee) {
    setProductionStatus("Bitte aktiven Mitarbeiter auswählen.", true);
    render();
    return;
  }

  const existing = getProductionOrderEmployees(order.id, false).find(
    (entry) => entry.employee_id === employee.id,
  );
  if (existing && existing.active !== false) {
    setProductionStatus("Mitarbeiter ist bereits zugeordnet.", true);
    render();
    return;
  }
  if (existing && existing.active === false) {
    await setProductionOrderEmployeeActive(existing.id, true);
    return;
  }

  const currentEntries = getProductionOrderEmployees(order.id, false);
  const nextSortOrder = currentEntries.length
    ? Math.max(...currentEntries.map((entry) => Number(entry.sort_order || 0))) + 1
    : 1;
  const { error } = await supabaseClient.from("production_order_employees").insert([
    {
      order_id: order.id,
      employee_id: employee.id,
      employee_name: getProductionEmployeeDisplayName(employee),
      personnel_no: employee.personnel_no || null,
      role: "worker",
      sort_order: nextSortOrder,
      active: true,
      updated_at: new Date().toISOString(),
    },
  ]);

  if (error) {
    console.error("Fehler beim Hinzufügen des Auftragsmitarbeiters:", error);
    await refreshProductionOrderEmployeesFromSupabase();
    const duplicate = getProductionOrderEmployees(order.id, false).find(
      (entry) => entry.employee_id === employee.id,
    );
    if (duplicate?.active === false) {
      await setProductionOrderEmployeeActive(duplicate.id, true);
      return;
    }
    if (duplicate && duplicate.active !== false) {
      setProductionStatus("Mitarbeiter ist bereits zugeordnet.", true);
      render();
      return;
    }
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Mitarbeiter konnte nicht zugeordnet werden",
        "Mitarbeiter ist bereits zugeordnet.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionOrderEmployeesFromSupabase();
  setProductionStatus("Mitarbeiter wurde dem Auftrag zugeordnet.");
  render();
}

async function createProductionOrderStation(orderId) {
  const order = getProductionOrderById(orderId);
  if (!canEditProductionOrder(order)) {
    setProductionStatus("Du darfst für diesen Auftrag keine Spannung anlegen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const stations = getProductionOrderStations(order.id);
  const nextStationNo = stations.length
    ? Math.max(...stations.map((station) => Number(station.station_no || 0))) + 1
    : 1;
  const { error } = await supabaseClient.from("production_order_stations").insert([
    {
      order_id: order.id,
      station_no: nextStationNo,
      name: `Spannung ${nextStationNo}`,
      lock_name: false,
      op_number: null,
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      updated_at: new Date().toISOString(),
    },
  ]);

  if (error) {
    logProductionStationSupabaseError("insert", error, {
      order_id: order.id,
      station_no: nextStationNo,
    });
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Spannung konnte nicht angelegt werden",
        "Spannung konnte nicht angelegt werden: Nummer ist bereits vorhanden.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionOrderStationsFromSupabase();
  setProductionStatus("Spannung wurde angelegt.");
  render();
}

async function saveProductionOrderStation(stationId) {
  const station = getProductionOrderStationById(stationId);
  const order = getProductionOrderById(station?.order_id);
  if (!station || !canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diese Spannung nicht bearbeiten.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = readProductionOrderStationValues(station.id);
  const validationMessage = validateProductionOrderStationInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_order_stations")
    .update({
      name: values.name,
      lock_name: values.name !== `Spannung ${station.station_no}` || station.lock_name === true,
      op_number: values.opNumber || null,
      time_status: values.timeStatus,
      actual_time_minutes: values.actualMinutes,
      updated_at: new Date().toISOString(),
    })
    .eq("id", station.id);

  if (error) {
    logProductionStationSupabaseError("update", error, {
      order_id: station.order_id,
      station_id: station.id,
      station_no: station.station_no,
    });
    setProductionStatus(
      formatProductionSupabaseError(error, "Spannung konnte nicht gespeichert werden"),
      true,
    );
    render();
    return;
  }

  await refreshProductionOrderStationsFromSupabase();
  setProductionStatus("Spannung wurde gespeichert.");
  render();
}

function productionOrderStationHasTotals(station) {
  return (
    Number(station?.scrap_total || 0) > 0 ||
    Number(station?.clarify_total || 0) > 0 ||
    Number(station?.scrap_lifetime || 0) > 0 ||
    Number(station?.clarify_lifetime || 0) > 0
  );
}

async function productionOrderStationHasRelatedEntries(stationId) {
  const checks = [
    { table: "production_station_counts", label: "Zählungen" },
    { table: "production_station_events", label: "Protokolle" },
  ];
  for (const check of checks) {
    const { data, error } = await supabaseClient
      .from(check.table)
      .select("id")
      .eq("station_id", stationId)
      .limit(1);
    if (error) {
      return {
        error: formatProductionSupabaseError(
          error,
          `${check.label} zur Spannung konnten nicht geprüft werden`,
        ),
        exists: false,
      };
    }
    if (Array.isArray(data) && data.length) {
      return { error: "", exists: true };
    }
  }
  return { error: "", exists: false };
}

async function deleteProductionOrderStation(stationId) {
  const station = getProductionOrderStationById(stationId);
  const order = getProductionOrderById(station?.order_id);
  if (!station || !canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diese Spannung nicht löschen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  if (productionOrderStationHasTotals(station)) {
    setProductionStatus("Spannung kann nicht gelöscht werden, weil bereits Ausschuss oder Abklärung erfasst ist.", true);
    render();
    return;
  }

  const relatedCheck = await productionOrderStationHasRelatedEntries(station.id);
  if (relatedCheck.error) {
    setProductionStatus(relatedCheck.error, true);
    render();
    return;
  }
  if (relatedCheck.exists) {
    setProductionStatus("Spannung kann nicht gelöscht werden, weil bereits Zählungen oder Protokolle verknüpft sind.", true);
    render();
    return;
  }
  if (!confirm("Spannung wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.")) {
    setProductionStatus("Löschen abgebrochen.");
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_order_stations")
    .delete()
    .eq("id", station.id);

  if (error) {
    logProductionStationSupabaseError("delete", error, {
      order_id: station.order_id,
      station_id: station.id,
      station_no: station.station_no,
    });
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Spannung konnte nicht gelöscht werden",
        "Spannung konnte nicht gelöscht werden: Datensatz ist noch verknüpft.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionOrderStationsFromSupabase();
  await ensureProductionOrderStationsForOrder(order.id);
  setProductionStatus("Spannung wurde gelöscht.");
  render();
}

async function createProductionDepartment() {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Abteilungen anlegen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = {
    name:
      document.getElementById("productionNewDepartmentName")?.value?.trim() ||
      "",
    code:
      document.getElementById("productionNewDepartmentCode")?.value?.trim() ||
      "",
    leaderEmployeeId:
      document.getElementById("productionNewDepartmentLeader")?.value || "",
    active:
      document.getElementById("productionNewDepartmentActive")?.checked !== false,
  };
  const validationMessage = validateProductionDepartmentInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const duplicateCheck = await findDepartmentByCode(values.code);
  if (duplicateCheck.error) {
    setProductionStatus(
      formatProductionSupabaseError(
        duplicateCheck.error,
        "Abteilungscode konnte nicht geprüft werden",
      ),
      true,
    );
    render();
    return;
  }
  if (duplicateCheck.exists) {
    setProductionStatus("Dieser Abteilungscode ist bereits vorhanden.", true);
    render();
    return;
  }

  const { error } = await supabaseClient.from("departments").insert([
    {
      name: values.name,
      code: values.code,
      leader_employee_id: values.leaderEmployeeId || null,
      active: values.active,
    },
  ]);

  if (error) {
    console.error("Fehler beim Anlegen der Abteilung:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Abteilung konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshDepartmentsFromSupabase();
  setProductionStatus("Abteilung wurde gespeichert.");
  render();
}

async function saveProductionDepartment(id) {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Abteilungen bearbeiten.", true);
    render();
    return;
  }
  if (!id) {
    setProductionStatus("Abteilung konnte nicht gespeichert werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = readProductionDepartmentForm(id);
  const validationMessage = validateProductionDepartmentInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const duplicateCheck = await findDepartmentByCode(values.code, id);
  if (duplicateCheck.error) {
    setProductionStatus(
      formatProductionSupabaseError(
        duplicateCheck.error,
        "Abteilungscode konnte nicht geprüft werden",
      ),
      true,
    );
    render();
    return;
  }
  if (duplicateCheck.exists) {
    setProductionStatus("Dieser Abteilungscode ist bereits vorhanden.", true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("departments")
    .update({
      name: values.name,
      code: values.code,
      leader_employee_id: values.leaderEmployeeId || null,
      updated_at: new Date().toISOString(),
    })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Speichern der Abteilung:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Abteilung konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshDepartmentsFromSupabase();
  setProductionStatus("Abteilung wurde gespeichert.");
  render();
}

async function setProductionDepartmentActive(id, active) {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Abteilungen aktivieren oder deaktivieren.", true);
    render();
    return;
  }
  if (!id) {
    setProductionStatus("Abteilung konnte nicht geändert werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("departments")
    .update({ active, updated_at: new Date().toISOString() })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Ändern der Abteilung:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Abteilung konnte nicht geändert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshDepartmentsFromSupabase();
  setProductionStatus(active ? "Abteilung wurde aktiviert." : "Abteilung wurde deaktiviert.");
  render();
}

function deactivateProductionDepartment(id) {
  setProductionDepartmentActive(id, false);
}

function activateProductionDepartment(id) {
  setProductionDepartmentActive(id, true);
}

async function deleteProductionDepartment(id) {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Abteilungen löschen.", true);
    render();
    return;
  }
  if (!id) {
    setProductionStatus("Abteilung konnte nicht gelöscht werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  if (
    !confirm(
      "Abteilung wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.",
    )
  ) {
    setProductionStatus("Löschen abgebrochen.");
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("departments")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Löschen der Abteilung:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Abteilung konnte nicht gelöscht werden",
        "Abteilung konnte nicht gelöscht werden: Datensatz ist noch verknüpft.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshDepartmentsFromSupabase();
  setProductionStatus("Abteilung wurde gelöscht.");
  render();
}

async function createProductionMachine() {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Maschinen anlegen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = {
    name:
      document.getElementById("productionNewMachineName")?.value?.trim() || "",
    machineCode:
      document.getElementById("productionNewMachineCode")?.value?.trim() || "",
    departmentId:
      document.getElementById("productionNewMachineDepartment")?.value || "",
    active:
      document.getElementById("productionNewMachineActive")?.checked !== false,
  };
  const validationMessage = validateProductionMachineInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const duplicateCheck = await findProductionMachineByCode(values.machineCode);
  if (duplicateCheck.error) {
    setProductionStatus(
      formatProductionSupabaseError(
        duplicateCheck.error,
        "Maschinencode konnte nicht geprüft werden",
      ),
      true,
    );
    render();
    return;
  }
  if (duplicateCheck.exists) {
    setProductionStatus("Dieser Maschinencode ist bereits vorhanden.", true);
    render();
    return;
  }

  const { error } = await supabaseClient.from("production_machines").insert([
    {
      name: values.name,
      machine_code: values.machineCode,
      department_id: values.departmentId || null,
      active: values.active,
    },
  ]);

  if (error) {
    console.error("Fehler beim Anlegen der Maschine:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Maschine konnte nicht gespeichert werden",
        "Dieser Maschinencode ist bereits vorhanden.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionMachinesFromSupabase();
  setProductionStatus("Maschine wurde gespeichert.");
  render();
}

async function saveProductionMachine(id) {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Maschinen bearbeiten.", true);
    render();
    return;
  }
  if (!id) {
    setProductionStatus("Maschine konnte nicht gespeichert werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = readProductionMachineForm(id);
  const validationMessage = validateProductionMachineInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const duplicateCheck = await findProductionMachineByCode(values.machineCode, id);
  if (duplicateCheck.error) {
    setProductionStatus(
      formatProductionSupabaseError(
        duplicateCheck.error,
        "Maschinencode konnte nicht geprüft werden",
      ),
      true,
    );
    render();
    return;
  }
  if (duplicateCheck.exists) {
    setProductionStatus("Dieser Maschinencode ist bereits vorhanden.", true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_machines")
    .update({
      name: values.name,
      machine_code: values.machineCode,
      department_id: values.departmentId || null,
      updated_at: new Date().toISOString(),
    })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Speichern der Maschine:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Maschine konnte nicht gespeichert werden",
        "Dieser Maschinencode ist bereits vorhanden.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionMachinesFromSupabase();
  setProductionStatus("Maschine wurde gespeichert.");
  render();
}

async function setProductionMachineActive(id, active) {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Maschinen aktivieren oder deaktivieren.", true);
    render();
    return;
  }
  if (!id) {
    setProductionStatus("Maschine konnte nicht geändert werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_machines")
    .update({ active, updated_at: new Date().toISOString() })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Ändern der Maschine:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Maschine konnte nicht geändert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionMachinesFromSupabase();
  setProductionStatus(active ? "Maschine wurde aktiviert." : "Maschine wurde deaktiviert.");
  render();
}

function deactivateProductionMachine(id) {
  setProductionMachineActive(id, false);
}

function activateProductionMachine(id) {
  setProductionMachineActive(id, true);
}

async function deleteProductionMachine(id) {
  if (!canManageProduction()) {
    setProductionStatus("Nur Admin darf Maschinen löschen.", true);
    render();
    return;
  }
  if (!id) {
    setProductionStatus("Maschine konnte nicht gelöscht werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  if (
    !confirm(
      "Maschine wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.",
    )
  ) {
    setProductionStatus("Löschen abgebrochen.");
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_machines")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Löschen der Maschine:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Maschine konnte nicht gelöscht werden",
        "Maschine konnte nicht gelöscht werden: Datensatz ist noch verknüpft.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionMachinesFromSupabase();
  setProductionStatus("Maschine wurde gelöscht.");
  render();
}

function adjustProductionCountField(inputId, delta) {
  const input = document.getElementById(inputId);
  if (!input) return;
  const currentValue = Math.max(0, Number(input.value || 0));
  input.value = String(Math.max(0, currentValue + Number(delta || 0)));
}

function updateProductionCountDepartmentPreview(selectId, targetId) {
  const select = document.getElementById(selectId);
  const target = document.getElementById(targetId);
  if (!select || !target) return;
  target.textContent = getProductionCountDepartmentDisplay(select.value);
}

function resetNewProductionCountForm() {
  if (!confirm("Stückzahl-Zähler wirklich zurücksetzen?")) return;
  const fieldIds = [
    "productionNewCountMachine",
    "productionNewCountOrder",
    "productionNewCountArticle",
    "productionNewCountGood",
    "productionNewCountScrap",
    "productionNewCountNote",
  ];
  fieldIds.forEach((fieldId) => {
    const field = document.getElementById(fieldId);
    if (!field) return;
    if (fieldId === "productionNewCountGood" || fieldId === "productionNewCountScrap") {
      field.value = "0";
    } else {
      field.value = "";
    }
  });
  updateProductionCountDepartmentPreview(
    "productionNewCountMachine",
    "productionNewCountDepartmentPreview",
  );
}

function updateProductionOrderDepartmentPreview(selectId, targetId) {
  const select = document.getElementById(selectId);
  const target = document.getElementById(targetId);
  if (!select || !target) return;
  target.textContent = getProductionOrderDepartmentDisplay(select.value);
}

async function createProductionOrder() {
  if (!canAccessProduction()) {
    setProductionStatus("Du darfst keine Aufträge / BA anlegen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = readProductionOrderForm();
  const validationMessage = validateProductionOrderInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const now = new Date().toISOString();
  const employeeId = currentEmployeeRecord?.id || null;
  const { error } = await supabaseClient.from("production_orders").insert([
    {
      machine_id: values.machineId,
      department_id: values.departmentId || null,
      ba_number: values.baNumber,
      article_number: values.articleNumber,
      ba_quantity: values.baQuantity,
      target_quantity: values.targetQuantity,
      pallet_count: values.palletCount,
      pieces_per_pallet: values.piecesPerPallet,
      use_chain_logic: values.useChainLogic,
      status: "running",
      started_at: now,
      completed_at: null,
      created_by_employee_id: employeeId,
      updated_by_employee_id: employeeId,
      updated_at: now,
    },
  ]);

  if (error) {
    console.error("Fehler beim Anlegen des Produktionsauftrags:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Auftrag / BA konnte nicht gespeichert werden",
        "Diese BA-Nummer ist bereits vorhanden.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionOrdersFromSupabase();
  setProductionStatus("Auftrag / BA wurde angelegt.");
  render();
}

async function updateProductionOrderStatus(id, status) {
  const order = getProductionOrderById(id);
  if (!canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diesen Auftrag / BA nicht bearbeiten.", true);
    render();
    return;
  }
  if (!["running", "paused", "completed", "cancelled"].includes(status)) {
    setProductionStatus("Ungültiger Auftragsstatus.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const now = new Date().toISOString();
  const payload = {
    status,
    updated_at: now,
    updated_by_employee_id: currentEmployeeRecord?.id || null,
    completed_at: ["completed", "cancelled"].includes(status) ? now : null,
  };

  const { error } = await supabaseClient
    .from("production_orders")
    .update(payload)
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Ändern des Auftragsstatus:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Auftragsstatus konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionOrdersFromSupabase();
  setProductionStatus("Auftragsstatus wurde gespeichert.");
  render();
}

function pauseProductionOrder(id) {
  updateProductionOrderStatus(id, "paused");
}

function resumeProductionOrder(id) {
  updateProductionOrderStatus(id, "running");
}

function completeProductionOrder(id) {
  completeProductionOrderWithChecklist(id);
}

function cancelProductionOrder(id) {
  if (!confirm("Auftrag / BA wirklich abbrechen?")) {
    setProductionStatus("Abbruch abgebrochen.");
    render();
    return;
  }
  updateProductionOrderStatus(id, "cancelled");
}

async function toggleProductionOrderChecklistItem(itemId, checked) {
  const item = getProductionOrderChecklistItemById(itemId);
  const order = getProductionOrderById(item?.order_id);
  if (!item || !order || !canEditProductionOrder(order) || !["running", "paused"].includes(order.status)) {
    setProductionStatus("Diese Abschluss-Checkliste kann nicht bearbeitet werden.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const now = new Date().toISOString();
  state.ui = state.ui || {};
  state.ui.productionChecklistSavingKey = item.id;
  render();

  const { error } = await supabaseClient
    .from("production_order_checklist")
    .update({
      checked: checked === true,
      checked_at: checked === true ? now : null,
      checked_by_employee_id: checked === true ? currentEmployeeRecord?.id || null : null,
      updated_at: now,
    })
    .eq("id", item.id);

  if (error) {
    logProductionOrderChecklistSupabaseError("update", error, order.id);
    state.ui.productionChecklistSavingKey = "";
    setProductionStatus(
      formatProductionSupabaseError(error, "Checklistenpunkt konnte nicht gespeichert werden"),
      true,
    );
    render();
    return;
  }

  const historyPayload = {
    item_key: item.item_key,
    item_label: item.item_label,
    checked: checked === true,
  };
  const { error: historyError } = await writeProductionOrderHistory(
    order,
    "checklist",
    historyPayload,
    0,
    null,
  );
  await refreshProductionOrderChecklistFromSupabase();
  state.ui.productionChecklistSavingKey = "";
  if (!historyError) await refreshProductionOrderHistoryFromSupabase();
  if (historyError) {
    console.error("Fehler beim Schreiben der Checklisten-Historie:", historyError);
    setProductionStatus("Checklistenpunkt gespeichert, Historie konnte nicht geschrieben werden.", true);
    render();
    return;
  }

  setProductionStatus(checked === true ? "Checklistenpunkt erledigt." : "Checklistenpunkt wieder geöffnet.");
  render();
}

function getProductionOrderCompletionBlocker(order) {
  if (!order || !canEditProductionOrder(order)) {
    return "Du darfst diesen Auftrag nicht fertig melden.";
  }
  if (!["running", "paused"].includes(order.status)) {
    return "Dieser Auftrag kann nicht fertig gemeldet werden.";
  }
  if (!(state.productionChecklistTemplates || []).length) {
    return "Keine Abschluss-Checkliste eingerichtet.";
  }
  if (productionOrderChecklistNeedsPreparation(order)) {
    return "Abschluss-Checkliste muss vorbereitet sein.";
  }
  if (!isProductionOrderChecklistComplete(order.id)) {
    return "Alle Checklistenpunkte müssen erledigt sein.";
  }
  if (!getProductionOrderStations(order.id).length) {
    return "Mindestens eine Spannung muss vorhanden sein.";
  }
  if (Number(order.target_quantity || 0) > 0 && calculatePreparedRemainingQuantity(order) > 0) {
    return "Auftrag hat noch Restmenge. Abschluss nicht möglich.";
  }
  return "";
}

function productionOrderChecklistNeedsPreparation(order) {
  if (!order || !["running", "paused"].includes(order.status)) return false;
  const templates = state.productionChecklistTemplates || [];
  if (!templates.length) return false;
  const existingKeys = new Set(
    getProductionOrderChecklistItems(order.id).map((item) => item.item_key),
  );
  return templates.some((template) => !existingKeys.has(template.item_key));
}

function scheduleProductionOrderChecklistPreparation(order) {
  if (!order || !productionOrderChecklistNeedsPreparation(order)) return;
  state.ui = state.ui || {};
  if (state.ui.productionChecklistPreparingOrderId === order.id) return;
  if (state.ui.productionChecklistAutoAttemptedOrderId === order.id) return;
  state.ui.productionChecklistAutoAttemptedOrderId = order.id;
  window.setTimeout(async () => {
    await ensureProductionOrderChecklistForOrder(order.id);
    render();
  }, 0);
}

async function openProductionCompletionDialog(orderId) {
  const order = getProductionOrderById(orderId);
  if (!order || !canEditProductionOrder(order)) {
    setProductionStatus("Du darfst diesen Auftrag nicht fertig melden.", true);
    render();
    return;
  }
  state.ui = state.ui || {};
  state.ui.productionCompletionOrderId = order.id;
  render();
  await ensureProductionOrderChecklistForOrder(order.id);
  render();
}

function closeProductionCompletionDialog() {
  state.ui = state.ui || {};
  state.ui.productionCompletionOrderId = "";
  render();
}

async function prepareProductionOrderChecklist(orderId) {
  state.ui = state.ui || {};
  state.ui.productionChecklistAutoAttemptedOrderId = "";
  const ok = await ensureProductionOrderChecklistForOrder(orderId);
  if (ok) setProductionStatus("Abschluss-Checkliste wurde vorbereitet.");
  render();
}

function renderProductionCompletionDialog() {
  const orderId = state.ui?.productionCompletionOrderId || "";
  if (!orderId) return "";
  const order = getProductionOrderById(orderId);
  if (!order) return "";
  const remaining = calculatePreparedRemainingQuantity(order);
  return `<div class='fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4'>
    <div class='bg-white rounded-lg shadow-xl max-w-4xl w-full max-h-[90vh] overflow-auto p-4 space-y-4'>
      <div class='flex items-start justify-between gap-3'>
        <div>
          <h3 class='text-lg font-semibold'>Auftrag fertig melden</h3>
          <p class='text-sm text-slate-500 mt-1'>BA ${escapeHtml(order.ba_number || "-")} · Restmenge ${escapeHtml(remaining)}</p>
        </div>
        <button type='button' class='px-3 py-2 rounded bg-slate-200 text-slate-800' onclick='closeProductionCompletionDialog()'>Abbrechen</button>
      </div>
      ${renderProductionOrderChecklistSection(order, true)}
    </div>
  </div>`;
}

async function completeProductionOrderWithChecklist(orderId) {
  const order = getProductionOrderById(orderId);
  const blocker = getProductionOrderCompletionBlocker(order);
  if (blocker) {
    setProductionStatus(blocker, true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  if (!confirm("Auftrag wirklich fertig melden?")) {
    setProductionStatus("Fertigmeldung abgebrochen.");
    render();
    return;
  }

  const now = new Date().toISOString();
  state.ui = state.ui || {};
  state.ui.productionOrderCompletingId = order.id;
  render();

  const payload = {
    status: "completed",
    completed_at: now,
    updated_at: now,
    updated_by_employee_id: currentEmployeeRecord?.id || null,
  };
  const { error } = await supabaseClient
    .from("production_orders")
    .update(payload)
    .eq("id", order.id);

  if (error) {
    console.error("Fehler beim Fertigmelden des Produktionsauftrags:", error);
    state.ui.productionOrderCompletingId = "";
    setProductionStatus(
      formatProductionSupabaseError(error, "Auftrag konnte nicht fertig gemeldet werden"),
      true,
    );
    render();
    return;
  }

  const historyPayload = {
    ba_number: order.ba_number || "",
    good_total: getProductionOrderGoodTotal(order.id),
    scrap_total: getProductionOrderScrapTotal(order.id),
    clarify_total: getProductionOrderClarifyTotal(order.id),
    remaining: calculatePreparedRemainingQuantity(order),
  };
  const { error: historyError } = await writeProductionOrderHistory(
    order,
    "order_completed",
    historyPayload,
    0,
    null,
  );
  await refreshProductionOrdersFromSupabase();
  if (!historyError) await refreshProductionOrderHistoryFromSupabase();
  state.ui.productionOrderCompletingId = "";
  state.ui.productionCompletionOrderId = "";
  if (state.ui.productionCounterActiveOrderId === order.id) {
    state.ui.productionCounterActiveOrderId = "";
    state.ui.productionCounterSelectedMachineId = "";
  }
  if (historyError) {
    console.error("Fehler beim Schreiben der Auftrags-Historie:", historyError);
    setProductionStatus("Auftrag fertig gemeldet, Historie konnte nicht geschrieben werden.", true);
    render();
    return;
  }

  setProductionStatus("Auftrag wurde fertig gemeldet.");
  render();
}

async function createProductionCount() {
  if (!canAccessProduction()) {
    setProductionStatus("Du darfst keine Stückzahl-Zähler anlegen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = readProductionCountValues("productionNewCount");
  const validationMessage = validateProductionCountInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const now = new Date().toISOString();
  const { error } = await supabaseClient.from("production_counts").insert([
    {
      machine_id: values.machineId,
      department_id: values.departmentId || null,
      employee_id: values.employeeId,
      order_no: values.orderNo,
      article_no: values.articleNo,
      good_qty: values.goodQty,
      scrap_qty: values.scrapQty,
      note: values.note,
      status: "running",
      started_at: now,
      updated_at: now,
    },
  ]);

  if (error) {
    console.error("Fehler beim Anlegen des Stückzahl-Zählers:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Stückzahl-Zähler konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionCountsFromSupabase();
  setProductionStatus("Stückzahl-Zähler wurde gespeichert.");
  render();
}

async function saveProductionCount(id) {
  const count = getProductionCountById(id);
  if (!canEditProductionCount(count) || count?.status === "completed") {
    setProductionStatus("Du darfst diesen Stückzahl-Zähler nicht bearbeiten.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = readProductionCountValues("productionCount", id);
  const validationMessage = validateProductionCountInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_counts")
    .update({
      machine_id: values.machineId,
      department_id: values.departmentId || null,
      employee_id: values.employeeId,
      order_no: values.orderNo,
      article_no: values.articleNo,
      good_qty: values.goodQty,
      scrap_qty: values.scrapQty,
      note: values.note,
      updated_at: new Date().toISOString(),
    })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Speichern des Stückzahl-Zählers:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Stückzahl-Zähler konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionCountsFromSupabase();
  setProductionStatus("Stückzahl-Zähler wurde gespeichert.");
  render();
}

async function completeProductionCount(id) {
  const count = getProductionCountById(id);
  if (!canEditProductionCount(count) || count?.status === "completed") {
    setProductionStatus("Du darfst diesen Stückzahl-Zähler nicht abschließen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  if (!confirm("Stückzahl-Zähler wirklich abschließen?")) {
    setProductionStatus("Abschließen abgebrochen.");
    render();
    return;
  }

  const values = readProductionCountValues("productionCount", id);
  const validationMessage = validateProductionCountInput(values);
  if (validationMessage) {
    setProductionStatus(validationMessage, true);
    render();
    return;
  }

  const now = new Date().toISOString();
  const { error } = await supabaseClient
    .from("production_counts")
    .update({
      machine_id: values.machineId,
      department_id: values.departmentId || null,
      employee_id: values.employeeId,
      order_no: values.orderNo,
      article_no: values.articleNo,
      good_qty: values.goodQty,
      scrap_qty: values.scrapQty,
      note: values.note,
      status: "completed",
      updated_at: now,
      completed_at: now,
    })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Abschließen des Stückzahl-Zählers:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Stückzahl-Zähler konnte nicht abgeschlossen werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionCountsFromSupabase();
  setProductionStatus("Stückzahl-Zähler wurde abgeschlossen.");
  render();
}

async function resetProductionCount(id) {
  const count = getProductionCountById(id);
  if (!canEditProductionCount(count) || count?.status === "completed") {
    setProductionStatus("Du darfst diesen Stückzahl-Zähler nicht zurücksetzen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setProductionStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  if (!confirm("Stückzahl-Zähler wirklich zurücksetzen?")) {
    setProductionStatus("Zurücksetzen abgebrochen.");
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("production_counts")
    .update({
      good_qty: 0,
      scrap_qty: 0,
      updated_at: new Date().toISOString(),
    })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Zurücksetzen des Stückzahl-Zählers:", error);
    setProductionStatus(
      formatProductionSupabaseError(
        error,
        "Stückzahl-Zähler konnte nicht zurückgesetzt werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshProductionCountsFromSupabase();
  setProductionStatus("Stückzahl-Zähler wurde zurückgesetzt.");
  render();
}

function canManagePersonnel() {
  return currentUser?.role === "admin";
}

function isPersonnelSystemUser(employee) {
  const role = String(employee?.role || "").trim().toLowerCase();
  const authUserId = employee?.authUserId || employee?.auth_user_id || "";
  return (
    role === "admin" ||
    role === "tool_scanner" ||
    role === "department_admin" ||
    !!authUserId ||
    isCurrentEmployee(employee)
  );
}

function isCurrentEmployee(employee) {
  return !!employee?.id && employee.id === currentEmployeeRecord?.id;
}

function getPersonnelEmployeeBadges(employee) {
  const badges = [];
  const role = String(employee?.role || "").trim().toLowerCase();
  if (role === "admin") badges.push("Admin");
  if (role === "tool_scanner") badges.push("Scanner");
  if (role === "department_admin") badges.push("Abteilungsleiter");
  if (employee?.authUserId || employee?.auth_user_id) badges.push("Login-Konto");
  if (
    role !== "department_admin" &&
    getDepartmentsLedByEmployee(employee?.id).length > 0
  ) {
    badges.push("Abteilungsleiter");
  }
  if (isPersonnelSystemUser(employee)) badges.push("Geschützt");

  return [...new Set(badges)];
}

function renderPersonnelEmployeeBadges(employee) {
  const badges = getPersonnelEmployeeBadges(employee);
  if (!badges.length) return "";
  return `<div class='flex flex-wrap gap-1 mt-2'>${badges
    .map((badge) => {
      const badgeClass =
        badge === "Geschützt"
          ? "bg-rose-100 text-rose-800 border border-rose-200"
          : "bg-slate-200 text-slate-700 border border-slate-200";
      return `<span class='px-2 py-0.5 rounded-full ${badgeClass} text-[11px] font-semibold'>${escapeHtml(badge)}</span>`;
    })
    .join("")}</div>`;
}

function setPersonalManagementStatus(message = "", isError = false) {
  state.ui = state.ui || {};
  if (isError) {
    state.ui.personalManagementActionError = message;
    state.ui.personalManagementActionMessage = "";
    return;
  }
  state.ui.personalManagementActionMessage = message;
  state.ui.personalManagementActionError = "";
}

function getPersonalManagementErrorBanner() {
  const employeeError = state.ui?.personalManagementEmployeesError;
  const departmentError = state.ui?.productionDepartmentsError;
  const shiftError = state.ui?.personalManagementShiftDefinitionsError;
  const actionError = state.ui?.personalManagementActionError;
  const actionMessage = state.ui?.personalManagementActionMessage;
  const messages = [];
  if (employeeError) {
    messages.push(`employees: ${employeeError}`);
  }
  if (departmentError) {
    messages.push(`departments: ${departmentError}`);
  }
  if (shiftError) {
    messages.push(`shift_definitions: ${shiftError}`);
  }
  if (actionError) {
    messages.push(actionError);
  }

  const errorBanner = messages.length
    ? `<div class='rounded border border-rose-200 bg-rose-50 p-3 text-sm text-rose-700'>
        ${messages.map(escapeHtml).join(" | ")}
      </div>`
    : "";
  const messageBanner = actionMessage
    ? `<div class='rounded border border-emerald-200 bg-emerald-50 p-3 text-sm text-emerald-700'>${escapeHtml(actionMessage)}</div>`
    : "";

  return `${errorBanner}${messageBanner}`;
}

function getPersonalManagementEmployeeRows() {
  return (state.employeesList || []).map((employee) => ({
    ...employee,
    first_name: employee.first_name || "",
    last_name: employee.last_name || "",
    personnel_no: employee.personnel_no || "",
    department: employee.department || "",
    department_id: employee.department_id || "",
    is_active: employee.is_active !== false,
  }));
}

function renderPersonalManagement() {
  if (!canManagePersonnel()) {
    return `<div class='bg-white rounded-xl shadow p-4'>
      <h2 class='text-lg font-semibold'>Personalverwaltung</h2>
      <p class='text-sm text-slate-600 mt-2'>Dieser Bereich ist nur für Administratoren verfügbar.</p>
    </div>`;
  }

  const subTab = PERSONNEL_MANAGEMENT_SUBTABS.some(
    (tab) => tab.id === state.personnelManagementSubTab,
  )
    ? state.personnelManagementSubTab
    : "employees";
  const subTabButtons = PERSONNEL_MANAGEMENT_SUBTABS.map((tab) => {
    const active = subTab === tab.id;
    return `<button class='px-3 py-2 rounded border ${active ? "humbel-subtab-active" : "humbel-subtab"}' onclick="setPersonnelManagementSubTab('${tab.id}')">${tab.label}</button>`;
  }).join("");

  let content = "";
  if (subTab === "employees") content = renderPersonnelEmployeesTab();
  if (subTab === "shiftModel") content = renderShiftDefinitionsTab();
  if (subTab === "shiftPlanning") {
    content = renderModulePlaceholder(
      "Schichtplanung",
      "Alter Schichtplan bleibt Alt / deaktiviert. Neue Schichtplanung wird später aufgebaut.",
    );
  }
  if (subTab === "settings") content = renderPersonnelSettingsTab();

  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-lg font-semibold'>Personalverwaltung</h2>
      <p class='text-sm text-slate-500 mt-1'>Neue Verwaltung für Mitarbeiter und Schichtmodell. Der alte Schichtplan wird hier fachlich nicht verwendet.</p>
    </div>
    <div class='flex gap-2 flex-wrap'>${subTabButtons}</div>
    ${getPersonalManagementErrorBanner()}
    ${content}
  </div>`;
}

function renderPersonnelEmployeesTab() {
  const employees = getPersonalManagementEmployeeRows();
  const activeDepartments = getActiveDepartments();
  const departmentInfo = !activeDepartments.length
    ? `<div class='rounded border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800'>Keine aktiven Abteilungen für die Auswahl vorhanden.</div>`
    : "";
  const rows = employees
    .map((employee) => {
      const isActive = employee.is_active !== false;
      const isSystemUser = isPersonnelSystemUser(employee);
      const effectiveDepartmentId = getEmployeeEffectiveDepartmentId(employee);
      const departmentDisplay = getDepartmentDisplayForEmployee(employee);
      const leaderHint = getDepartmentLeaderHint(employee);
      const statusClass = isActive
        ? "bg-emerald-100 text-emerald-800"
        : "bg-slate-200 text-slate-700";
      const activeToggleButton = isSystemUser
        ? ""
        : isActive
          ? `<button class='px-3 py-2 rounded bg-rose-700 text-white text-sm ml-2' onclick="deactivatePersonnelEmployee('${employee.id}')">Deaktivieren</button>`
          : `<button class='px-3 py-2 rounded bg-emerald-700 text-white text-sm ml-2' onclick="activatePersonnelEmployee('${employee.id}')">Aktivieren</button>`;
      const deleteButton = isSystemUser
        ? `<span class='inline-block px-3 py-2 rounded bg-slate-100 text-slate-600 text-xs font-semibold ml-2'>Geschützt</span>`
        : `<button class='px-3 py-2 rounded bg-red-800 text-white text-sm ml-2' onclick="deletePersonnelEmployee('${employee.id}')">Löschen</button>`;
      return `<tr class='border-b align-top ${isActive ? "" : "bg-slate-50 text-slate-500"}'>
        <td class='p-2'>
          <input id='pm-employee-first-${employee.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(employee.first_name)}' />
          ${renderPersonnelEmployeeBadges(employee)}
        </td>
        <td class='p-2'>
          <input id='pm-employee-last-${employee.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(employee.last_name)}' />
        </td>
        <td class='p-2'>
          <input id='pm-employee-no-${employee.id}' class='border rounded p-2 w-full bg-white' value='${escapeHtml(employee.personnel_no)}' />
        </td>
        <td class='p-2'>
          <select id='pm-employee-department-${employee.id}' class='border rounded p-2 w-full bg-white' onchange="autosavePersonnelEmployeeDepartment('${employee.id}', this.value)">${renderDepartmentOptions(effectiveDepartmentId)}</select>
          <div class='text-xs text-slate-500 mt-1'>${escapeHtml(departmentDisplay)}</div>
          ${leaderHint ? `<div class='text-xs text-amber-700 mt-1'>${escapeHtml(leaderHint)}</div>` : ""}
        </td>
        <td class='p-2 whitespace-nowrap'>
          <span class='px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${isActive ? "Aktiv" : "Inaktiv"}</span>
          ${renderPersonnelEmployeeBadges(employee)}
        </td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="savePersonnelEmployee('${employee.id}')">Speichern</button>
          ${activeToggleButton}
          ${deleteButton}
        </td>
      </tr>`;
    })
    .join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='font-semibold mb-3'>Mitarbeiter anlegen</h3>
      ${departmentInfo}
      <div class='grid sm:grid-cols-2 lg:grid-cols-5 gap-3'>
        <input id='pmNewFirstName' class='border rounded p-2 bg-white' placeholder='Vorname' />
        <input id='pmNewLastName' class='border rounded p-2 bg-white' placeholder='Nachname' />
        <input id='pmNewPersonnelNo' class='border rounded p-2 bg-white' placeholder='Personalnummer' />
        <select id='pmNewDepartmentId' class='border rounded p-2 bg-white'>${renderDepartmentOptions("")}</select>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='createPersonnelEmployee()'>Anlegen</button>
      </div>
    </div>
    <div class='border rounded-lg bg-white overflow-auto'>
      <table class='w-full text-sm min-w-[900px]'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Vorname</th>
            <th class='p-2 text-left'>Nachname</th>
            <th class='p-2 text-left'>Personalnummer</th>
            <th class='p-2 text-left'>Abteilung</th>
            <th class='p-2 text-left'>Status</th>
            <th class='p-2 text-left'>Aktion</th>
          </tr>
        </thead>
        <tbody>${rows || "<tr><td class='p-3 text-slate-500' colspan='6'>Keine Mitarbeiter geladen.</td></tr>"}</tbody>
      </table>
    </div>
  </div>`;
}

function getShiftDefinitionByKey(shiftKey) {
  return (state.shiftDefinitions || []).find(
    (definition) => definition.shift_key === shiftKey,
  );
}

function renderShiftDefinitionsTab() {
  const rows = SHIFT_DEFINITION_TYPES.map((template) => {
    const definition = getShiftDefinitionByKey(template.shift_key);
    const startTime = definition?.start_time || "";
    const endTime = definition?.end_time || "";
    const active = definition?.active !== false;
    const timeInputs = template.requires_time
      ? `<div class='grid sm:grid-cols-2 gap-3'>
          <label class='text-xs text-slate-500'>Zeit von
            <input id='shift-start-${template.shift_key}' type='time' class='mt-1 border rounded p-2 w-full bg-white' value='${escapeHtml(startTime)}' />
          </label>
          <label class='text-xs text-slate-500'>Zeit bis
            <input id='shift-end-${template.shift_key}' type='time' class='mt-1 border rounded p-2 w-full bg-white' value='${escapeHtml(endTime)}' />
          </label>
        </div>`
      : `<div class='grid sm:grid-cols-2 gap-3'>
          <label class='text-xs text-slate-500'>Zeit von
            <input type='time' class='mt-1 border rounded p-2 w-full bg-slate-100 text-slate-400' disabled value='' />
          </label>
          <label class='text-xs text-slate-500'>Zeit bis
            <input type='time' class='mt-1 border rounded p-2 w-full bg-slate-100 text-slate-400' disabled value='' />
          </label>
        </div>`;

    return `<div class='border rounded-lg p-3 bg-white'>
      <div class='flex items-start justify-between gap-3 mb-3'>
        <div>
          <h3 class='font-semibold'>${template.name}</h3>
          <p class='text-xs text-slate-500'>${template.requires_time ? "Zeitangaben erforderlich" : "Ohne Zeitangaben"}</p>
        </div>
        <label class='text-sm flex items-center gap-2'>
          <input id='shift-active-${template.shift_key}' type='checkbox' ${active ? "checked" : ""} />
          Aktiv
        </label>
      </div>
      ${timeInputs}
      <div class='flex justify-end mt-3'>
        <button class='px-3 py-2 rounded bg-slate-900 text-white text-sm' onclick="saveShiftDefinition('${template.shift_key}')">Speichern</button>
      </div>
    </div>`;
  }).join("");

  return `<div class='grid lg:grid-cols-2 gap-4'>${rows}</div>`;
}

function renderPersonnelSettingsTab() {
  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='font-semibold mb-2'>Einstellungen</h3>
      <p class='text-sm text-slate-600'>Für die Personalverwaltung sind hier später weitere Einstellungen vorgesehen.</p>
    </div>
  </div>`;
}

function readPersonnelEmployeeForm(id) {
  const employee = state.employees?.[id] || {};
  const selectedDepartmentId =
    document.getElementById(`pm-employee-department-${id}`)?.value || "";
  const autoDepartmentId =
    !employee.department_id && !selectedDepartmentId
      ? getSingleLedDepartmentId(id)
      : "";
  return {
    firstName:
      document.getElementById(`pm-employee-first-${id}`)?.value?.trim() || "",
    lastName:
      document.getElementById(`pm-employee-last-${id}`)?.value?.trim() || "",
    personnelNo:
      document.getElementById(`pm-employee-no-${id}`)?.value?.trim() || "",
    departmentId: selectedDepartmentId || autoDepartmentId,
  };
}

function validatePersonnelEmployeeInput(values, employee = null) {
  if (isPersonnelSystemUser(employee)) return "";
  if (!values.firstName || !values.lastName || !values.personnelNo) {
    return "Bitte Vorname, Nachname und Personalnummer ausfüllen.";
  }
  return "";
}

function buildPersonnelDisplayName(values) {
  return [values.firstName, values.lastName].filter(Boolean).join(" ").trim();
}

function buildPersonnelDisplayNameForPayload(values, employee = null) {
  return (
    buildPersonnelDisplayName(values) ||
    employee?.display_name ||
    employee?.name ||
    employee?.role ||
    ""
  );
}

function getDepartmentNameForStorage(departmentId) {
  const department = getDepartmentById(departmentId);
  return department?.name || "";
}

function getPersonnelDepartmentTextForPayload(departmentId, employee = null) {
  if (departmentId) return getDepartmentNameForStorage(departmentId) || null;
  if (employee?.department_id) return null;
  return employee?.department || null;
}

function formatPersonalManagementSupabaseError(
  error,
  fallback,
  duplicateMessage = "Diese Personalnummer ist bereits vorhanden.",
) {
  const message = error?.message || "";
  const code = error?.code || "";
  const lowerMessage = message.toLowerCase();
  if (
    code === "23505" ||
    lowerMessage.includes("duplicate") ||
    lowerMessage.includes("unique")
  ) {
    return duplicateMessage;
  }
  if (
    lowerMessage.includes("relation") ||
    lowerMessage.includes("does not exist") ||
    lowerMessage.includes("schema cache")
  ) {
    return `${fallback}: Supabase-Tabelle oder Spalte fehlt. Bitte Datenbankstruktur prüfen.`;
  }
  return `${fallback}: ${message || "Unbekannter Supabase-Fehler"}`;
}

async function findEmployeeByPersonnelNo(personnelNo, exceptId = null) {
  const { data, error } = await supabaseClient
    .from("employees")
    .select("id")
    .eq("personnel_no", personnelNo)
    .limit(1);

  if (error) return { error, exists: false };

  const existing = (data || []).find((row) => row.id !== exceptId);
  return { error: null, exists: !!existing };
}

async function createPersonnelEmployee() {
  if (!canManagePersonnel()) {
    setPersonalManagementStatus("Nur Admin darf Mitarbeiter anlegen.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setPersonalManagementStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = {
    firstName: document.getElementById("pmNewFirstName")?.value?.trim() || "",
    lastName: document.getElementById("pmNewLastName")?.value?.trim() || "",
    personnelNo: document.getElementById("pmNewPersonnelNo")?.value?.trim() || "",
    departmentId: document.getElementById("pmNewDepartmentId")?.value || "",
  };
  const validationMessage = validatePersonnelEmployeeInput(values);
  if (validationMessage) {
    setPersonalManagementStatus(validationMessage, true);
    render();
    return;
  }

  const duplicateCheck = await findEmployeeByPersonnelNo(values.personnelNo);
  if (duplicateCheck.error) {
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        duplicateCheck.error,
        "Personalnummer konnte nicht geprüft werden",
      ),
      true,
    );
    render();
    return;
  }
  if (duplicateCheck.exists) {
    setPersonalManagementStatus("Diese Personalnummer ist bereits vorhanden.", true);
    render();
    return;
  }

  const { error } = await supabaseClient.from("employees").insert([
    {
      first_name: values.firstName,
      last_name: values.lastName,
      display_name: buildPersonnelDisplayName(values),
      personnel_no: values.personnelNo,
      department_id: values.departmentId || null,
      department: getPersonnelDepartmentTextForPayload(values.departmentId),
      role: "employee",
      employee_type: "springer",
      is_active: true,
    },
  ]);

  if (error) {
    console.error("Fehler beim Anlegen des Mitarbeiters:", error);
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        error,
        "Mitarbeiter konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshEmployeesFromSupabase();
  setPersonalManagementStatus("Mitarbeiter wurde gespeichert.");
  render();
}

async function savePersonnelEmployee(id) {
  if (!canManagePersonnel()) {
    setPersonalManagementStatus("Nur Admin darf Mitarbeiter bearbeiten.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setPersonalManagementStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const values = readPersonnelEmployeeForm(id);
  const employee = state.employees?.[id] || {};
  const validationMessage = validatePersonnelEmployeeInput(values, employee);
  if (validationMessage) {
    setPersonalManagementStatus(validationMessage, true);
    render();
    return;
  }

  if (values.personnelNo) {
    const duplicateCheck = await findEmployeeByPersonnelNo(values.personnelNo, id);
    if (duplicateCheck.error) {
      setPersonalManagementStatus(
        formatPersonalManagementSupabaseError(
          duplicateCheck.error,
          "Personalnummer konnte nicht geprüft werden",
        ),
        true,
      );
      render();
      return;
    }
    if (duplicateCheck.exists) {
      setPersonalManagementStatus("Diese Personalnummer ist bereits vorhanden.", true);
      render();
      return;
    }
  }

  const { error } = await supabaseClient
    .from("employees")
    .update({
      first_name: values.firstName,
      last_name: values.lastName,
      display_name: buildPersonnelDisplayNameForPayload(values, employee),
      personnel_no: values.personnelNo || null,
      department_id: values.departmentId || null,
      department: getPersonnelDepartmentTextForPayload(values.departmentId, employee),
      updated_at: new Date().toISOString(),
    })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Speichern des Mitarbeiters:", error);
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        error,
        "Mitarbeiter konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshEmployeesFromSupabase();
  setPersonalManagementStatus("Mitarbeiter wurde gespeichert.");
  render();
}

async function autosavePersonnelEmployeeDepartment(id, departmentId) {
  if (!canManagePersonnel()) {
    setPersonalManagementStatus("Nur Admin darf die Abteilung ändern.", true);
    render();
    return;
  }
  if (!id) {
    setPersonalManagementStatus("Abteilung konnte nicht gespeichert werden: gültige Mitarbeiter-ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setPersonalManagementStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const normalizedDepartmentId = departmentId || "";
  const { error } = await supabaseClient
    .from("employees")
    .update({
      department_id: normalizedDepartmentId || null,
      department: normalizedDepartmentId
        ? getDepartmentNameForStorage(normalizedDepartmentId) || null
        : null,
      updated_at: new Date().toISOString(),
    })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim automatischen Speichern der Abteilung:", error);
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        error,
        "Abteilung konnte nicht gespeichert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshEmployeesFromSupabase();
  setPersonalManagementStatus("Abteilung gespeichert.");
  render();
}

async function deactivatePersonnelEmployee(id) {
  if (!canManagePersonnel()) {
    setPersonalManagementStatus("Nur Admin darf Mitarbeiter deaktivieren.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setPersonalManagementStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const employee = state.employees?.[id];
  const name = employee?.display_name || "Mitarbeiter";
  if (!employee?.id) {
    setPersonalManagementStatus("Mitarbeiter konnte nicht deaktiviert werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (isCurrentEmployee(employee)) {
    setPersonalManagementStatus("Der aktuell angemeldete Admin darf sich nicht selbst deaktivieren.", true);
    render();
    return;
  }
  if (isPersonnelSystemUser(employee)) {
    setPersonalManagementStatus("Login- und Systemuser dürfen hier nicht deaktiviert werden.", true);
    render();
    return;
  }
  if (!confirm(`${name} wirklich deaktivieren?`)) return;

  const { error } = await supabaseClient
    .from("employees")
    .update({ is_active: false, updated_at: new Date().toISOString() })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Deaktivieren des Mitarbeiters:", error);
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        error,
        "Mitarbeiter konnte nicht deaktiviert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshEmployeesFromSupabase();
  setPersonalManagementStatus("Mitarbeiter wurde deaktiviert.");
  render();
}

async function activatePersonnelEmployee(id) {
  if (!canManagePersonnel()) {
    setPersonalManagementStatus("Nur Admin darf Mitarbeiter aktivieren.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setPersonalManagementStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("employees")
    .update({ is_active: true, updated_at: new Date().toISOString() })
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Aktivieren des Mitarbeiters:", error);
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        error,
        "Mitarbeiter konnte nicht aktiviert werden",
      ),
      true,
    );
    render();
    return;
  }

  await refreshEmployeesFromSupabase();
  setPersonalManagementStatus("Mitarbeiter wurde aktiviert.");
  render();
}

async function deletePersonnelEmployee(id) {
  if (!canManagePersonnel()) {
    setPersonalManagementStatus("Nur Admin darf Mitarbeiter löschen.", true);
    render();
    return;
  }
  if (!id) {
    setPersonalManagementStatus("Mitarbeiter konnte nicht gelöscht werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setPersonalManagementStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }
  const employee = state.employees?.[id];
  if (!employee?.id) {
    setPersonalManagementStatus("Mitarbeiter konnte nicht gelöscht werden: gültige ID fehlt.", true);
    render();
    return;
  }
  if (isCurrentEmployee(employee)) {
    setPersonalManagementStatus("Der aktuell angemeldete Admin darf sich nicht selbst löschen.", true);
    render();
    return;
  }
  if (isPersonnelSystemUser(employee)) {
    setPersonalManagementStatus("Login- und Systemuser dürfen nicht gelöscht werden.", true);
    render();
    return;
  }
  if (
    !confirm(
      "Mitarbeiter wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.",
    )
  ) {
    setPersonalManagementStatus("Löschen abgebrochen.");
    render();
    return;
  }

  const { error } = await supabaseClient
    .from("employees")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Löschen des Mitarbeiters:", error);
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        error,
        "Mitarbeiter konnte nicht gelöscht werden",
        "Mitarbeiter konnte nicht gelöscht werden: Datensatz ist noch verknüpft.",
      ),
      true,
    );
    render();
    return;
  }

  await refreshEmployeesFromSupabase();
  setPersonalManagementStatus("Mitarbeiter wurde gelöscht.");
  render();
}

async function saveShiftDefinition(shiftKey) {
  if (!canManagePersonnel()) {
    setPersonalManagementStatus("Nur Admin darf Schichten bearbeiten.", true);
    render();
    return;
  }
  if (!supabaseReady) {
    setPersonalManagementStatus("Supabase ist nicht erreichbar.", true);
    render();
    return;
  }

  const template = SHIFT_DEFINITION_TYPES.find(
    (entry) => entry.shift_key === shiftKey,
  );
  if (!template) {
    setPersonalManagementStatus("Unbekannte Schichtart.", true);
    render();
    return;
  }

  const startTime = template.requires_time
    ? document.getElementById(`shift-start-${shiftKey}`)?.value || ""
    : "";
  const endTime = template.requires_time
    ? document.getElementById(`shift-end-${shiftKey}`)?.value || ""
    : "";
  const active = document.getElementById(`shift-active-${shiftKey}`)?.checked !== false;

  if (template.requires_time && (!startTime || !endTime)) {
    setPersonalManagementStatus(
      `${template.name}: Bitte Zeit von und Zeit bis eintragen.`,
      true,
    );
    render();
    return;
  }

  const payload = {
    name: template.name,
    shift_key: template.shift_key,
    start_time: template.requires_time ? startTime : null,
    end_time: template.requires_time ? endTime : null,
    requires_time: template.requires_time,
    active,
    updated_at: new Date().toISOString(),
  };

  const existing = getShiftDefinitionByKey(shiftKey);
  const query = existing?.id
    ? supabaseClient
        .from("shift_definitions")
        .update(payload)
        .eq("id", existing.id)
    : supabaseClient.from("shift_definitions").insert([payload]);
  const { error } = await query;

  if (error) {
    console.error("Fehler beim Speichern der Schichtdefinition:", error);
    setPersonalManagementStatus(
      formatPersonalManagementSupabaseError(
        error,
        "Schichtdefinition konnte nicht gespeichert werden",
        "Schichtdefinition existiert bereits.",
      ),
      true,
    );
    render();
    return;
  }

  const shiftDefinitions = await loadShiftDefinitionsFromSupabase();
  applyShiftDefinitionsToState(shiftDefinitions);
  setPersonalManagementStatus("Schichtdefinition wurde gespeichert.");
  render();
}

function getWeekRanges(from, to) {
  const ranges = [];
  let current = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T00:00:00`);

  while (current <= end) {
    const start = new Date(current);
    const weekday = (start.getDay() + 6) % 7;
    const monday = new Date(start);
    monday.setDate(start.getDate() - weekday);

    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);

    const rangeFrom =
      start > new Date(`${from}T00:00:00`)
        ? start
        : new Date(`${from}T00:00:00`);
    const rangeTo = sunday < end ? sunday : end;

    ranges.push({
      from: isoDate(rangeFrom),
      to: isoDate(rangeTo),
    });

    current = new Date(sunday);
    current.setDate(current.getDate() + 1);
  }

  return ranges;
}

function getShiftsOfUserInRange(userName, from, to) {
  return generateThreeMonths().filter((s) => {
    return s.date >= from && s.date <= to && s.originalAssigned === userName;
  });
}

function chooseReplacementUser(absentUser, from, to, weekLabel = "") {
  const available = activeUsers().filter((u) => u.name !== absentUser);

  if (!available.length) {
    alert("Keine anderen aktiven Mitarbeiter verfügbar.");
    return Promise.resolve(null);
  }

  const host = getModalHost();

  return new Promise((resolve) => {
    host.innerHTML = `
      <div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
          <h3 class="text-lg font-bold mb-2">Ersatz auswählen</h3>
          <p class="text-sm text-slate-700 mb-3">
            ${weekLabel ? `${weekLabel}<br>` : ""}
            Wer ersetzt <b>${absentUser}</b> von <b>${formatDateDisplay(from)}</b> bis <b>${formatDateDisplay(to)}</b>?
          </p>

          <select id="replacementUserSelect" class="border rounded p-2 w-full mb-4">
            ${available
              .map((u) => `<option value="${u.name}">${u.name}</option>`)
              .join("")}
            <option value="AUSFALL">AUSFALL</option>
          </select>

          <div class="flex justify-end gap-2">
            <button id="replacementCancel" class="px-3 py-1 rounded bg-slate-200">Abbrechen</button>
            <button id="replacementSave" class="px-3 py-1 rounded bg-slate-900 text-white">Speichern</button>
          </div>
        </div>
      </div>
    `;

    host.querySelector("#replacementCancel")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(null);
    });

    host.querySelector("#replacementSave")?.addEventListener("click", () => {
      const value = host.querySelector("#replacementUserSelect")?.value || "";

      host.innerHTML = "";

      if (value === "AUSFALL") {
        resolve({ mode: "cancel", replacementUser: null });
        return;
      }

      resolve({ mode: "replace", replacementUser: value });
    });
  });
}

async function planAbsenceReplacement(type, entryId) {
  if (!supabaseReady) return;
  openAbsenceReplacementPlanner(type, entryId);
}

async function deleteManualAbsence(date, userName) {
  if (!supabaseReady) return;

  const confirmDelete = confirm(
    `Abwesenheit von ${userName} am ${formatDateDisplay(date)} wirklich löschen?`,
  );
  if (!confirmDelete) return;

  const { error } = await supabaseClient
    .from("planner_absences")
    .delete()
    .eq("absence_date", date)
    .eq("user_name", userName);

  if (error) {
    console.error("Fehler beim Löschen der Abwesenheit:", error);
    return alert(`Fehler: ${error.message}`);
  }

  delete state.absences[`${date}:${userName}`];

  persist();
  render();
}

function renderPlanningAbstinenz() {
  const openShifts = generateThreeMonths().filter((s) => s.open);

  const monthValue = `${new Date().getFullYear()}-${String(
    new Date().getMonth() + 1,
  ).padStart(2, "0")}`;

  const personOptions = activeUsers()
    .map((u) => `<option value='${u.name}'>${u.name}</option>`)
    .join("");

  const absenceRows = Object.entries(state.absences || {})
    .slice(-100)
    .reverse()
    .map(([key, value]) => {
      const [date, user] = key.split(":");
      return `<tr class='border-b'>
        <td class='p-2'>${formatDateWithWeekday(date)}</td>
        <td class='p-2'>${user}</td>
        <td class='p-2'>Kann nicht kommen</td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-red-600 text-white' onclick="deleteManualAbsence('${date}','${user}')">Löschen</button>
        </td>
      </tr>`;
    })
    .join("");

  const vacationRows = state.vacations
    .slice(-20)
    .reverse()
    .map(
      (v) => `<tr class='border-b'>
        <td class='p-2'>${v.user}</td>
        <td class='p-2'>${formatDateDisplay(v.from)}</td>
        <td class='p-2'>${formatDateDisplay(v.to)}</td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-red-600 text-white' onclick="deleteVacation('${v.id}')">Löschen</button>
          <button class='px-2 py-1 rounded bg-slate-900 text-white ml-1' onclick="planAbsenceReplacement('vacation','${v.id}')">Ersatz</button>
        </td>
      </tr>`,
    )
    .join("");

  const sickRows = state.sickLeaves
    .slice(-20)
    .reverse()
    .map(
      (v) => `<tr class='border-b'>
        <td class='p-2'>${v.user}</td>
        <td class='p-2'>${formatDateDisplay(v.from)}</td>
        <td class='p-2'>${formatDateDisplay(v.to)}</td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-red-600 text-white' onclick="deleteSickLeave('${v.id}')">Löschen</button>
          <button class='px-2 py-1 rounded bg-slate-900 text-white ml-1' onclick="planAbsenceReplacement('sick','${v.id}')">Ersatz</button>
        </td>
      </tr>`,
    )
    .join("");

  const openRows = openShifts
    .map((s) => {
      const options = activeUsers()
        .map((u) => `<option value="${u.name}">${u.name}</option>`)
        .join("");

      return `<tr class='border-b'>
        <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
        <td class='p-2'>${s.label}</td>
        <td class='p-2'>${s.start}–${s.end}</td>
        <td class='p-2'>${s.originalAssigned || "-"}</td>
        <td class='p-2'>
          <select id='sel-${s.id}' class='border rounded p-1'>${options}</select>
        </td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-blue-700 text-white mr-2' onclick="assignShift('${s.id}')">Übernehmen</button>
          <button class='px-2 py-1 rounded bg-red-700 text-white' onclick="cancelShift('${s.id}')">Ausfall</button>
        </td>
      </tr>`;
    })
    .join("");

  return `
  <div class='space-y-4'>
    <div class='flex justify-end'>${helpButton("abstinenz")}</div>

    <div class='grid md:grid-cols-2 gap-4'>
      <div class='border rounded-lg p-3 bg-white'>
        <h3 class='font-semibold mb-2'>Klärung Abstinenz</h3>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100'>
            <tr>
              <th class='p-2'>Datum</th>
              <th class='p-2'>Mitarbeiter</th>
              <th class='p-2'>Typ</th>
              <th class='p-2'>Aktion</th>
            </tr>
          </thead>
          <tbody>
            ${absenceRows || "<tr><td class='p-2' colspan='4'>Keine Einträge</td></tr>"}
          </tbody>
        </table>
      </div>

      <div class='border rounded-lg p-3 bg-white'>
        <h3 class='font-semibold mb-3'>Urlaub / Krankmeldung eintragen</h3>

        <div class='grid grid-cols-2 gap-3 text-sm'>
          <label>
            Monat
            <input id='adminMonth' type='month' value='${monthValue}' class='border rounded p-2 w-full mt-1'/>
          </label>

          <label>
            Mitarbeiter
            <select id='adminCalUser' class='border rounded p-2 w-full mt-1'>
              ${personOptions}
            </select>
          </label>

          <label>
            Von
            <input id='adminFrom' type='date' value='${todayIso()}' class='border rounded p-2 w-full mt-1'/>
          </label>

          <label>
            Bis
            <input id='adminTo' type='date' value='${todayIso()}' class='border rounded p-2 w-full mt-1'/>
          </label>
        </div>

        <div class='flex gap-2 mt-4'>
          <button class='px-3 py-2 rounded bg-emerald-700 text-white' onclick='addVacation()'>Urlaub speichern</button>
          <button class='px-3 py-2 rounded bg-amber-700 text-white' onclick='addSickLeave()'>Krank speichern</button>
        </div>
      </div>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='font-semibold mb-2'>Geplante Urlaube</h3>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2'>Mitarbeiter</th>
            <th class='p-2'>Von</th>
            <th class='p-2'>Bis</th>
            <th class='p-2'>Aktion</th>
          </tr>
        </thead>
        <tbody>
          ${vacationRows || "<tr><td class='p-2' colspan='4'>Keine Einträge</td></tr>"}
        </tbody>
      </table>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='font-semibold mb-2'>Krankmeldungen</h3>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2'>Mitarbeiter</th>
            <th class='p-2'>Von</th>
            <th class='p-2'>Bis</th>
            <th class='p-2'>Aktion</th>
          </tr>
        </thead>
        <tbody>
          ${sickRows || "<tr><td class='p-2' colspan='4'>Keine Einträge</td></tr>"}
        </tbody>
      </table>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='font-semibold mb-2'>Offene Schichten</h3>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2'>Datum</th>
            <th class='p-2'>Schicht</th>
            <th class='p-2'>Zeit</th>
            <th class='p-2'>Vorher</th>
            <th class='p-2'>Übernehmen</th>
            <th class='p-2'>Aktion</th>
          </tr>
        </thead>
        <tbody>
          ${openRows || "<tr><td class='p-2' colspan='6'>Keine offenen Schichten</td></tr>"}
        </tbody>
      </table>
    </div>

  </div>`;
}

function isPrimaryCoreAbsentForShift(shift) {
  const primary = shift.originalAssigned;
  if (!primary || !isCoreEmployee(primary)) return false;

  const manualAbsent = !!state.absences?.[`${shift.date}:${primary}`];
  const calendarAbsent = !!getCalendarAbsenceType(primary, shift.date);

  return manualAbsent || calendarAbsent;
}

function getSuggestedSpringerForShift(shift) {
  if (!shift) return null;

  const availableSpringer = activeUsers()
    .filter((u) => u.type === "springer")
    .filter((u) => state.availability[`${shift.id}:${u.name}`] === "yes")
    .filter((u) => canAssignUserToShift(u.name, shift.id, true));

  if (!availableSpringer.length) return null;

  const shifts = generateThreeMonths();

  const ranked = availableSpringer
    .map((u) => {
      const weekendCount = shifts.filter(
        (s) =>
          s.assigned === u.name &&
          (s.id.includes("-sa-") || s.id.includes("-su-")),
      ).length;

      const totalCount = shifts.filter((s) => s.assigned === u.name).length;

      return {
        name: u.name,
        weekendCount,
        totalCount,
      };
    })
    .sort((a, b) => {
      if (a.weekendCount !== b.weekendCount) {
        return a.weekendCount - b.weekendCount;
      }
      if (a.totalCount !== b.totalCount) {
        return a.totalCount - b.totalCount;
      }
      return a.name.localeCompare(b.name);
    });

  return ranked[0]?.name || null;
}

function renderPlanningWochenende() {
  const weekendShifts = generateThreeMonths().filter(
    (s) => s.id.includes("-sa-1") || s.id.includes("-su-0"),
  );

  const optionalRows = weekendShifts
    .map((s) => {
      const primary = s.originalAssigned || "-";
      const primaryAbsent = isPrimaryCoreAbsentForShift(s);

      const yesUsers = activeUsers()
        .filter((u) => u.type === "springer")
        .filter((u) => state.availability[`${s.id}:${u.name}`] === "yes")
        .map((u) => u.name);

      const noUsers = activeUsers()
        .filter((u) => u.type === "springer")
        .filter((u) => state.availability[`${s.id}:${u.name}`] === "no")
        .map((u) => u.name);

      const suggested = getSuggestedSpringerForShift(s);

      const options = yesUsers
        .map(
          (name) =>
            `<option value='${name}' ${name === suggested ? "selected" : ""}>${name}${name === suggested ? " · Vorschlag" : ""}</option>`,
        )
        .join("");

      const canAssign = primaryAbsent && yesUsers.length > 0;

      return `<tr class='border-b'>
        <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
        <td class='p-2'>${s.label}</td>
        <td class='p-2'>${primary}</td>
        <td class='p-2'>
          ${
            primaryAbsent
              ? "<span class='px-2 py-1 rounded bg-emerald-100 text-emerald-800'>abwesend</span>"
              : "<span class='px-2 py-1 rounded bg-slate-100 text-slate-600'>nicht abwesend</span>"
          }
        </td>
        <td class='p-2'>${yesUsers.length ? yesUsers.join(", ") : "-"}</td>
        <td class='p-2'>${noUsers.length ? noUsers.join(", ") : "-"}</td>
        <td class='p-2 font-semibold'>
          ${
            suggested
              ? `<span class='px-2 py-1 rounded bg-blue-100 text-blue-800'>${suggested}</span>`
              : "-"
          }
        </td>
        <td class='p-2'>
          ${
            canAssign
              ? `<select id='opt-${s.id}' class='border rounded p-1'>${options}</select>`
              : `<select class='border rounded p-1 bg-slate-100 text-slate-400' disabled><option>-</option></select>`
          }
        </td>
        <td class='p-2 whitespace-nowrap'>
          ${
            canAssign
              ? `<button class='px-2 py-1 rounded bg-blue-700 text-white mr-2' onclick="assignOptionalShift('${s.id}')">Auswahl einteilen</button>
                 <button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick="assignSuggestedSpringer('${s.id}')">Vorschlag nehmen</button>`
              : `<span class='text-xs text-slate-500'>Nur bei Abwesenheit A/B/C</span>`
          }
        </td>
      </tr>`;
    })
    .join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-white'>
      <div class='flex items-center justify-between gap-2 mb-2'>
        <h3 class='text-md font-semibold'>Wochenende</h3>
        ${helpButton("wochenende")}
      </div>
      <p class='text-sm text-slate-500 mb-2'>
        Springer melden „Kann“ oder „Kann nicht“. Der Admin teilt ein. Der Vorschlag bevorzugt verfügbare Springer mit den wenigsten Wochenend-Einsätzen.
      </p>
      <div class='overflow-auto max-h-[60vh]'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100 sticky top-0'>
            <tr>
              <th class='p-2 text-left'>Datum</th>
              <th class='p-2 text-left'>Schicht</th>
              <th class='p-2 text-left'>A/B/C</th>
              <th class='p-2 text-left'>Status A/B/C</th>
              <th class='p-2 text-left'>Kann</th>
              <th class='p-2 text-left'>Kann nicht</th>
              <th class='p-2 text-left'>Vorschlag</th>
              <th class='p-2 text-left'>Auswahl</th>
              <th class='p-2 text-left'>Einteilen</th>
            </tr>
          </thead>
          <tbody>${optionalRows || '<tr><td class="p-2" colspan="9">Keine Wochenend-Schichten.</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  </div>`;
}

function renderPlanningSchichttausch() {
  const corePersonOptions = activeUsers()
    .filter((u) => u.type === "core")
    .map((u) => `<option value='${u.name}'>${u.name}</option>`)
    .join("");

  const swapRows = state.swaps
    .slice()
    .reverse()
    .map(
      (swap, index) => `<tr class='border-b'>
      <td class='p-2'>${swap.userA}</td>
      <td class='p-2'>${swap.userB}</td>
      <td class='p-2'>${formatDateDisplay(swap.startDate)}</td>
      <td class='p-2'>${swap.endDate ? formatDateDisplay(swap.endDate) : "-"}</td>
      <td class='p-2'>${!swap.endDate || swap.endDate >= todayIso() ? "Aktiv/Geplant" : "Beendet"}</td>
      <td class='p-2'><button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="deleteSwap(${state.swaps.length - 1 - index})">Löschen</button></td>
    </tr>`,
    )
    .join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <div class='flex items-center justify-between gap-2 mb-2'>
        <h3 class='text-md font-semibold'>Schichttausch</h3>
        ${helpButton("schichttausch")}
      </div>
      <div class='grid md:grid-cols-5 gap-2 items-end'>
        <label class='text-sm'>Mitarbeiter 1 (A/B/C)<select id='swapA' class='border rounded p-1 w-full'>${corePersonOptions}</select></label>
        <label class='text-sm'>Mitarbeiter 2 (A/B/C)<select id='swapB' class='border rounded p-1 w-full'>${corePersonOptions}</select></label>
        <label class='text-sm'>Gültig ab<input id='swapDate' type='date' value='${todayIso()}' class='border rounded p-1 w-full'/></label>
        <label class='text-sm'>Bis (optional)<input id='swapEndDate' type='date' class='border rounded p-1 w-full'/></label>
        <button class='px-2 py-1 rounded bg-purple-700 text-white h-9' onclick='addSwap()'>Schichttausch speichern</button>
      </div>
      <div class='mt-3 flex gap-2 flex-wrap'>
        <button class='px-2 py-1 rounded bg-slate-800 text-white text-sm' onclick='resetActiveSwaps()'>Tausch zurücksetzen (aktuelle+zukünftige)</button>
      </div>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h4 class='font-semibold mb-2'>Aktuelle und geplante Tausche</h4>
      <div class='overflow-auto max-h-[55vh]'>
        <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Mitarbeiter 1</th><th class='p-2 text-left'>Mitarbeiter 2</th><th class='p-2 text-left'>Ab</th><th class='p-2 text-left'>Bis</th><th class='p-2 text-left'>Status</th><th class='p-2'></th></tr></thead>
        <tbody>${swapRows || '<tr><td class="p-2" colspan="6">Keine Tausche vorhanden.</td></tr>'}</tbody></table>
      </div>
    </div>
  </div>`;
}

function canAssignUserToShift(name, shiftId, silent = false) {
  const shift = getShiftById(shiftId);
  if (!shift) return true;

  const absType = getCalendarAbsenceType(name, shift.date);
  if (absType) {
    if (!silent) {
      alert(
        `${name} ist am ${formatDateDisplay(shift.date)} als ${absType} gemeldet und kann nicht eingesetzt werden.`,
      );
    }
    return false;
  }

  const restCheck = checkRestTimeRule(name, shiftId);
  if (!restCheck.ok) {
    if (!silent) {
      alert(restCheck.message);
    }
    return false;
  }

  return true;
}

function checkRestTimeRule(name, shiftId) {
  const target = getShiftById(shiftId);
  if (!target) return { ok: true };
  const assigned = generateThreeMonths()
    .filter((s) => s.assigned === name && s.id !== shiftId)
    .map((s) => ({
      ...s,
      _start: shiftDateRange(s).start,
      _end: shiftDateRange(s).end,
    }));
  const tRange = shiftDateRange(target);
  assigned.push({
    ...target,
    assigned: name,
    _start: tRange.start,
    _end: tRange.end,
  });
  assigned.sort((a, b) => a._start - b._start);

  let shortRestCountWeek = 0;
  const targetWeek = weekKey(target.date);
  for (let i = 1; i < assigned.length; i++) {
    const prev = assigned[i - 1];
    const cur = assigned[i];
    const restHours =
      (cur._start.getTime() - prev._end.getTime()) / (1000 * 60 * 60);
    if (restHours < 8) {
      return {
        ok: false,
        message: `${name} hat zwischen Schichten weniger als 8 Stunden Ruhezeit.`,
      };
    }
    if (restHours < 11 && weekKey(cur.date) === targetWeek)
      shortRestCountWeek += 1;
  }
  if (shortRestCountWeek > 1) {
    return {
      ok: false,
      message: `${name} überschreitet die Ruhezeit-Regel: max. 1 Ausnahme mit 8h pro Woche.`,
    };
  }
  return { ok: true };
}

function weekKey(iso) {
  const d = new Date(`${iso}T00:00:00`);
  const weekday = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - weekday);
  return d.toISOString().slice(0, 10);
}

function readAdminCalendarForm() {
  const user = document.getElementById("adminCalUser")?.value;
  const from = document.getElementById("adminFrom")?.value;
  const to = document.getElementById("adminTo")?.value;
  if (!user || !from || !to || to < from) {
    alert("Bitte Mitarbeiter sowie gültiges Von/Bis-Datum angeben.");
    return null;
  }
  return { user, from, to };
}

async function addVacation() {
  const data = readAdminCalendarForm();
  if (!data) return;

  const payload = {
    user_name: data.user,
    from_date: data.from,
    to_date: data.to,
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("planner_vacations")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern von Urlaub:", error);
    return alert(`Urlaub konnte nicht gespeichert werden: ${error.message}`);
  }

  state.vacations.push({
    id: inserted.id,
    user: inserted.user_name,
    from: inserted.from_date,
    to: inserted.to_date,
  });

  persist();
  render();
}

async function addSickLeave() {
  const data = readAdminCalendarForm();
  if (!data) return;

  const payload = {
    user_name: data.user,
    from_date: data.from,
    to_date: data.to,
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("planner_sick_leaves")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern von Krankmeldung:", error);
    return alert(
      `Krankmeldung konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.sickLeaves.push({
    id: inserted.id,
    user: inserted.user_name,
    from: inserted.from_date,
    to: inserted.to_date,
  });

  persist();
  render();
}

function getDatesBetween(start, end) {
  const dates = [];
  let current = new Date(start);
  const last = new Date(end);

  while (current <= last) {
    dates.push(current.toISOString().slice(0, 10));
    current.setDate(current.getDate() + 1);
  }

  return dates;
}

async function deleteVacation(id) {
  if (!supabaseReady) return;

  const entry = state.vacations.find((v) => v.id === id);
  if (!entry) return;

  const confirmDelete = confirm(`Urlaub von ${entry.user} wirklich löschen?`);
  if (!confirmDelete) return;

  const replacementError = await deleteReplacementPlanFromSupabase(
    "vacation",
    id,
  );
  if (replacementError) {
    console.error("Fehler beim Löschen der Ersatzplanung:", replacementError);
    return alert(
      `Urlaub konnte nicht gelöscht werden, weil die Ersatzplanung nicht bereinigt werden konnte: ${replacementError.message}`,
    );
  }

  // 🔴 Urlaub löschen
  const { error } = await supabaseClient
    .from("planner_vacations")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Löschen Urlaub:", error);
    return alert(error.message);
  }

  // 🔴 Zugehörige Abwesenheiten löschen
  const dates = getDatesBetween(entry.from, entry.to);

  for (const d of dates) {
    await supabaseClient
      .from("planner_absences")
      .delete()
      .eq("absence_date", d)
      .eq("user_name", entry.user);

    delete state.absences[`${d}:${entry.user}`];
  }

  // 🔴 lokal entfernen
  clearReplacementPlanForSource("vacation", id);
  state.vacations = state.vacations.filter((v) => v.id !== id);

  persist();
  render();
}

async function deleteSickLeave(id) {
  if (!supabaseReady) return;

  const entry = state.sickLeaves.find((v) => v.id === id);
  if (!entry) return;

  const confirmDelete = confirm(
    `Krankmeldung von ${entry.user} wirklich löschen?`,
  );
  if (!confirmDelete) return;

  const replacementError = await deleteReplacementPlanFromSupabase("sick", id);
  if (replacementError) {
    console.error("Fehler beim Löschen der Ersatzplanung:", replacementError);
    return alert(
      `Krankmeldung konnte nicht gelöscht werden, weil die Ersatzplanung nicht bereinigt werden konnte: ${replacementError.message}`,
    );
  }

  const { error } = await supabaseClient
    .from("planner_sick_leaves")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Löschen Krankmeldung:", error);
    return alert(error.message);
  }

  const dates = getDatesBetween(entry.from, entry.to);

  for (const d of dates) {
    await supabaseClient
      .from("planner_absences")
      .delete()
      .eq("absence_date", d)
      .eq("user_name", entry.user);

    delete state.absences[`${d}:${entry.user}`];
  }

  clearReplacementPlanForSource("sick", id);
  state.sickLeaves = state.sickLeaves.filter((v) => v.id !== id);

  persist();
  render();
}

async function addSwap() {
  if (!supabaseReady) return;

  const userA = document.getElementById("swapA")?.value;
  const userB = document.getElementById("swapB")?.value;
  const startDate = document.getElementById("swapDate")?.value;
  const endDate = document.getElementById("swapEndDate")?.value || null;

  if (!isCoreEmployee(userA) || !isCoreEmployee(userB)) {
    alert("Beim Tausch sind nur A/B/C erlaubt.");
    return;
  }

  if (!userA || !userB || userA === userB || !startDate) {
    alert(
      "Bitte zwei verschiedene Mitarbeiter (A/B/C) und ein Startdatum wählen.",
    );
    return;
  }

  if (endDate && endDate < startDate) {
    alert("Enddatum muss nach dem Startdatum liegen.");
    return;
  }

  const { data: inserted, error } = await supabaseClient
    .from("planner_swaps")
    .insert({
      user_a: userA,
      user_b: userB,
      start_date: startDate,
      end_date: endDate,
    })
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern des Schichttausches:", error);
    return alert(
      `Schichttausch konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.swaps.push({
    id: inserted.id,
    userA: inserted.user_a,
    userB: inserted.user_b,
    startDate: inserted.start_date,
    endDate: inserted.end_date || null,
  });

  persist();
  render();
}

async function deleteSwap(index) {
  if (!supabaseReady) return;
  if (index < 0 || index >= state.swaps.length) return;

  const swap = state.swaps[index];

  if (swap?.id) {
    const { error } = await supabaseClient
      .from("planner_swaps")
      .delete()
      .eq("id", swap.id);

    if (error) {
      console.error("Fehler beim Löschen des Schichttausches:", error);
      return alert(
        `Schichttausch konnte nicht gelöscht werden: ${error.message}`,
      );
    }
  }

  state.swaps = state.swaps.filter((_, i) => i !== index);
  persist();
  render();
}

async function resetActiveSwaps(silent = false) {
  if (!supabaseReady) return;

  const today = todayIso();
  const y = new Date(`${today}T00:00:00`);
  y.setDate(y.getDate() - 1);
  const yesterday = isoDate(y);

  const activeSwaps = state.swaps.filter((swap) => {
    if (!swap) return false;
    if (swap.startDate >= today) return true;
    return !swap.endDate || swap.endDate >= today;
  });

  const deleteIds = activeSwaps
    .filter((swap) => swap.startDate >= today && swap.id)
    .map((swap) => swap.id);
  const updateIds = activeSwaps
    .filter((swap) => swap.startDate < today && swap.id)
    .map((swap) => swap.id);

  const operations = [];
  if (deleteIds.length)
    operations.push(
      supabaseClient.from("planner_swaps").delete().in("id", deleteIds),
    );
  if (updateIds.length)
    operations.push(
      supabaseClient
        .from("planner_swaps")
        .update({ end_date: yesterday })
        .in("id", updateIds),
    );

  const results = await Promise.all(operations);
  const firstError = results.find((res) => res.error)?.error;

  if (firstError) {
    console.error("Fehler beim Zurücksetzen der Tausche:", firstError);
    return alert(
      `Aktive/zukünftige Tausche konnten nicht zurückgesetzt werden: ${firstError.message}`,
    );
  }

  state.swaps = state.swaps.flatMap((swap) => {
    if (!swap) return [];
    if (swap.startDate >= today) return [];
    if (!swap.endDate || swap.endDate >= today)
      return [{ ...swap, endDate: yesterday }];
    return [swap];
  });

  persist();
  if (!silent) {
    render();
    alert("Aktive/zukünftige Tausche wurden zurückgesetzt.");
  }
}

function resetManualAssignments(silent = false) {
  const assignments = { ...state.assignments };
  Object.keys(assignments).forEach((shiftId) => {
    const shift = getShiftById(shiftId);
    if (!shift) return;
    if (isCurrentOrFutureShift(shift)) {
      delete state.assignments[shiftId];
      delete state.shiftCancellations[shiftId];
      delete state.absenceReplacements?.[shiftId];
    }
  });
  persist();
  if (!silent) {
    render();
    alert(
      "Manuelle Zuordnungen (aktuelle/zukünftige Schichten) wurden zurückgesetzt.",
    );
  }
}

async function resetPlanCurrentFuture() {
  if (
    !confirm(
      "Gesamtplan wirklich für aktuelle und zukünftige Schichten auf Ursprung zurücksetzen?",
    )
  ) {
    return;
  }

  if (!supabaseReady) {
    return alert("Supabase ist nicht bereit. Zurücksetzen nicht möglich.");
  }

  const today = todayIso();

  const deletes = await Promise.all([
    supabaseClient
      .from("planner_assignments")
      .delete()
      .gte("shift_date", today),
    supabaseClient.from("planner_absences").delete().gte("absence_date", today),
    supabaseClient.from("planner_vacations").delete().gte("to_date", today),
    supabaseClient.from("planner_sick_leaves").delete().gte("to_date", today),
    supabaseClient
      .from("planner_availability")
      .delete()
      .gte("shift_date", today),
    supabaseClient
      .from("planner_swaps")
      .delete()
      .or(`end_date.is.null,end_date.gte.${today}`),
    supabaseClient
      .from("planner_shift_cancellations")
      .delete()
      .gte("shift_date", today),
    supabaseClient
      .from("planner_saturday_requests")
      .delete()
      .gte("shift_date", today),
    supabaseClient
      .from("planner_absence_replacements")
      .delete()
      .gte("shift_date", today),
  ]);

  const firstError = deletes.find((res) => res.error)?.error;
  if (firstError) {
    console.error("Fehler beim Zurücksetzen des Plans:", firstError);
    return alert(
      `Gesamtplan konnte nicht vollständig zurückgesetzt werden: ${firstError.message}`,
    );
  }

  state.assignments = {};
  state.shiftCancellations = {};
  state.absences = {};
  state.swaps = [];
  state.absenceReplacements = {};
  state.saturdayEveningRequests = {};
  state.vacations = [];
  state.sickLeaves = [];
  state.availability = {};
  state.unmanned = {};
  state.shiftEndChecks = {};
  state.shiftStartChecks = {};
  state.checklists = {};
  state.conflicts = {};
  state.machineDowntime = {};
  state.machinePromptSeen = {};
  state.extraUsers = [];
  state.inactiveUsers = {};
  state.slotAssignments = { ...DEFAULT_SLOT_ASSIGNMENTS };
  state.planningSubTab = "personal";

  persist();
  render();
  alert(
    "Gesamtplan wurde für aktuelle und zukünftige Schichten zurückgesetzt.",
  );
}

async function refreshEmployeesFromSupabase() {
  const employees = await loadEmployeesFromSupabase();
  applyEmployeesToState(employees);

  if (
    currentEmployeeRecord?.id &&
    state.employees?.[currentEmployeeRecord.id]
  ) {
    const refreshed = state.employees[currentEmployeeRecord.id];
    currentEmployeeRecord = {
      ...currentEmployeeRecord,
      first_name: refreshed.first_name,
      last_name: refreshed.last_name,
      personnel_no: refreshed.personnel_no,
      department: refreshed.department,
      department_id: refreshed.department_id,
      display_name: refreshed.display_name,
      role: refreshed.role,
      employee_type: refreshed.employee_type,
      slot_code: refreshed.slot_code,
      is_active: refreshed.is_active,
      color_key: refreshed.color_key,
    };
    currentUser = {
      name: refreshed.display_name,
      role: refreshed.role,
    };
  }

  persist();
}

async function clearSlotInSupabase(slot, exceptEmployeeId = null) {
  if (!slot) return null;

  let query = supabaseClient
    .from("employees")
    .update({ slot_code: null })
    .eq("slot_code", slot);

  if (exceptEmployeeId) {
    query = query.neq("id", exceptEmployeeId);
  }

  const { error } = await query;
  return error || null;
}

async function updateSlotAssignment(slot) {
  const select = document.getElementById(`slot-${slot}`);
  if (!select) return;
  const selectedName = select.value;
  const selectedEmployee = activeUsers().find((u) => u.name === selectedName);

  if (supabaseReady && selectedEmployee?.id) {
    const clearError = await clearSlotInSupabase(slot, selectedEmployee.id);
    if (clearError) {
      console.error("Fehler beim Freigeben des Slots:", clearError);
      alert(clearError.message);
      return;
    }

    const { error } = await supabaseClient
      .from("employees")
      .update({ slot_code: slot })
      .eq("id", selectedEmployee.id);

    if (error) {
      console.error("Fehler beim Speichern der Slot-Zuordnung:", error);
      alert(error.message);
      return;
    }

    await refreshEmployeesFromSupabase();
  }

  state.slotAssignments[slot] = select.value;
  persist();
  render();
}

async function saveAllPersonnelChanges() {
  if (!supabaseReady) return;

  const ui = ensurePersonnelPendingState();
  const pendingEmployees = { ...ui.pendingEmployeeEdits };
  const pendingSlots = { ...ui.pendingSlotAssignments };
  const originalUsers = allUsers();
  const originalNameToId = {};
  originalUsers.forEach((user) => {
    if (user?.name && user?.id) originalNameToId[user.name] = user.id;
  });

  const failedEmployeeEdits = {};
  const failedSlotAssignments = {};
  const errors = [];

  for (const [id, changes] of Object.entries(pendingEmployees)) {
    const payload = {};

    if (Object.prototype.hasOwnProperty.call(changes, "display_name")) {
      const displayName = String(changes.display_name || "").trim();
      if (!displayName) {
        failedEmployeeEdits[id] = changes;
        errors.push("Ein Mitarbeitername fehlt.");
        continue;
      }
      payload.display_name = displayName;
    }

    if (Object.prototype.hasOwnProperty.call(changes, "employee_type")) {
      payload.employee_type = changes.employee_type || "springer";
    }
    if (Object.prototype.hasOwnProperty.call(changes, "color_key")) {
      payload.color_key = changes.color_key || "gray";
    }
    if (Object.prototype.hasOwnProperty.call(changes, "role")) {
      payload.role = changes.role || "employee";
    }
    if (Object.prototype.hasOwnProperty.call(changes, "is_active")) {
      payload.is_active = !!changes.is_active;
      if (!payload.is_active) payload.slot_code = null;
    }

    if (!Object.keys(payload).length) continue;

    const { error } = await supabaseClient
      .from("employees")
      .update(payload)
      .eq("id", id);

    if (error) {
      console.error("Fehler beim Speichern des Mitarbeiters:", error);
      failedEmployeeEdits[id] = changes;
      errors.push(error.message);
    }
  }

  await refreshEmployeesFromSupabase();

  for (const [slot, selectedName] of Object.entries(pendingSlots)) {
    const originalId = originalNameToId[selectedName];
    const selectedEmployee =
      activeUsers().find((u) => u.name === selectedName) ||
      activeUsers().find((u) => u.id === originalId);

    if (!selectedEmployee?.id) {
      failedSlotAssignments[slot] = selectedName;
      errors.push(`Kein aktiver Mitarbeiter für Slot ${slot} gefunden.`);
      continue;
    }

    const clearError = await clearSlotInSupabase(slot, selectedEmployee.id);
    if (clearError) {
      console.error("Fehler beim Freigeben des Slots:", clearError);
      failedSlotAssignments[slot] = selectedName;
      errors.push(clearError.message);
      continue;
    }

    const { error } = await supabaseClient
      .from("employees")
      .update({ slot_code: slot })
      .eq("id", selectedEmployee.id);

    if (error) {
      console.error("Fehler beim Speichern der Slot-Zuordnung:", error);
      failedSlotAssignments[slot] = selectedName;
      errors.push(error.message);
      continue;
    }

    state.slotAssignments[slot] = selectedEmployee.name;
  }

  ui.pendingEmployeeEdits = failedEmployeeEdits;
  ui.pendingSlotAssignments = failedSlotAssignments;

  await refreshEmployeesFromSupabase();
  persist();
  render();

  if (errors.length) {
    alert(`Einige Änderungen konnten nicht gespeichert werden:\n${errors.join("\n")}`);
    return;
  }

  alert("Alle Änderungen wurden gespeichert.");
}

async function addEmployee() {
  if (!supabaseReady) return;

  const name = document.getElementById("newEmployeeName")?.value?.trim() || "";
  const type = document.getElementById("newEmployeeType")?.value || "springer";
  const slot = document.getElementById("newEmployeeSlot")?.value || null;
  const colorKey = document.getElementById("newEmployeeColor")?.value || "gray";
  const role = document.getElementById("newEmployeeRole")?.value || "employee";

  if (!name) {
    alert("Name fehlt");
    return;
  }

  if (slot) {
    const clearError = await clearSlotInSupabase(slot);
    if (clearError) {
      console.error("Fehler beim Freigeben des Slots:", clearError);
      alert(clearError.message);
      return;
    }
  }

  const { error } = await supabaseClient.from("employees").insert([
    {
      display_name: name,
      role,
      employee_type: type,
      slot_code: slot,
      is_active: true,
      color_key: colorKey,
    },
  ]);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  document.getElementById("newEmployeeName").value = "";
  document.getElementById("newEmployeeSlot").value = "";
  await refreshEmployeesFromSupabase();
  render();
}

function fillEmployeeForm(emp) {
  if (!emp?.id) return;

  const nameInput = document.getElementById(`employee-name-${emp.id}`);
  const typeInput = document.getElementById(`employee-type-${emp.id}`);
  const colorInput = document.getElementById(`employee-color-${emp.id}`);
  const roleInput = document.getElementById(`employee-role-${emp.id}`);

  if (nameInput) nameInput.value = emp.display_name || emp.name || "";
  if (typeInput) typeInput.value = emp.employee_type || emp.type || "springer";
  if (colorInput) colorInput.value = emp.color_key || "gray";
  if (roleInput) roleInput.value = emp.role || "employee";
}

async function updateEmployee(id) {
  if (!supabaseReady) return;

  const name =
    document.getElementById(`employee-name-${id}`)?.value?.trim() || "";
  const type =
    document.getElementById(`employee-type-${id}`)?.value || "springer";
  const colorKey =
    document.getElementById(`employee-color-${id}`)?.value || "gray";
  const role =
    document.getElementById(`employee-role-${id}`)?.value || "employee";

  if (!name) {
    alert("Name fehlt");
    return;
  }

  const { error } = await supabaseClient
    .from("employees")
    .update({
      display_name: name,
      employee_type: type,
      color_key: colorKey,
      role,
    })
    .eq("id", id);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await refreshEmployeesFromSupabase();
  render();
}

async function deactivateEmployee(id) {
  const employee = state.employees?.[id];
  const name = employee?.display_name || employee?.name || "Mitarbeiter";
  if (!confirm(`${name} wirklich als inaktiv markieren?`)) return;
  if (supabaseReady && id) {
    const { error } = await supabaseClient
      .from("employees")
      .update({ is_active: false, slot_code: null })
      .eq("id", id);

    if (error) {
      console.error("Fehler beim Deaktivieren des Mitarbeiters:", error);
      alert(error.message);
      return;
    }
  }

  state.inactiveUsers[name] = true;
  Object.keys(state.slotAssignments).forEach((slot) => {
    if (state.slotAssignments[slot] === name)
      state.slotAssignments[slot] = DEFAULT_SLOT_ASSIGNMENTS[slot] || null;
  });
  if (supabaseReady) await refreshEmployeesFromSupabase();
  persist();
  render();
}

async function activateEmployee(id) {
  const employee = state.employees?.[id];
  const name = employee?.display_name || employee?.name || id;
  if (supabaseReady && id) {
    const { error } = await supabaseClient
      .from("employees")
      .update({ is_active: true })
      .eq("id", id);

    if (error) {
      console.error("Fehler beim Aktivieren des Mitarbeiters:", error);
      alert(error.message);
      return;
    }
  }

  delete state.inactiveUsers[name];
  if (supabaseReady) await refreshEmployeesFromSupabase();
  persist();
  render();
}

function getToolLabels() {
  return [...DEFAULT_TOOL_LABELS, ...(state.toolLabelsExtra || [])];
}

function getToolManufacturers() {
  return [
    ...DEFAULT_TOOL_MANUFACTURERS,
    ...(state.toolManufacturersExtra || []),
  ];
}

function getToolHolders() {
  return [...DEFAULT_TOOL_HOLDERS];
}

function addToolLabel() {
  if (currentUser.role !== "admin") return;
  const value = prompt("Neue Bezeichnung:");
  if (!value) return;
  if (!state.toolLabelsExtra.includes(value)) state.toolLabelsExtra.push(value);
  persist();
  render();
}

function addToolManufacturer() {
  if (currentUser.role !== "admin") return;
  const value = prompt("Neuer Hersteller:");
  if (!value) return;
  if (!state.toolManufacturersExtra.includes(value))
    state.toolManufacturersExtra.push(value);
  persist();
  render();
}

async function addToolMaterial() {
  if (currentUser.role !== "admin") return;

  const name = prompt("Neuen Schneidwerkstoff eingeben:");
  if (!name) return;

  const cleanName = name.trim();
  if (!cleanName) return;

  if (
    toolMaterials.some((m) => m.name.toLowerCase() === cleanName.toLowerCase())
  ) {
    return alert("Dieser Schneidwerkstoff existiert bereits.");
  }

  const nextSort =
    toolMaterials.reduce(
      (max, m) => Math.max(max, Number(m.sort_order || 0)),
      0,
    ) + 10;

  const { data, error } = await supabaseClient
    .from("tool_materials")
    .insert({
      name: cleanName,
      sort_order: nextSort,
      is_active: true,
    })
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Anlegen des Schneidwerkstoffs:", error);
    return alert(
      `Schneidwerkstoff konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  toolMaterials.push(data);
  toolMaterials.sort(
    (a, b) =>
      Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
      String(a.name || "").localeCompare(String(b.name || "")),
  );

  render();
}

async function renameToolMaterial(materialId) {
  if (currentUser.role !== "admin") return;

  const material = toolMaterials.find((m) => m.id === materialId);
  if (!material) return;

  const nextName = prompt("Schneidwerkstoff umbenennen:", material.name || "");
  if (!nextName) return;

  const cleanName = nextName.trim();
  if (!cleanName) return;

  if (
    toolMaterials.some(
      (m) =>
        m.id !== materialId &&
        String(m.name || "").toLowerCase() === cleanName.toLowerCase(),
    )
  ) {
    return alert("Dieser Schneidwerkstoff existiert bereits.");
  }

  const { data, error } = await supabaseClient
    .from("tool_materials")
    .update({
      name: cleanName,
      updated_at: new Date().toISOString(),
    })
    .eq("id", materialId)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Umbenennen des Schneidwerkstoffs:", error);
    return alert(
      `Schneidwerkstoff konnte nicht umbenannt werden: ${error.message}`,
    );
  }

  const index = toolMaterials.findIndex((m) => m.id === materialId);
  if (index !== -1) {
    toolMaterials[index] = data;
  }

  toolMaterials.sort(
    (a, b) =>
      Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
      String(a.name || "").localeCompare(String(b.name || "")),
  );

  render();
}

async function deactivateToolMaterial(materialId) {
  if (currentUser.role !== "admin") return;

  const material = toolMaterials.find((m) => m.id === materialId);
  if (!material) return;

  const linkedTools = state.tools.filter((t) => t.materialId === materialId);
  const message = linkedTools.length
    ? `Schneidwerkstoff "${material.name}" wirklich deaktivieren?\n\nEr ist noch bei ${linkedTools.length} Werkzeug(en) hinterlegt. Bestehende Werkzeuge behalten den Wert, aber der Werkstoff ist künftig nicht mehr auswählbar.`
    : `Schneidwerkstoff "${material.name}" wirklich deaktivieren?`;

  const ok = confirm(message);
  if (!ok) return;

  const { error } = await supabaseClient
    .from("tool_materials")
    .update({
      is_active: false,
      updated_at: new Date().toISOString(),
    })
    .eq("id", materialId);

  if (error) {
    console.error("Fehler beim Deaktivieren des Schneidwerkstoffs:", error);
    return alert(
      `Schneidwerkstoff konnte nicht deaktiviert werden: ${error.message}`,
    );
  }

  toolMaterials = toolMaterials.filter((m) => m.id !== materialId);
  render();
}

function renderToolMaterialsAdmin() {
  const rows = toolMaterials
    .map(
      (m) => `<tr class='border-b'>
        <td class='p-2'>${m.name || "-"}</td>
        <td class='p-2'>${Number(m.sort_order || 0)}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-2 py-1 rounded bg-amber-600 text-white mr-2' onclick="renameToolMaterial('${m.id}')">Bearbeiten</button>
          <button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="deactivateToolMaterial('${m.id}')">Deaktivieren</button>
        </td>
      </tr>`,
    )
    .join("");

  return `<div class='border-2 border-slate-300 rounded-xl p-3 bg-slate-50'>
    <div class='flex items-center justify-between gap-3 flex-wrap mb-3'>
      <div>
        <h3 class='text-lg font-bold mb-1'>Schneidwerkstoffe</h3>
        <p class='text-sm text-slate-500'>Admin kann Schneidwerkstoffe anlegen, umbenennen und deaktivieren.</p>
      </div>
      <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolMaterial()'>Schneidwerkstoff hinzufügen</button>
    </div>

    <div class='overflow-auto border rounded-lg bg-white max-h-[30vh]'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Name</th>
            <th class='p-2 text-left'>Sortierung</th>
            <th class='p-2'></th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td class="p-2" colspan="3">Keine Schneidwerkstoffe vorhanden.</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;
}

function renderToolMasterDataAdmin() {
  const labelRows = getToolLabels()
    .map(
      (name) => `<tr class='border-b'>
        <td class='p-2'>${name}</td>
        <td class='p-2 text-slate-500'>Bezeichnung</td>
      </tr>`,
    )
    .join("");

  const manufacturerRows = getToolManufacturers()
    .map(
      (name) => `<tr class='border-b'>
        <td class='p-2'>${name}</td>
        <td class='p-2 text-slate-500'>Hersteller</td>
      </tr>`,
    )
    .join("");

  const materialRows = toolMaterials
    .map(
      (m) => `<tr class='border-b'>
        <td class='p-2'>${m.name || "-"}</td>
        <td class='p-2 text-slate-500'>Schneidwerkstoff</td>
      </tr>`,
    )
    .join("");

  return `<div class='border-2 border-slate-300 rounded-xl p-3 bg-slate-50'>
    <div class='flex items-center justify-between gap-3 flex-wrap mb-3'>
      <div>
        <h3 class='text-lg font-bold mb-1'>Stammdaten für Werkzeuge</h3>
        <p class='text-sm text-slate-500'>Hier können Bezeichnungen, Hersteller und Schneidwerkstoffe ergänzt und eingesehen werden.</p>
      </div>
      <div class='flex gap-2 flex-wrap'>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolLabel()'>Bezeichnung hinzufügen</button>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolManufacturer()'>Hersteller hinzufügen</button>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolMaterial()'>Schneidwerkstoff hinzufügen</button>
      </div>
    </div>

    <div class='grid md:grid-cols-3 gap-4'>
      <div class='border rounded-lg bg-white overflow-auto max-h-[28vh]'>
        <div class='p-2 font-semibold border-b bg-slate-100'>Bezeichnungen</div>
        <table class='w-full text-sm'>
          <tbody>${labelRows || '<tr><td class="p-2">Keine Bezeichnungen vorhanden.</td></tr>'}</tbody>
        </table>
      </div>

      <div class='border rounded-lg bg-white overflow-auto max-h-[28vh]'>
        <div class='p-2 font-semibold border-b bg-slate-100'>Hersteller</div>
        <table class='w-full text-sm'>
          <tbody>${manufacturerRows || '<tr><td class="p-2">Keine Hersteller vorhanden.</td></tr>'}</tbody>
        </table>
      </div>

      <div class='border rounded-lg bg-white overflow-auto max-h-[28vh]'>
        <div class='p-2 font-semibold border-b bg-slate-100'>Schneidwerkstoffe</div>
        <table class='w-full text-sm'>
          <tbody>${materialRows || '<tr><td class="p-2">Keine Schneidwerkstoffe vorhanden.</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  </div>`;
}

function openToolMasterDataModal() {
  const labelRows = getToolLabels()
    .map((name) => `<tr class='border-b'><td class='p-2'>${escapeHtml(name)}</td></tr>`)
    .join("");

  const manufacturerRows = getToolManufacturers()
    .map((name) => `<tr class='border-b'><td class='p-2'>${escapeHtml(name)}</td></tr>`)
    .join("");

  const materialRows = toolMaterials
    .map(
      (material) =>
        `<tr class='border-b'><td class='p-2'>${escapeHtml(material.name || "-")}</td></tr>`,
    )
    .join("");

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl max-h-[88vh] overflow-auto p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">Stammdaten für Werkzeuge</h3>
          <p class="text-sm text-slate-500">Bezeichnungen, Hersteller und Schneidwerkstoffe.</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeToolMasterDataModal()">Schließen</button>
      </div>

      <div class='grid md:grid-cols-3 gap-4'>
        <div class='border rounded-lg bg-white overflow-auto max-h-[55vh]'>
          <div class='p-2 font-semibold border-b bg-slate-100'>Bezeichnungen</div>
          <table class='w-full text-sm'>
            <tbody>${labelRows || '<tr><td class="p-2">Keine Bezeichnungen vorhanden.</td></tr>'}</tbody>
          </table>
        </div>

        <div class='border rounded-lg bg-white overflow-auto max-h-[55vh]'>
          <div class='p-2 font-semibold border-b bg-slate-100'>Hersteller</div>
          <table class='w-full text-sm'>
            <tbody>${manufacturerRows || '<tr><td class="p-2">Keine Hersteller vorhanden.</td></tr>'}</tbody>
          </table>
        </div>

        <div class='border rounded-lg bg-white overflow-auto max-h-[55vh]'>
          <div class='p-2 font-semibold border-b bg-slate-100'>Schneidwerkstoffe</div>
          <table class='w-full text-sm'>
            <tbody>${materialRows || '<tr><td class="p-2">Keine Schneidwerkstoffe vorhanden.</td></tr>'}</tbody>
          </table>
        </div>
      </div>
    </div>
  </div>`;
}

function closeToolMasterDataModal() {
  getModalHost().innerHTML = "";
}

function openToolHistory(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const entries = (Array.isArray(state.toolJournal) ? state.toolJournal : [])
    .filter((entry) => entry.toolId === toolId)
    .sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));

  const rows = entries
    .map((entry) => {
      const action = entry.action || "-";
      const actionClass = String(action).toLowerCase().includes("entnahme")
        ? "text-rose-700"
        : String(action).toLowerCase().includes("einlager")
          ? "text-emerald-700"
          : "text-slate-700";

      return `<tr class='border-b'>
        <td class='p-2'>${escapeHtml(entry.createdAt ? new Date(entry.createdAt).toLocaleString() : "-")}</td>
        <td class='p-2'>${escapeHtml(entry.user || "-")}</td>
        <td class='p-2 ${actionClass}'>${escapeHtml(action)}</td>
        <td class='p-2'>${escapeHtml(entry.qty ?? "-")}</td>
        <td class='p-2'>${escapeHtml(entry.stockBefore ?? "-")}</td>
        <td class='p-2'>${escapeHtml(entry.stockAfter ?? "-")}</td>
      </tr>`;
    })
    .join("");

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-4xl max-h-[80vh] overflow-auto p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">Werkzeughistorie – T ${escapeHtml(tool.tNumber)}</h3>
          <p class="text-sm text-slate-500">${escapeHtml(tool.label || "-")} · ${escapeHtml(tool.holder || "-")} · Fach ${escapeHtml(tool.shelf || "-")}</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeToolHistory()">Schließen</button>
      </div>

      <div class='overflow-auto border rounded-lg'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100 sticky top-0'>
            <tr>
              <th class='p-2 text-left'>Zeit</th>
              <th class='p-2 text-left'>Benutzer</th>
              <th class='p-2 text-left'>Aktion</th>
              <th class='p-2 text-left'>Menge</th>
              <th class='p-2 text-left'>Bestand vorher</th>
              <th class='p-2 text-left'>Bestand nachher</th>
            </tr>
          </thead>
          <tbody>${rows || '<tr><td class="p-2" colspan="6">Keine Historie vorhanden.</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  </div>`;
}

function closeToolHistory() {
  getModalHost().innerHTML = "";
}

function openStorageLocationModal(locationKey) {
  const normalizedKey = normalizeStorageLocation(locationKey);
  const qrValue = buildStorageQrValue(normalizedKey);
  const storageMap = buildToolStorageMap();
  const tools = storageMap[normalizedKey] || [];
  const rows = tools
    .map(
      (tool) => `<tr class='border-b'>
        <td class='p-2'>T ${escapeHtml(tool.tNumber || "-")}</td>
        <td class='p-2'>${escapeHtml(tool.label || "-")}${renderToolBorrowedDetailLine(tool, "text-xs text-orange-700")}</td>
        <td class='p-2'>${escapeHtml(tool.stock ?? "-")}</td>
        <td class='p-2'>${escapeHtml(tool.holder || "-")}</td>
        <td class='p-2'>${escapeHtml(tool.manufacturer || "-")}</td>
        <td class='p-2 text-right'>
          <button class='px-2 py-1 rounded bg-blue-700 text-white text-xs' onclick="openMoveToolModal('${tool.id}')">Umlagern</button>
        </td>
      </tr>`,
    )
    .join("");

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-3xl max-h-[82vh] overflow-auto p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">Lagerfach ${escapeHtml(normalizedKey)}</h3>
          <p class="text-sm text-slate-500">${tools.length} Werkzeug(e)</p>
          <div class="mt-2 inline-flex flex-wrap items-center gap-2 rounded border bg-slate-50 px-3 py-2 text-sm">
            <span class="font-semibold text-slate-700">QR-Code:</span>
            <span class="font-mono font-semibold text-slate-900">${escapeHtml(qrValue)}</span>
            <button class="px-2 py-1 rounded bg-blue-700 text-white text-xs" onclick="openStorageQrModal('${normalizedKey}')">QR anzeigen</button>
          </div>
        </div>
        <div class="flex gap-2">
          <button class="px-3 py-1 rounded bg-slate-200" onclick="closeStorageLocationModal()">Schließen</button>
        </div>
      </div>
      <div class="overflow-auto">
        <table class="w-full text-sm">
          <thead class="bg-slate-100">
            <tr>
              <th class="p-2 text-left">T-Nummer</th>
              <th class="p-2 text-left">Bezeichnung</th>
              <th class="p-2 text-left">Bestand</th>
              <th class="p-2 text-left">Aufnahme</th>
              <th class="p-2 text-left">Hersteller</th>
              <th class="p-2 text-right">Aktion</th>
            </tr>
          </thead>
          <tbody>${rows || '<tr><td class="p-2" colspan="6">Kein Werkzeug in diesem Fach</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  </div>`;
}

function closeStorageLocationModal() {
  getModalHost().innerHTML = "";
}

async function openInventoryMode() {
  if (!INVENTORY_MODE_ENABLED) {
    alert("Inventur ist vorübergehend deaktiviert.");
    return;
  }

  if (currentUser?.role !== "admin") {
    alert("Inventurmodus ist nur für Administratoren verfügbar.");
    return;
  }

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
      <h3 class="text-xl font-bold mb-2">🧾 Inventurmodus</h3>
      <p class="text-sm text-slate-600">Inventur-Session wird gestartet ...</p>
    </div>
  </div>`;

  const { data: session, error } = await supabaseClient
    .from("inventory_sessions")
    .insert({
      user: currentUser?.name || currentEmployeeRecord?.display_name || "Admin",
      start_time: new Date().toISOString(),
      status: "open",
    })
    .select()
    .maybeSingle();

  if (error || !session?.id) {
    console.error("Inventur-Session konnte nicht gestartet werden", error);
    getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-xl font-bold mb-2">🧾 Inventurmodus</h3>
        <p class="text-sm text-rose-700 mb-4">Inventur-Session konnte nicht gestartet werden.</p>
        <button class="px-3 py-2 rounded bg-slate-200" onclick="closeInventoryMode()">Schließen</button>
      </div>
    </div>`;
    return;
  }

  activeInventorySessionId = session.id;
  activeInventoryLocationKey = "";
  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl max-h-[90vh] overflow-auto p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-xl font-bold">🧾 Inventurmodus</h3>
          <p class="text-sm text-slate-500">Reine Prüfansicht. Es werden keine Werkzeugdaten verändert.</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeInventoryMode()">Schließen</button>
      </div>

      <div class="border rounded-lg bg-slate-50 p-3 mb-4">
        <label class="text-sm font-medium">
          Fach erfassen
          <input id="inventoryLocationInput" class="border rounded p-2 w-full mt-1" placeholder="STORAGE:12N oder 12N" />
        </label>
        <div class="flex gap-2 mt-3">
          <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="processInventoryLocation()">Fach prüfen</button>
          <button class="px-3 py-2 rounded bg-slate-200" onclick="document.getElementById('inventoryLocationInput').value=''; document.getElementById('inventoryResult').innerHTML='';">Zurücksetzen</button>
        </div>
      </div>

      <div id="inventoryResult" class="space-y-3"></div>
    </div>
  </div>`;
  document.getElementById("inventoryLocationInput")?.focus();
}

async function closeInventoryMode() {
  const sessionId = activeInventorySessionId;
  activeInventorySessionId = null;
  activeInventoryLocationKey = "";

  if (!INVENTORY_MODE_ENABLED) {
    getModalHost().innerHTML = "";
    return;
  }

  if (sessionId) {
    const { error } = await supabaseClient
      .from("inventory_sessions")
      .update({
        end_time: new Date().toISOString(),
        status: "completed",
      })
      .eq("id", sessionId);
    if (error) {
      console.warn("Inventur-Session konnte nicht abgeschlossen werden", error);
    }
  }

  getModalHost().innerHTML = "";
}

function processInventoryLocation() {
  const input = document.getElementById("inventoryLocationInput");
  const result = document.getElementById("inventoryResult");
  if (!input || !result) return;

  if (!INVENTORY_MODE_ENABLED) {
    result.innerHTML = `<div class="border border-amber-200 bg-amber-50 text-amber-800 rounded p-3 text-sm">Inventur ist vorübergehend deaktiviert.</div>`;
    return;
  }

  const rawValue = String(input.value || "").trim();
  if (!activeInventorySessionId) {
    result.innerHTML = `<div class="border border-rose-200 bg-rose-50 text-rose-800 rounded p-3 text-sm">Keine aktive Inventur-Session.</div>`;
    return;
  }
  const locationKey = rawValue.toUpperCase().startsWith("STORAGE:")
    ? parseStorageQrCode(rawValue)
    : normalizeStorageLocation(rawValue);

  if (!locationKey || !isValidStorageLocation(locationKey)) {
    result.innerHTML = `<div class="border border-rose-200 bg-rose-50 text-rose-800 rounded p-3 text-sm">Ungültiges Lagerfach. Bitte STORAGE:12N oder 12N eingeben.</div>`;
    return;
  }

  const tools = (Array.isArray(state.tools) ? state.tools : []).filter((tool) => {
    return getToolStorageLocationKey(tool) === locationKey;
  });
  activeInventoryLocationKey = locationKey;

  const rows = tools
    .map((tool) => {
      const imagePath = getToolImagePath(tool);
      const imageCell = imagePath
        ? `<img src="${escapeHtml(imagePath)}" alt="Bild T ${escapeHtml(normalizeToolImageNumber(tool.tNumber))}" class="w-12 h-12 object-contain" onerror="this.style.display='none'" />`
        : "-";
      return `<tr class="border-b">
        <td class="p-2 whitespace-nowrap">T ${escapeHtml(tool.tNumber || "-")}</td>
        <td class="p-2">${escapeHtml(tool.label || "-")}</td>
        <td class="p-2">${imageCell}</td>
        <td class="p-2">${escapeHtml(tool.manufacturer || "-")}</td>
        <td class="p-2">${escapeHtml(tool.holder || "-")}</td>
        <td class="p-2 whitespace-nowrap">${renderToolSizeCell(tool)}</td>
        <td class="p-2 min-w-56">
          <input id="inventoryComment-${tool.id}" class="border rounded p-1 w-full text-xs mb-1" placeholder="Kommentar bei Abweichung" />
          <div class="flex gap-1">
            <button class="px-2 py-1 rounded bg-emerald-700 text-white text-xs" onclick="confirmInventoryResult('${tool.id}', 'present')">🟢 Vorhanden</button>
            <button class="px-2 py-1 rounded bg-rose-700 text-white text-xs" onclick="confirmInventoryResult('${tool.id}', 'deviation')">🔴 Abweichung</button>
          </div>
        </td>
      </tr>`;
    })
    .join("");

  result.innerHTML = `<div class="border rounded-lg bg-white p-3">
    <div class="flex items-start justify-between gap-3 mb-3">
      <div>
        <h4 class="font-bold text-lg">Fach: ${escapeHtml(locationKey)}</h4>
        <p class="text-sm text-slate-500">${tools.length} Soll-Werkzeug(e) laut System</p>
      </div>
    </div>

    ${
      tools.length
        ? `<div class="overflow-auto">
            <table class="w-full text-sm">
              <thead class="bg-slate-100">
                <tr>
                  <th class="p-2 text-left">T-Nummer</th>
                  <th class="p-2 text-left">Bezeichnung</th>
                  <th class="p-2 text-left">Bild</th>
                  <th class="p-2 text-left">Hersteller</th>
                  <th class="p-2 text-left">Aufnahme</th>
                  <th class="p-2 text-left">Ø / AL</th>
                  <th class="p-2 text-left">Ergebnis</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>`
        : '<div class="text-sm text-slate-600 bg-slate-50 border rounded p-3">Kein Werkzeug laut System eingelagert</div>'
    }

    <div id="inventoryFeedback" class="mt-3"></div>
  </div>`;
}

async function confirmInventoryResult(toolId, resultType) {
  const feedback = document.getElementById("inventoryFeedback");
  if (!feedback) return;

  if (!INVENTORY_MODE_ENABLED) {
    feedback.innerHTML = `<div class="border border-amber-200 bg-amber-50 text-amber-800 rounded p-3 text-sm">Inventur ist vorübergehend deaktiviert.</div>`;
    return;
  }

  const tool = (Array.isArray(state.tools) ? state.tools : []).find(
    (entry) => entry.id === toolId,
  );
  const fach = activeInventoryLocationKey;
  if (
    !activeInventorySessionId ||
    !fach ||
    !tool?.id ||
    getToolStorageLocationKey(tool) !== fach
  ) {
    feedback.innerHTML = `<div class="border border-rose-200 bg-rose-50 text-rose-800 rounded p-3 text-sm">Inventurergebnis konnte ohne gültige Session, Fach und Werkzeug nicht gespeichert werden.</div>`;
    return;
  }

  const comment = String(
    document.getElementById(`inventoryComment-${toolId}`)?.value || "",
  ).trim();
  const { error } = await supabaseClient.from("inventory_entries").insert({
    session_id: activeInventorySessionId,
    fach,
    tool_id: tool.id,
    ergebnis: resultType === "present" ? "vorhanden" : "abweichung",
    kommentar: comment || null,
    timestamp: new Date().toISOString(),
  });

  if (error) {
    console.error("Inventurergebnis konnte nicht gespeichert werden", error);
    feedback.innerHTML = `<div class="border border-rose-200 bg-rose-50 text-rose-800 rounded p-3 text-sm">Inventurergebnis konnte nicht gespeichert werden.</div>`;
    return;
  }

  const message =
    resultType === "present"
      ? `Inventurprüfung für T ${escapeHtml(tool.tNumber || "-")} als vorhanden gespeichert.`
      : `Inventurprüfung für T ${escapeHtml(tool.tNumber || "-")} als Abweichung gespeichert.`;
  const color =
    resultType === "present"
      ? "border-emerald-200 bg-emerald-50 text-emerald-800"
      : "border-rose-200 bg-rose-50 text-rose-800";

  feedback.innerHTML = `<div class="border ${color} rounded p-3 text-sm">${message}</div>`;
}

function openStorageQrModal(locationKey) {
  const normalizedKey = normalizeStorageLocation(locationKey);
  if (!isValidStorageLocation(normalizedKey)) {
    alert("Ungültiges Lagerfach");
    return;
  }

  const qrValue = buildStorageQrValue(normalizedKey);
  const qrSvg =
    renderSimpleQrSvg(qrValue) ||
    '<div class="text-sm text-rose-700">QR-Code konnte nicht erzeugt werden.</div>';
  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">QR-Code für Lagerfach ${escapeHtml(normalizedKey)}</h3>
          <p class="text-sm text-slate-500">Diesen QR-Code am Lagerfach anbringen.</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeStorageLocationModal()">Schließen</button>
      </div>

      <div class="rounded-lg border bg-slate-50 p-4 text-center">
        <div class="flex justify-center mb-3">${qrSvg}</div>
        <div class="font-mono text-sm font-semibold break-all">${escapeHtml(qrValue)}</div>
      </div>
      <div class="flex justify-end gap-2 mt-4">
        <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="printStorageQrLabel('${escapeHtml(normalizedKey)}')">Drucken</button>
      </div>
    </div>
  </div>`;
}

function printStorageQrLabel(locationKey) {
  const normalizedKey = normalizeStorageLocation(locationKey);
  if (!isValidStorageLocation(normalizedKey)) {
    alert("Ungültiges Lagerfach");
    return;
  }

  const qrValue = buildStorageQrValue(normalizedKey);
  const qrSvg = renderSimpleQrSvg(qrValue);
  const printWindow = window.open("", "_blank", "width=420,height=520");
  if (!printWindow) {
    alert("Druckfenster konnte nicht geöffnet werden.");
    return;
  }

  printWindow.document.write(`<!doctype html>
  <html>
    <head>
      <title>Lagerfach ${escapeHtml(normalizedKey)}</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 24px; color: #0f172a; }
        .label { border: 2px solid #0f172a; border-radius: 12px; padding: 18px; text-align: center; max-width: 340px; }
        .location { font-size: 38px; font-weight: 800; margin: 0 0 8px; }
        .meta { font-size: 13px; color: #475569; line-height: 1.3; }
        .qr-wrap { display: flex; justify-content: center; margin: 12px 0 10px; }
        .payload { font-family: monospace; font-size: 11px; word-break: break-all; }
        svg { width: 170px; height: 170px; }
      </style>
    </head>
    <body>
      <div class="label">
        <div class="location">${escapeHtml(normalizedKey)}</div>
        <div class="meta">Lagerfach</div>
        <div class="qr-wrap">${qrSvg}</div>
        <div class="payload">${escapeHtml(qrValue)}</div>
      </div>
      <script>
        window.onload = function() {
          window.print();
        };
      </script>
    </body>
  </html>`);
  printWindow.document.close();
}

function openMoveToolModal(toolId) {
  const tool = state.tools.find((item) => item.id === toolId);
  if (!tool) {
    alert("Werkzeug wurde nicht gefunden.");
    return;
  }

  const currentLocation = getToolStorageLocationKey(tool);

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-lg p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">Werkzeug umlagern</h3>
          <p class="text-sm text-slate-500">T ${escapeHtml(tool.tNumber || "-")} · ${escapeHtml(tool.label || "-")}</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeStorageLocationModal()">Abbrechen</button>
      </div>

      <div class="space-y-3">
        <div class="rounded border bg-slate-50 p-3 text-sm">
          <div><span class="font-semibold">Aktuelles Fach:</span> ${escapeHtml(currentLocation || "-")}</div>
          <div><span class="font-semibold">Bestand:</span> ${escapeHtml(tool.stock ?? "-")}</div>
        </div>

        <label class="block text-sm">
          Neues Fach
          <input id="moveToolLocation" class="border rounded p-2 w-full mt-1" placeholder="z. B. 12F" value="${escapeHtml(currentLocation || "")}" />
        </label>
        <p class="text-xs text-slate-500">Ein Fach darf nur von einem Werkzeug belegt sein.</p>

        <div class="flex justify-end gap-2">
          <button class="px-3 py-2 rounded bg-slate-200" onclick="closeStorageLocationModal()">Abbrechen</button>
          <button class="px-3 py-2 rounded bg-blue-700 text-white" onclick="confirmMoveTool('${tool.id}')">Umlagern</button>
        </div>
      </div>
    </div>
  </div>`;
}

async function confirmMoveTool(toolId) {
  const tool = state.tools.find((item) => item.id === toolId);
  if (!tool) {
    alert("Werkzeug wurde nicht gefunden.");
    return;
  }

  const input = document.getElementById("moveToolLocation");
  const newLocation = normalizeStorageLocation(input?.value || "");
  if (!isValidStorageLocation(newLocation)) {
    alert("Bitte ein gültiges Fach eingeben (Regal 1-26, Buchstabe A-Y, z. B. 12F).");
    return;
  }

  const oldLocation = getToolStorageLocationKey(tool);
  if (oldLocation === newLocation) {
    alert("Das Werkzeug liegt bereits in diesem Fach.");
    return;
  }

  const occupiedBy = findToolByStorageLocation(newLocation, toolId);
  if (occupiedBy) {
    alert(
      `Fach ${newLocation} ist bereits belegt durch T ${occupiedBy.tNumber} · ${occupiedBy.label || "-"}`,
    );
    return;
  }

  const { error } = await supabaseClient
    .from("tools")
    .update({ shelf: newLocation })
    .eq("id", tool.id);

  if (error) {
    alert("Werkzeug konnte nicht umgelagert werden: " + error.message);
    return;
  }

  state.tools = state.tools.map((item) => {
    if (item.id !== tool.id) return item;
    return {
      ...item,
      shelf: newLocation,
      location: item.location ? newLocation : item.location,
      storageLocation: item.storageLocation ? newLocation : item.storageLocation,
    };
  });

  await refreshToolPageData();
  await addToolJournalEntry({
    toolId: tool.id,
    toolLabel: tool.label,
    toolTNumber: tool.tNumber,
    action: `Werkzeug umgelagert ${oldLocation || "-"} -> ${newLocation}`,
    qty: 0,
    stockBefore: tool.stock,
    stockAfter: tool.stock,
    user: currentUser?.name || "System",
  });

  getModalHost().innerHTML = "";
  persist();
  render();
}

function getFilteredToolJournalEntries() {
  const filters = state.toolJournalFilters || {};
  const search = String(filters.search || "")
    .trim()
    .toLowerCase();
  const action = String(filters.action || "");
  const range = String(filters.range || "");

  let entries = Array.isArray(state.toolJournal) ? [...state.toolJournal] : [];

  if (search) {
    entries = entries.filter((entry) => {
      return [entry.toolTNumber, entry.toolLabel, entry.action, entry.user].some(
        (value) => String(value || "").toLowerCase().includes(search),
      );
    });
  }

  if (action) {
    entries = entries.filter((entry) =>
      String(entry.action || "").toLowerCase().includes(action.toLowerCase()),
    );
  }

  if (range) {
    const now = new Date();
    let since = null;

    if (range === "today") {
      since = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    } else if (range === "7d") {
      since = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    } else if (range === "30d") {
      since = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    }

    if (since) {
      entries = entries.filter((entry) => {
        const date = new Date(entry.createdAt);
        return !Number.isNaN(date.getTime()) && date >= since;
      });
    }
  }

  return entries;
}

function applyToolJournalFilters() {
  state.toolJournalFilters = {
    search: document.getElementById("toolJournalSearch")?.value || "",
    action: document.getElementById("toolJournalAction")?.value || "",
    range: document.getElementById("toolJournalRange")?.value || "",
  };
  persist();
  render();
}

function resetToolJournalFilters() {
  state.toolJournalFilters = {
    search: "",
    action: "",
    range: "",
  };
  persist();
  render();
}

function getToolStatisticsSinceDate(range) {
  const now = new Date();

  if (range === "today") {
    return new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }
  if (range === "7d") {
    return new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  }
  if (range === "30d") {
    return new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  }

  return null;
}

function buildToolStatistics() {
  const statsByTool = {};
  const range = state.toolStatisticsRange || "30d";
  const since = getToolStatisticsSinceDate(range);
  let entries = Array.isArray(state.toolJournal) ? state.toolJournal : [];

  if (since) {
    entries = entries.filter((entry) => {
      const entryDate = new Date(entry.createdAt || 0);
      return !Number.isNaN(entryDate.getTime()) && entryDate >= since;
    });
  }

  entries.forEach((entry) => {
    const toolId = entry.toolId || "";
    if (!toolId) return;

    if (!statsByTool[toolId]) {
      statsByTool[toolId] = {
        toolId,
        toolTNumber: entry.toolTNumber || "",
        toolLabel: entry.toolLabel || "",
        withdrawals: 0,
        restocks: 0,
        totalActions: 0,
        lastActionAt: null,
        currentStock: 0,
      };
    }

    const stat = statsByTool[toolId];
    const action = String(entry.action || "");
    if (action.includes("Entnahme")) stat.withdrawals += 1;
    if (action.includes("Einlagerung")) stat.restocks += 1;
    stat.totalActions += 1;

    const entryDate = new Date(entry.createdAt || 0);
    const lastDate = stat.lastActionAt ? new Date(stat.lastActionAt) : null;
    if (!Number.isNaN(entryDate.getTime()) && (!lastDate || entryDate > lastDate)) {
      stat.lastActionAt = entry.createdAt;
    }
  });

  return Object.values(statsByTool)
    .map((stat) => {
      const tools = Array.isArray(state.tools) ? state.tools : [];
      const tool = tools.find((t) => t.id === stat.toolId);
      return {
        ...stat,
        toolTNumber: stat.toolTNumber || tool?.tNumber || "",
        toolLabel: stat.toolLabel || tool?.label || "",
        currentStock: Number(tool?.stock || 0),
      };
    })
    .sort((a, b) => b.withdrawals - a.withdrawals);
}

function setToolStatisticsRange(range) {
  state.toolStatisticsRange = range || "30d";
  persist();
  render();
}

function isThreadToolLabel(label) {
  return [
    "Gewindebohrer",
    "Gewindefräser",
    "Gewindeformer",
    "Gewindewirbler",
  ].includes(label);
}

function formatToolSize(tool) {
  if (!isThreadToolLabel(tool.label)) return `⌀ ${tool.diameter}`;
  const prefix = tool.threadPrefix || "";
  const base = `${prefix}${prefix ? " " : ""}${tool.diameter}`;
  if (prefix === "MF" && tool.threadPitch)
    return `${base} P${tool.threadPitch}`;
  return base;
}

function getToolOverhangLength(tool) {
  return String(tool?.overhangLength || tool?.overhang_length || "").trim();
}

function renderToolSizeCell(tool) {
  const size = formatToolSize(tool).replace(/^⌀\s*/, "Ø");
  const radiusLine = tool.cornerRadius
    ? `<div>R${escapeHtml(tool.cornerRadius)}</div>`
    : "";
  const overhang = getToolOverhangLength(tool);
  const overhangLine = overhang
    ? `<div>AL${escapeHtml(overhang)}</div>`
    : '<div class="text-slate-400">—</div>';

  return `<div>${escapeHtml(size)}</div>${radiusLine}${overhangLine}`;
}

function renderToolOverhangDetailLine(tool, className = "") {
  const overhang = getToolOverhangLength(tool);
  if (!overhang) return "";

  const classAttr = className ? ` class="${className}"` : "";
  return `<div${classAttr}>AL: ${escapeHtml(overhang)} mm</div>`;
}

function renderToolBorrowedBadge(tool) {
  if (!tool?.isBorrowed) return "";

  const borrowedTo = tool.borrowedTo || "-";
  return `<span class="inline-flex px-2 py-1 rounded bg-orange-100 text-orange-800 text-xs font-semibold">🟠 Ausgeliehen: ${escapeHtml(borrowedTo)}</span>`;
}

function renderToolBorrowedDetailLine(tool, className = "") {
  if (!tool?.isBorrowed) return "";

  const classAttr = className ? ` class="${className}"` : "";
  const borrowedTo = tool.borrowedTo || "-";
  const borrowedAt = tool.borrowedAt
    ? `<div${classAttr}>Seit: ${escapeHtml(formatDateTime(tool.borrowedAt))}</div>`
    : "";
  return `<div${classAttr}>Ausgeliehen an: ${escapeHtml(borrowedTo)}</div>${borrowedAt}`;
}

function buildBorrowToolUpdateFields(borrowedTo) {
  return {
    is_borrowed: true,
    borrowed_to: borrowedTo,
    borrowed_at: new Date().toISOString(),
  };
}

function buildReturnToolUpdateFields(tool) {
  if (!tool?.isBorrowed) return {};

  return {
    is_borrowed: false,
    borrowed_to: null,
    borrowed_at: null,
  };
}
function isRadiusToolLabel(label) {
  return label === "Radiusfräser";
}

function updateToolTypeFields(prefix = "tool") {
  const labelEl = document.getElementById(`${prefix}Label`);
  const threadPrefixWrap = document.getElementById(`${prefix}ThreadPrefixWrap`);
  const threadPitchWrap = document.getElementById(`${prefix}ThreadPitchWrap`);
  const cornerRadiusWrap = document.getElementById(`${prefix}CornerRadiusWrap`);

  const label = labelEl?.value || "";
  const isThread = isThreadToolLabel(label);
  const isRadius = isRadiusToolLabel(label);

  if (threadPrefixWrap) threadPrefixWrap.style.display = isThread ? "" : "none";
  if (threadPitchWrap)
    threadPitchWrap.style.display =
      isThread &&
      labelEl?.value &&
      document.getElementById(`${prefix}ThreadPrefix`)?.value === "MF"
        ? ""
        : "none";
  if (cornerRadiusWrap) cornerRadiusWrap.style.display = isRadius ? "" : "none";
}

function updateThreadPitchVisibility(prefix = "tool") {
  const label = document.getElementById(`${prefix}Label`)?.value || "";
  const threadPrefix =
    document.getElementById(`${prefix}ThreadPrefix`)?.value || "";
  const threadPitchWrap = document.getElementById(`${prefix}ThreadPitchWrap`);

  const visible = isThreadToolLabel(label) && threadPrefix === "MF";
  if (threadPitchWrap) threadPitchWrap.style.display = visible ? "" : "none";
}

function collectToolFormData(root = document) {
  return {
    tNumber: root.getElementById("toolTNumber")?.value?.trim(),
    label: root.getElementById("toolLabel")?.value,
    diameter: root.getElementById("toolDiameter")?.value?.trim(),
    overhangLength:
      root.getElementById("toolOverhangLength")?.value?.trim() || "",
    threadPrefix: root.getElementById("toolThreadPrefix")?.value || "",
    threadPitch: root.getElementById("toolThreadPitch")?.value?.trim() || "",
    cornerRadius: root.getElementById("toolCornerRadius")?.value?.trim() || "",
    materialId: root.getElementById("toolMaterial")?.value || "",
    shelf: root.getElementById("toolShelf")?.value?.trim().toUpperCase(),
    articleNo: root.getElementById("toolArticle")?.value?.trim(),
    holder: root.getElementById("toolHolder")?.value,
    stock: Number(
      root.getElementById("toolStock")?.value ||
        root.querySelector?.("[name='stock']")?.value ||
        0,
    ),
    minStock: Number(root.getElementById("toolMinStock")?.value || 0),
    optimalStock: Number(root.getElementById("toolOptimalStock")?.value || 0),
    manufacturer: root.getElementById("toolManufacturer")?.value,
    insertTool: !!root.getElementById("toolInsertTool")?.checked,
    insertEdges: Number(root.getElementById("toolInsertEdges")?.value || 0),
    insertRadius:
      root.getElementById("toolInsertRadius")?.value?.trim() || "",
  };
}

function isValidToolShelf(value) {
  return /^\d{2}[A-Z]$/.test(String(value || "").trim().toUpperCase());
}

function validateToolData(data) {
  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  if (
    !data.tNumber ||
    !data.label ||
    !data.diameter ||
    !isValidToolShelf(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return {
      ok: false,
      message:
        "Bitte Felder korrekt ausfüllen (2-stellige Zahl + Buchstabe für Fach, z. B. 26C oder 02T; Aufnahme HSK 100 oder HSK 63).",
    };
  }

  if (isThreadTool && !data.threadPrefix) {
    return { ok: false, message: "Bitte Gewindekennung wählen." };
  }

  if (isThreadTool && data.threadPrefix === "MF" && !data.threadPitch) {
    return {
      ok: false,
      message: "Bitte bei MF die Steigung (P) angeben.",
    };
  }

  if (isRadiusTool && !data.cornerRadius) {
    return {
      ok: false,
      message: "Bitte Schneidenradius eingeben.",
    };
  }

  if (!isThreadTool && !Number.isFinite(Number(data.diameter))) {
    return {
      ok: false,
      message: "Bitte gültigen numerischen Durchmesser eingeben.",
    };
  }

  if (
    data.insertTool &&
    (!Number.isFinite(data.insertEdges) || data.insertEdges <= 0)
  ) {
    return {
      ok: false,
      message: "Bitte Anzahl der Schneiden > 0 eingeben.",
    };
  }

  return { ok: true };
}

function normalizeToolData(data) {
  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  return {
    tNumber: data.tNumber,
    label: data.label,
    diameter: isThreadTool ? data.diameter : String(data.diameter),
    overhangLength: data.overhangLength || "",
    threadPrefix: isThreadTool ? data.threadPrefix : "",
    threadPitch:
      isThreadTool && data.threadPrefix === "MF" ? data.threadPitch : "",
    cornerRadius: isRadiusTool ? data.cornerRadius : "",
    materialId: data.materialId || null,
    shelf: data.shelf,
    articleNo: data.articleNo,
    holder: data.holder,
    stock: Number(data.stock || 0),
    minStock: data.minStock,
    optimalStock: Math.max(0, data.optimalStock),
    manufacturer: data.manufacturer,
    ordered: false,
    orderedQty: 0,
    insertTool: data.insertTool,
    insertEdges: data.insertTool ? data.insertEdges : 0,
    insertRadius: data.insertTool ? data.insertRadius || "" : "",
  };
}

function renderToolCreateForm(prefix = "tool") {
  const labels = getToolLabels();
  const manufacturers = getToolManufacturers();
  const holders = getToolHolders();

  const labelOptions = labels
    .map((l) => `<option value="${l}">${l}</option>`)
    .join("");
  const manufacturerOptions = manufacturers
    .map((m) => `<option value="${m}">${m}</option>`)
    .join("");
  const holderOptions = holders
    .map((h) => `<option value="${h}">${h}</option>`)
    .join("");
  const materialOptions = toolMaterials
    .map((m) => `<option value="${m.id}">${m.name}</option>`)
    .join("");

  return `<div class='grid md:grid-cols-2 gap-3'>
    <input id='${prefix}TNumber' class='border rounded p-2' placeholder='T-Nummer (z.B. 134)' />

    <select id='${prefix}Label' class='border rounded p-2' onchange='updateToolTypeFields("${prefix}")'>
      ${labelOptions}
    </select>

    <input id='${prefix}Diameter' class='border rounded p-2' placeholder='Durchmesser' />
    <label class='text-sm'>
      Ausspannlänge (AL) [mm]
      <input id='${prefix}OverhangLength' type='number' step='0.1' min='0' class='border rounded p-2 w-full mt-1' placeholder='AL' />
    </label>

    <div id='${prefix}ThreadPrefixWrap' style='display:none;'>
      <select id='${prefix}ThreadPrefix' class='border rounded p-2 w-full' onchange='updateThreadPitchVisibility("${prefix}")'>
        <option value=''>Kennung (nur Gewinde)</option>
        <option value='M'>M</option>
        <option value='MF'>MF</option>
        <option value='G'>G</option>
        <option value='UNF'>UNF</option>
        <option value='UNC'>UNC</option>
        <option value='Mx'>Mx</option>
      </select>
    </div>

    <div id='${prefix}ThreadPitchWrap' style='display:none;'>
      <input id='${prefix}ThreadPitch' class='border rounded p-2 w-full' placeholder='Steigung P (nur MF)' />
    </div>

    <div id='${prefix}CornerRadiusWrap' style='display:none;'>
      <input id='${prefix}CornerRadius' class='border rounded p-2 w-full' placeholder='Schneidenradius' />
    </div>

    <select id='${prefix}Material' class='border rounded p-2'>
      <option value=''>Schneidwerkstoff wählen</option>
      ${materialOptions}
    </select>

    <input id='${prefix}Shelf' class='border rounded p-2' placeholder='00A' />
    <input id='${prefix}Article' class='border rounded p-2' placeholder='Artikel Nr.' />
    <select id='${prefix}Holder' class='border rounded p-2'>${holderOptions}</select>
    <input id='${prefix}Stock' type='number' class='border rounded p-2' placeholder='Bestand' />
    <input id='${prefix}MinStock' type='number' min='0' class='border rounded p-2' placeholder='Mindestbestand' />
    <input id='${prefix}OptimalStock' type='number' class='border rounded p-2' placeholder='Optimale Stückzahl' />
    <select id='${prefix}Manufacturer' class='border rounded p-2'>${manufacturerOptions}</select>
    <label class='flex items-center gap-2 text-sm md:col-span-2'>
      <input id='${prefix}InsertTool' type='checkbox' onchange='toggleInsertToolFieldsById("${prefix}InsertTool","${prefix}InsertEdges","${prefix}InsertRadius")' />
      Wendeplattenwerkzeug
    </label>
    <input id='${prefix}InsertEdges' type='number' class='border rounded p-2 md:col-span-2' placeholder='Anzahl Schneiden' disabled />
    <div id='${prefix}InsertRadiusWrap' class='md:col-span-2' style='display:none;'>
      <input id='${prefix}InsertRadius' class='border rounded p-2 w-full' placeholder='Plattenradius optional, z. B. 0.8' disabled />
    </div>
  </div>`;
}

async function createTool() {
  if (currentUser.role !== "admin")
    return alert("Nur Admin darf Werkzeuge anlegen.");

  const data = collectToolFormData(document);
  const validation = validateToolData(data);
  if (!validation.ok) return alert(validation.message);

  const normalized = normalizeToolData(data);

  const payload = {
    t_number: String(normalized.tNumber),
    label: normalized.label,
    diameter: String(normalized.diameter),
    overhang_length: normalized.overhangLength || null,
    thread_prefix: normalized.threadPrefix || null,
    thread_pitch: normalized.threadPitch || null,
    corner_radius: normalized.cornerRadius || null,
    material_id: normalized.materialId || null,
    shelf: normalized.shelf,
    article_no: normalized.articleNo,
    holder: normalized.holder,
    stock: Number(normalized.stock || 0),
    min_stock: Number(normalized.minStock || 0),
    optimal_stock: Number(normalized.optimalStock || 0),
    manufacturer: normalized.manufacturer || null,
    ordered: !!normalized.ordered,
    ordered_qty: Number(normalized.orderedQty || 0),
    insert_tool: !!normalized.insertTool,
    insert_edges: Number(normalized.insertEdges || 0),
    insert_radius: normalized.insertRadius || null,
    is_borrowed: false,
    borrowed_to: null,
    borrowed_at: null,
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("tools")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Anlegen des Werkzeugs in Supabase:", error);
    return alert(`Werkzeug konnte nicht gespeichert werden: ${error.message}`);
  }

  state.tools.push(normalizeToolFromDb(inserted));
  persist();
  render();
}

function getModalHost() {
  let host = document.getElementById("centerModalHost");
  if (!host) {
    host = document.createElement("div");
    host.id = "centerModalHost";
    document.body.appendChild(host);
  }
  return host;
}

function normalizeToolImageNumber(tNumber) {
  return String(tNumber || "")
    .trim()
    .replace(/^T\s*/i, "")
    .trim();
}

function getToolImagePath(tool) {
  const holderFolder =
    tool.holder === "HSK 100"
      ? "HSK100"
      : tool.holder === "HSK 63"
        ? "HSK63"
        : "";

  const imageNumber = normalizeToolImageNumber(tool.tNumber);

  if (!holderFolder || !imageNumber) return "";

  return `img/Depo/${holderFolder}/${imageNumber}.PNG`;
}

function openToolImagePopup(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const imagePath = getToolImagePath(tool);
  if (!imagePath) {
    alert("Für dieses Werkzeug konnte kein Bildpfad erstellt werden.");
    return;
  }

  const imageNumber = normalizeToolImageNumber(tool.tNumber);
  const host = getModalHost();

  host.innerHTML = `
    <div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-3xl max-h-[90vh] overflow-auto p-4">
        <div class="flex items-center justify-between gap-3 mb-3">
          <div>
            <h3 class="text-lg font-bold">Werkzeugbild – T ${escapeHtml(imageNumber)}</h3>
            <p class="text-sm text-slate-500">${escapeHtml(tool.label || "-")} · ${escapeHtml(tool.holder || "-")}</p>
            ${renderToolOverhangDetailLine(tool, "text-xs text-slate-500")}
            ${renderToolBorrowedDetailLine(tool, "text-xs text-orange-700")}
            <p class="text-xs text-slate-400">${escapeHtml(imagePath)}</p>
          </div>
          <button class="px-3 py-1 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
        </div>

        <div class="border rounded-lg bg-slate-50 p-3 flex justify-center">
          <img
            src="${imagePath}"
            alt="Werkzeugbild T ${escapeHtml(imageNumber)}"
            class="max-w-full max-h-[70vh] object-contain rounded"
            onerror="this.outerHTML='<div class=&quot;text-sm text-rose-700 bg-rose-50 border border-rose-200 rounded p-3 w-full&quot;>Kein Bild gefunden: ${imagePath}</div>'"
          />
        </div>
      </div>
    </div>
  `;
}

function closeToolImagePopup() {
  getModalHost().innerHTML = "";
}

function renderSimpleQrSvg(text) {
  try {
    const size = 29;
    const quiet = 4;
    const scale = 6;
    const dataCodewords = 55;
    const ecCodewords = 15;
    const modules = Array.from({ length: size }, () => Array(size).fill(false));
    const reserved = Array.from({ length: size }, () => Array(size).fill(false));

    const setModule = (x, y, dark, reserve = true) => {
      if (x < 0 || y < 0 || x >= size || y >= size) return;
      modules[y][x] = !!dark;
      if (reserve) reserved[y][x] = true;
    };

    const addFinder = (x, y) => {
      for (let dy = -1; dy <= 7; dy++) {
        for (let dx = -1; dx <= 7; dx++) {
          const xx = x + dx;
          const yy = y + dy;
          if (xx < 0 || yy < 0 || xx >= size || yy >= size) continue;
          const inFinder = dx >= 0 && dx <= 6 && dy >= 0 && dy <= 6;
          const dark =
            inFinder &&
            (dx === 0 ||
              dx === 6 ||
              dy === 0 ||
              dy === 6 ||
              (dx >= 2 && dx <= 4 && dy >= 2 && dy <= 4));
          setModule(xx, yy, dark);
        }
      }
    };

    addFinder(0, 0);
    addFinder(size - 7, 0);
    addFinder(0, size - 7);

    for (let i = 8; i < size - 8; i++) {
      setModule(i, 6, i % 2 === 0);
      setModule(6, i, i % 2 === 0);
    }

    for (let dy = -2; dy <= 2; dy++) {
      for (let dx = -2; dx <= 2; dx++) {
        const dist = Math.max(Math.abs(dx), Math.abs(dy));
        setModule(22 + dx, 22 + dy, dist === 2 || dist === 0);
      }
    }

    for (let i = 0; i <= 8; i++) {
      reserved[8][i] = true;
      reserved[i][8] = true;
    }
    for (let i = 0; i < 8; i++) {
      reserved[8][size - 1 - i] = true;
      reserved[size - 1 - i][8] = true;
    }
    setModule(8, size - 8, true);

    const bytes = String(text || "")
      .split("")
      .map((ch) => ch.charCodeAt(0));
    if (!bytes.length || bytes.some((byte) => byte > 255)) {
      throw new Error("QR unterstützt nur kurzen ASCII-Text.");
    }
    if (bytes.length > 53) {
      throw new Error("QR-Text ist für diese lokale Minimalversion zu lang.");
    }

    const bits = [];
    const appendBits = (value, length) => {
      for (let i = length - 1; i >= 0; i--) bits.push((value >>> i) & 1);
    };

    appendBits(0b0100, 4);
    appendBits(bytes.length, 8);
    bytes.forEach((byte) => appendBits(byte, 8));

    const capacityBits = dataCodewords * 8;
    const terminator = Math.min(4, capacityBits - bits.length);
    appendBits(0, terminator);
    while (bits.length % 8) bits.push(0);

    const data = [];
    for (let i = 0; i < bits.length; i += 8) {
      data.push(Number.parseInt(bits.slice(i, i + 8).join(""), 2));
    }
    for (let pad = 0xec; data.length < dataCodewords; pad ^= 0xfd) {
      data.push(pad);
    }

    const gfExp = Array(512).fill(0);
    const gfLog = Array(256).fill(0);
    let value = 1;
    for (let i = 0; i < 255; i++) {
      gfExp[i] = value;
      gfLog[value] = i;
      value <<= 1;
      if (value & 0x100) value ^= 0x11d;
    }
    for (let i = 255; i < 512; i++) gfExp[i] = gfExp[i - 255];

    const gfMul = (a, b) => {
      if (!a || !b) return 0;
      return gfExp[gfLog[a] + gfLog[b]];
    };

    const polyMul = (a, b) => {
      const result = Array(a.length + b.length - 1).fill(0);
      a.forEach((av, i) => {
        b.forEach((bv, j) => {
          result[i + j] ^= gfMul(av, bv);
        });
      });
      return result;
    };

    let generator = [1];
    for (let i = 0; i < ecCodewords; i++) {
      generator = polyMul(generator, [1, gfExp[i]]);
    }

    const ec = Array(ecCodewords).fill(0);
    data.forEach((byte) => {
      const factor = byte ^ ec.shift();
      ec.push(0);
      for (let i = 0; i < ecCodewords; i++) {
        ec[i] ^= gfMul(generator[i + 1], factor);
      }
    });

    const codewordBits = [...data, ...ec].flatMap((byte) => {
      const out = [];
      for (let i = 7; i >= 0; i--) out.push((byte >>> i) & 1);
      return out;
    });

    let bitIndex = 0;
    let upward = true;
    for (let right = size - 1; right >= 1; right -= 2) {
      if (right === 6) right--;
      for (let vert = 0; vert < size; vert++) {
        const y = upward ? size - 1 - vert : vert;
        for (let j = 0; j < 2; j++) {
          const x = right - j;
          if (reserved[y][x]) continue;
          const bit = bitIndex < codewordBits.length ? codewordBits[bitIndex++] : 0;
          const mask = (x + y) % 2 === 0;
          setModule(x, y, bit ^ mask, false);
        }
      }
      upward = !upward;
    }

    const formatData = (0b01 << 3) | 0;
    let formatRemainder = formatData << 10;
    for (let i = 14; i >= 10; i--) {
      if ((formatRemainder >>> i) & 1) {
        formatRemainder ^= 0x537 << (i - 10);
      }
    }
    const formatBits = ((formatData << 10) | formatRemainder) ^ 0x5412;
    const formatBit = (i) => ((formatBits >>> i) & 1) === 1;

    for (let i = 0; i <= 5; i++) setModule(8, i, formatBit(i));
    setModule(8, 7, formatBit(6));
    setModule(8, 8, formatBit(7));
    setModule(7, 8, formatBit(8));
    for (let i = 9; i < 15; i++) setModule(14 - i, 8, formatBit(i));
    for (let i = 0; i < 8; i++) setModule(size - 1 - i, 8, formatBit(i));
    for (let i = 8; i < 15; i++) setModule(8, size - 15 + i, formatBit(i));
    setModule(8, size - 8, true);

    const viewSize = (size + quiet * 2) * scale;
    const rects = [];
    modules.forEach((row, y) => {
      row.forEach((dark, x) => {
        if (!dark) return;
        rects.push(
          `<rect x="${(x + quiet) * scale}" y="${(y + quiet) * scale}" width="${scale}" height="${scale}"/>`,
        );
      });
    });

    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${viewSize} ${viewSize}" width="150" height="150" role="img" aria-label="QR-Code">
      <rect width="100%" height="100%" fill="#fff"/>
      <g fill="#000">${rects.join("")}</g>
    </svg>`;
  } catch (error) {
    console.error("QR-Code konnte nicht erzeugt werden:", error);
    return `<div class="text-sm text-rose-700 bg-rose-50 border border-rose-200 rounded p-3">[QR-Code konnte nicht erzeugt werden]</div>`;
  }
}

function getToolQrPayload(tool) {
  return `TOOL:${tool.id}`;
}

function openToolQrPopup(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const payload = getToolQrPayload(tool);
  const host = getModalHost();
  const toolTitle = `T ${tool.tNumber}`;
  const qrSvg = renderSimpleQrSvg(payload);
  const diameter = tool.diameter || "-";
  const cornerRadius = tool.cornerRadius ? `R ${tool.cornerRadius}` : "-";
  const manufacturer = tool.manufacturer || "-";
  const articleNo = tool.articleNo || "-";
  const shelf = tool.shelf || "-";
  const overhangLine = renderToolOverhangDetailLine(
    tool,
    "text-xs text-slate-600",
  );
  const insertRadiusLine =
    tool.insertTool && tool.insertRadius
      ? `<div class="text-xs text-slate-600">Plattenradius: R ${escapeHtml(tool.insertRadius)}</div>`
      : "";

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-3">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-2xl max-h-[90vh] overflow-auto p-3">
      <div class="flex items-center justify-between gap-3 mb-2">
        <div>
          <h3 class="text-lg font-bold">QR-Code für ${escapeHtml(toolTitle)}</h3>
          <p class="text-sm text-slate-500">${escapeHtml(tool.label || "-")} · ${escapeHtml(tool.holder || "-")} · Fach ${escapeHtml(tool.shelf || "-")}</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
      </div>
      <div class="grid md:grid-cols-[1fr,190px] gap-3 mb-3">
        <div class="border rounded-lg bg-slate-50 p-2 space-y-1">
          <div><span class="text-xs text-slate-500">T-Nummer</span><div class="font-semibold">${escapeHtml(toolTitle)}</div></div>
          <div><span class="text-xs text-slate-500">Bezeichnung</span><div>${escapeHtml(tool.label || "-")}</div></div>
          <div><span class="text-xs text-slate-500">Aufnahme</span><div>${escapeHtml(tool.holder || "-")}</div></div>
          <div><span class="text-xs text-slate-500">Lagerfach</span><div>${escapeHtml(tool.shelf || "-")}</div></div>
          ${renderToolOverhangDetailLine(tool)}
          ${renderToolBorrowedDetailLine(tool)}
          <div><span class="text-xs text-slate-500">QR-Inhalt</span><div class="font-mono text-xs break-all bg-white border rounded p-1 mt-1">${escapeHtml(payload)}</div></div>
        </div>
        <div class="border rounded-lg bg-white p-2 flex items-center justify-center text-center min-h-[170px]">
          ${qrSvg}
        </div>
      </div>
      <div class="border rounded-lg bg-white p-3 mb-3">
        <div class="text-xs uppercase tracking-wide text-slate-500 mb-1">Druckbereich</div>
        <div class="border rounded-lg p-3 text-center">
          <div class="text-2xl font-bold">${escapeHtml(toolTitle)}</div>
          <div class="text-sm mt-1">${escapeHtml(tool.label || "-")}</div>
          <div class="text-sm text-slate-700 mt-1">Ø ${escapeHtml(diameter)} · ${escapeHtml(cornerRadius)}</div>
          ${overhangLine}
          ${renderToolBorrowedDetailLine(tool, "text-xs text-orange-700")}
          ${insertRadiusLine}
          <div class="text-xs text-slate-600">Hersteller: ${escapeHtml(manufacturer)}</div>
          <div class="text-xs text-slate-600">Artikel: ${escapeHtml(articleNo)}</div>
          <div class="text-xs text-slate-600">Fach: ${escapeHtml(shelf)}</div>
          <div class="mt-2 flex justify-center">${qrSvg}</div>
          <div class="font-mono text-[10px] break-all mt-2">QR-Inhalt: ${escapeHtml(payload)}</div>
        </div>
      </div>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="printToolQrLabel('${tool.id}')">Drucken</button>
        <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="copyToolQrPayload('${tool.id}')">Text kopieren</button>
      </div>
    </div>
  </div>`;
}

function printToolQrLabel(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const payload = getToolQrPayload(tool);
  const qrSvg = renderSimpleQrSvg(payload);
  const diameter = tool.diameter || "-";
  const cornerRadius = tool.cornerRadius ? `R ${tool.cornerRadius}` : "-";
  const manufacturer = tool.manufacturer || "-";
  const articleNo = tool.articleNo || "-";
  const shelf = tool.shelf || "-";
  const overhangLine = renderToolOverhangDetailLine(tool, "meta");
  const insertRadiusLine =
    tool.insertTool && tool.insertRadius
      ? `<div class="meta">Plattenradius: R ${escapeHtml(tool.insertRadius)}</div>`
      : "";
  const printWindow = window.open("", "_blank", "width=420,height=560");
  if (!printWindow) {
    alert("Druckfenster konnte nicht geöffnet werden.");
    return;
  }

  printWindow.document.write(`<!doctype html>
  <html>
    <head>
      <title>QR-Etikett T ${escapeHtml(tool.tNumber)}</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 24px; color: #0f172a; }
        .label { border: 2px solid #0f172a; border-radius: 12px; padding: 16px; text-align: center; max-width: 340px; }
        .tnumber { font-size: 36px; font-weight: 800; margin: 0 0 3px; }
        .name { font-size: 16px; margin-bottom: 4px; }
        .meta { font-size: 12px; color: #475569; line-height: 1.25; }
        .qr-wrap { display: flex; justify-content: center; margin: 10px 0 8px; }
        .payload { font-family: monospace; font-size: 10px; word-break: break-all; }
        svg { width: 150px; height: 150px; }
      </style>
    </head>
    <body>
      <div class="label">
        <div class="tnumber">T ${escapeHtml(tool.tNumber)}</div>
        <div class="name">${escapeHtml(tool.label || "-")}</div>
        <div class="meta">Ø ${escapeHtml(diameter)} · ${escapeHtml(cornerRadius)}</div>
        ${overhangLine}
        ${renderToolBorrowedDetailLine(tool, "meta")}
        ${insertRadiusLine}
        <div class="meta">Hersteller: ${escapeHtml(manufacturer)}</div>
        <div class="meta">Artikel: ${escapeHtml(articleNo)}</div>
        <div class="meta">Fach: ${escapeHtml(shelf)}</div>
        <div class="qr-wrap">${qrSvg}</div>
        <div class="payload">QR-Inhalt: ${escapeHtml(payload)}</div>
      </div>
      <script>
        window.onload = function() {
          window.print();
        };
      </script>
    </body>
  </html>`);
  printWindow.document.close();
}

async function copyToolQrPayload(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const payload = getToolQrPayload(tool);
  if (!navigator.clipboard?.writeText) {
    alert(payload);
    return;
  }

  await navigator.clipboard.writeText(payload);
  alert("QR-Inhalt wurde kopiert.");
}

function renderToolScanner(mode = "all") {
  const withdrawCard = `<div class='border rounded-xl p-4 bg-slate-50 space-y-4'>
    <div>
      <h3 class='text-xl font-bold'>Werkzeug entnehmen</h3>
      <p class='text-sm text-slate-500 mt-1'>Werkzeugbestand um 1 reduzieren.</p>
    </div>
    <div class='grid gap-3'>
      <button class='px-4 py-4 rounded-lg bg-emerald-700 text-white text-lg font-semibold' onclick='openManualToolWithdraw()'>Manuell entnehmen</button>
      <button class='px-4 py-4 rounded-lg bg-slate-800 text-white text-lg font-semibold' onclick="openQrScannerPlaceholder('withdraw')">QR-Code scannen</button>
    </div>
  </div>`;
  const restockCard = `<div class='border rounded-xl p-4 bg-slate-50 space-y-4'>
    <div>
      <h3 class='text-xl font-bold'>Werkzeug einlagern</h3>
      <p class='text-sm text-slate-500 mt-1'>Werkzeugbestand um eine gewählte Menge erhöhen.</p>
    </div>
    <div class='grid gap-3'>
      <button class='px-4 py-4 rounded-lg bg-blue-700 text-white text-lg font-semibold' onclick='openManualToolRestock()'>Manuell einlagern</button>
      <button class='px-4 py-4 rounded-lg bg-slate-800 text-white text-lg font-semibold' onclick="openQrScannerPlaceholder('restock')">QR-Code scannen</button>
    </div>
  </div>`;
  const cards =
    mode === "withdraw"
      ? withdrawCard
      : mode === "restock"
        ? restockCard
        : `${withdrawCard}${restockCard}`;
  const gridClass = mode === "all" ? "grid md:grid-cols-2 gap-4" : "grid gap-4";
  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-2xl font-bold'>Werkzeug-Scanner</h2>
      <p class='text-sm text-slate-500 mt-1'>Für Werkzeugbereich / Handy-Terminal</p>
    </div>
    <div class='${gridClass}'>${cards}</div>
  </div>`;
}

function findToolByTNumberInput(value) {
  const normalized = normalizeToolImageNumber(value);
  if (!normalized) return null;

  return (
    state.tools.find((tool) => {
      return normalizeToolImageNumber(tool.tNumber) === normalized;
    }) || null
  );
}

function formatToolScannerSummary(tool) {
  if (!tool) return "";
  return `T ${tool.tNumber} · ${tool.label || "-"} · ${formatToolSize(tool)} · ${tool.holder || "-"} · Bestand ${tool.stock}`;
}

function updateScannerWithdrawPreview() {
  const input = document.getElementById("scannerWithdrawTNumber");
  const preview = document.getElementById("scannerWithdrawToolPreview");
  const qtyInput = document.getElementById("scannerWithdrawQty");
  if (!preview) return;

  const tool = findToolByTNumberInput(input?.value || "");
  const borrowedLine = tool?.isBorrowed
    ? `<div class="mt-1 text-orange-700">Dieses Werkzeug ist ausgeliehen an ${escapeHtml(tool.borrowedTo || "-")}</div>`
    : "";
  preview.innerHTML = tool
    ? `<div class="text-sm text-emerald-800 bg-emerald-50 border border-emerald-200 rounded p-2">${escapeHtml(formatToolScannerSummary(tool))}${borrowedLine}</div>`
    : `<div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>`;

  if (qtyInput && tool) {
    qtyInput.value = String(getDefaultToolBookingQty(tool));
  } else if (qtyInput && !qtyInput.value) {
    qtyInput.value = "1";
  }
}

function updateScannerRestockPreview() {
  const input = document.getElementById("scannerRestockTNumber");
  const preview = document.getElementById("scannerRestockToolPreview");
  const qtyInput = document.getElementById("scannerRestockQty");
  if (!preview) return;

  const tool = findToolByTNumberInput(input?.value || "");
  const borrowedLine = tool?.isBorrowed
    ? `<div class="mt-1 text-orange-700">Dieses Werkzeug ist ausgeliehen an ${escapeHtml(tool.borrowedTo || "-")}</div>`
    : "";
  preview.innerHTML = tool
    ? `<div class="text-sm text-emerald-800 bg-emerald-50 border border-emerald-200 rounded p-2">${escapeHtml(formatToolScannerSummary(tool))}${borrowedLine}</div>`
    : `<div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>`;

  if (qtyInput && tool?.ordered && Number(tool.orderedQty || 0) > 0) {
    qtyInput.value = String(tool.orderedQty);
  } else if (qtyInput && tool) {
    qtyInput.value = String(getDefaultToolBookingQty(tool));
  } else if (qtyInput && !qtyInput.value) {
    qtyInput.value = "1";
  }
}

function openQrScannerPlaceholder(mode) {
  openQrScanner(mode);
}

async function openQrScanner(mode) {
  stopQrScanner();

  const modeLabel = mode === "restock" ? "Einlagern" : "Entnehmen";
  const host = getModalHost();

  if (!("BarcodeDetector" in window)) {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">QR-Code scannen</h3>
        <p class="text-sm text-slate-700 mb-4">QR-Scanner wird von diesem Browser nicht unterstützt. Bitte manuell per T-Nummer buchen.</p>
        <div class="flex justify-end gap-2">
          <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
          <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="${mode === "restock" ? "openManualToolRestock()" : "openManualToolWithdraw()"}">Manuell buchen</button>
        </div>
      </div>
    </div>`;
    return;
  }

  if (!navigator.mediaDevices?.getUserMedia) {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">QR-Code scannen</h3>
        <p class="text-sm text-slate-700 mb-4">Kamera-Zugriff wird von diesem Browser nicht unterstützt. Bitte manuell per T-Nummer buchen.</p>
        <div class="flex justify-end gap-2">
          <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
          <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="${mode === "restock" ? "openManualToolRestock()" : "openManualToolWithdraw()"}">Manuell buchen</button>
        </div>
      </div>
    </div>`;
    return;
  }

  qrScannerMode = mode;
  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-lg p-4">
      <h3 class="text-lg font-bold mb-2">QR-Code scannen</h3>
      <p class="text-sm text-slate-500 mb-1">Modus: ${modeLabel}</p>
      <video id="qrScannerVideo" class="w-full bg-black rounded-lg mb-3" autoplay muted playsinline></video>
      <div id="qrScannerStatus" class="text-sm text-slate-600 bg-slate-50 border rounded p-2 mb-4">Kamera wird gestartet ...</div>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="stopQrScanner(); closeToolImagePopup()">Abbrechen</button>
        <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="stopQrScanner(); ${mode === "restock" ? "openManualToolRestock()" : "openManualToolWithdraw()"}">Manuell buchen</button>
      </div>
    </div>
  </div>`;

  try {
    qrScannerStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: "environment" },
    });
    const video = document.getElementById("qrScannerVideo");
    if (!video) return;
    video.srcObject = qrScannerStream;
    await video.play();
    const status = document.getElementById("qrScannerStatus");
    if (status) {
      status.textContent = "Kamera aktiv. QR-Code vor die Kamera halten.";
    }
    startQrScannerLoop();
  } catch (error) {
    console.error("Fehler beim Starten des QR-Scanners:", error);
    stopQrScanner();
    const status = document.getElementById("qrScannerStatus");
    if (status) {
      status.textContent =
        "Kamera konnte nicht gestartet werden. Bitte manuell per T-Nummer buchen.";
    }
  }
}

async function startQrScannerLoop() {
  const video = document.getElementById("qrScannerVideo");
  const status = document.getElementById("qrScannerStatus");
  if (!video || !("BarcodeDetector" in window)) return;

  const detector = new window.BarcodeDetector({ formats: ["qr_code"] });
  let scanning = false;

  qrScannerTimer = window.setInterval(async () => {
    if (scanning || !qrScannerStream) return;
    if (video.readyState < 2) return;

    scanning = true;
    try {
      const codes = await detector.detect(video);
      const first = codes?.[0];
      const rawValue = first?.rawValue || "";
      if (rawValue) {
        if (status) status.textContent = "QR-Code erkannt.";
        const mode = qrScannerMode;
        stopQrScanner();
        processScannedToolQr(rawValue, mode);
      }
    } catch (error) {
      console.error("Fehler beim Lesen des QR-Codes:", error);
      if (status) status.textContent = "QR-Code konnte nicht gelesen werden.";
    } finally {
      scanning = false;
    }
  }, 450);
}

async function processScannedToolQr(rawValue, mode) {
  const value = String(rawValue || "").trim();
  if (value.toUpperCase().startsWith("STORAGE:")) {
    openStorageLocationByQr(value);
    return;
  }

  if (!value.startsWith("TOOL:")) {
    alert("Ungültiger QR-Code");
    return;
  }

  const toolId = value.slice("TOOL:".length);
  const cachedTool = state.tools.find((t) => t.id === toolId);
  const tool = (await reloadSingleToolFromSupabase(toolId)) || cachedTool;
  if (!tool) {
    alert("Werkzeug nicht gefunden");
    return;
  }

  const isRestockMode = mode === "restock";
  const modeLabel = isRestockMode ? "Einlagerung" : "Entnahme";
  const orderedQty = Number(tool.orderedQty || tool.ordered_qty || 0);
  const defaultRestockQty =
    tool.ordered || orderedQty > 0 ? orderedQty : getDefaultToolBookingQty(tool);
  const defaultWithdrawQty = getDefaultToolBookingQty(tool);
  const orderedQtyLine =
    isRestockMode && (tool.ordered || orderedQty > 0)
      ? `<div><span class="text-slate-500">Bestellmenge:</span> ${escapeHtml(orderedQty)}</div>`
      : "";
  const restockQtyInput = isRestockMode
    ? `<label class="block text-sm font-medium mb-4">
        Menge
        <input id="qrRestockQty" type="number" min="1" class="border rounded p-2 w-full mt-1" value="${escapeHtml(defaultRestockQty)}" />
      </label>`
    : `<div class="space-y-3 mb-4">
        <label class="block text-sm font-medium">
          Menge
          <input id="qrWithdrawQty" type="number" min="1" class="border rounded p-2 w-full mt-1" value="${escapeHtml(defaultWithdrawQty)}" />
        </label>
        <label class="block text-sm font-medium">
          Kostenträger/Maschine bei Ausleihe
          <input id="qrBorrowedTo" class="border rounded p-2 w-full mt-1" placeholder="z. B. 350 / Abteilung Montage" />
        </label>
      </div>`;
  const actionButtons = isRestockMode
    ? `<button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="confirmQrToolRestock('${escapeHtml(String(tool.id || ""))}')">${tool.isBorrowed ? "Rückgabe einlagern" : "Einlagern bestätigen"}</button>`
    : `<button class="px-3 py-2 rounded bg-emerald-700 text-white" onclick="confirmQrToolWithdraw('${escapeHtml(String(tool.id || ""))}', 'normal')">Normal entnehmen</button>
       <button class="px-3 py-2 rounded bg-orange-700 text-white" onclick="confirmQrToolWithdraw('${escapeHtml(String(tool.id || ""))}', 'borrow')">Ausleihen</button>`;

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
      <h3 class="text-lg font-bold mb-2">Werkzeug erkannt</h3>
      <p class="text-sm text-slate-500 mb-3">Modus: ${modeLabel}</p>
      <div class="border rounded-lg bg-slate-50 p-3 space-y-1 text-sm mb-4">
        <div><span class="text-slate-500">T-Nummer:</span> T ${escapeHtml(tool.tNumber)}</div>
        <div><span class="text-slate-500">Bezeichnung:</span> ${escapeHtml(tool.label || "-")}</div>
        <div><span class="text-slate-500">Durchmesser:</span> ${escapeHtml(formatToolSize(tool))}</div>
        ${renderToolOverhangDetailLine(tool)}
        ${renderToolBorrowedDetailLine(tool)}
        <div><span class="text-slate-500">Aufnahme:</span> ${escapeHtml(tool.holder || "-")}</div>
        <div><span class="text-slate-500">Fach:</span> ${escapeHtml(tool.shelf || "-")}</div>
        <div><span class="text-slate-500">Bestand:</span> ${escapeHtml(tool.stock)}</div>
        ${orderedQtyLine}
      </div>
      ${restockQtyInput}
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Abbrechen</button>
        ${actionButtons}
      </div>
    </div>
  </div>`;
}

async function confirmQrToolWithdraw(toolId, mode = "normal") {
  const tool =
    (await reloadSingleToolFromSupabase(toolId)) ||
    state.tools.find((t) => t.id === toolId);
  if (!tool) {
    alert("Werkzeug nicht gefunden");
    return;
  }

  const oldStock = Number(tool.stock || 0);
  const qty = readPositiveQtyInput("qrWithdrawQty");
  if (!qty) {
    alert("Bitte eine gültige Entnahmemenge eingeben.");
    return;
  }
  if (oldStock < qty) {
    alert("Nicht genügend Bestand vorhanden.");
    return;
  }

  const newStock = oldStock - qty;
  const isBorrowing = mode === "borrow";
  const borrowedTo = isBorrowing
    ? String(document.getElementById("qrBorrowedTo")?.value || "").trim()
    : "";
  if (isBorrowing && !borrowedTo) {
    alert("Bitte Kostenträger/Maschine eintragen.");
    return;
  }

  const { error } = await supabaseClient
    .from("tools")
    .update({
      stock: newStock,
      ...(isBorrowing ? buildBorrowToolUpdateFields(borrowedTo) : {}),
    })
    .eq("id", tool.id);

  if (error) {
    console.error("Fehler bei QR-Scanner Entnahme:", error);
    alert("QR-Entnahme konnte nicht gespeichert werden.");
    return;
  }

  const tools = await loadToolsFromSupabase();
  if (!Array.isArray(tools) || tools.length === 0) {
    console.error("QR-Scanner Entnahme konnte nicht neu geladen werden", {
      toolId: tool.id,
    });
    alert("QR-Entnahme konnte nicht geprüft werden.");
    return;
  }
  const refreshedTool = tools.find((t) => t.id === tool.id);

  if (!refreshedTool || Number(refreshedTool.stock) !== newStock) {
    console.error("QR-Scanner Entnahme konnte nicht verifiziert werden", {
      toolId: tool.id,
      expectedStock: newStock,
      actualStock: refreshedTool?.stock,
    });
    alert("QR-Entnahme konnte nicht gespeichert werden.");
    return;
  }

  state.tools = tools;
  try {
    await addToolJournalEntry({
      toolId: refreshedTool.id,
      toolTNumber: refreshedTool.tNumber,
      toolLabel: refreshedTool.label,
      action: isBorrowing
        ? `Werkzeug ausgeliehen an ${borrowedTo}`
        : `QR-Scanner Entnahme ${qty}`,
      qty,
      stockBefore: oldStock,
      stockAfter: newStock,
      user: currentUser?.name || "Werkzeug-Scanner",
    });
    await refreshToolsAndRender();
    closeToolImagePopup();
  } catch (error) {
    console.error("QR Withdraw Journal Fehler:", error);
  }
}

async function confirmQrToolRestock(toolId) {
  const tool =
    (await reloadSingleToolFromSupabase(toolId)) ||
    state.tools.find((t) => t.id === toolId);
  if (!tool) {
    alert("Werkzeug nicht gefunden");
    return;
  }

  const qty = Number(document.getElementById("qrRestockQty")?.value || 0);
  if (!Number.isFinite(qty) || qty <= 0) {
    alert("Bitte eine gültige Menge eingeben.");
    return;
  }

  const oldStock = Number(tool.stock || 0);
  const newStock = oldStock + qty;
  const returnFrom = tool.isBorrowed ? tool.borrowedTo || "-" : "";
  const { error } = await supabaseClient
    .from("tools")
    .update({
      stock: newStock,
      ordered: false,
      ordered_qty: 0,
      ...buildReturnToolUpdateFields(tool),
    })
    .eq("id", tool.id);

  if (error) {
    console.error("Fehler bei QR-Scanner Einlagerung:", error);
    alert("QR-Einlagerung konnte nicht gespeichert werden.");
    return;
  }

  const loadedTools = await loadToolsFromSupabase();
  if (Array.isArray(loadedTools) && loadedTools.length > 0) {
    state.tools = loadedTools;
  }

  let refreshedTool = state.tools.find((t) => t.id === tool.id);

  if (!refreshedTool) {
    console.warn("Werkzeug wurde nach QR-Einlagerung nicht neu geladen.", {
      toolId: tool.id,
    });
    state.tools = state.tools.map((t) =>
      t.id === tool.id
        ? {
            ...t,
            stock: newStock,
            ordered: false,
            orderedQty: 0,
            ordered_qty: 0,
            isBorrowed: false,
            borrowedTo: "",
            borrowedAt: "",
          }
        : t,
    );
    refreshedTool = state.tools.find((t) => t.id === tool.id) || {
      ...tool,
      stock: newStock,
      ordered: false,
      orderedQty: 0,
      ordered_qty: 0,
      isBorrowed: false,
      borrowedTo: "",
      borrowedAt: "",
    };
  }

  try {
    await addToolJournalEntry({
      toolId: refreshedTool.id,
      toolTNumber: refreshedTool.tNumber,
      toolLabel: refreshedTool.label,
      action: tool.isBorrowed
        ? `Werkzeug zurück von ${returnFrom}`
        : `QR-Scanner Einlagerung ${qty}`,
      qty,
      stockBefore: oldStock,
      stockAfter: newStock,
      user: currentUser?.name || "Werkzeug-Scanner",
    });
    await refreshToolsAndRender();
    closeToolImagePopup();
  } catch (error) {
    console.error("QR Restock Journal Fehler:", error);
  }
}

function stopQrScanner() {
  if (qrScannerTimer) {
    window.clearInterval(qrScannerTimer);
    qrScannerTimer = null;
  }

  if (qrScannerStream) {
    qrScannerStream.getTracks().forEach((track) => track.stop());
    qrScannerStream = null;
  }

  qrScannerMode = null;
}

function openManualToolWithdraw(initialTNumber = "") {
  const host = getModalHost();
  stopQrScanner();

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-xl p-4">
      <h3 class="text-lg font-bold mb-3">Werkzeug manuell entnehmen</h3>
      <input id="scannerWithdrawTNumber" inputmode="numeric" class="border rounded p-2 w-full mb-2" placeholder="T-Nummer eingeben, z. B. 133" value="${escapeHtml(initialTNumber)}" oninput="updateScannerWithdrawPreview()" />
      <div id="scannerWithdrawToolPreview" class="mb-4">
        <div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>
      </div>
      <label class="block text-sm font-medium mb-4">
        Menge
        <input id="scannerWithdrawQty" type="number" min="1" class="border rounded p-2 w-full mt-1" value="1" />
      </label>
      <label class="block text-sm font-medium mb-4">
        Kostenträger/Maschine bei Ausleihe
        <input id="scannerWithdrawBorrowedTo" class="border rounded p-2 w-full mt-1" placeholder="z. B. 350 / Abteilung Montage" />
      </label>
      <div class="flex justify-end gap-2 flex-wrap">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Abbrechen</button>
        <button id="scannerWithdrawSave" class="px-3 py-2 rounded bg-emerald-700 text-white">Normal entnehmen</button>
        <button id="scannerBorrowSave" class="px-3 py-2 rounded bg-orange-700 text-white">Ausleihen</button>
      </div>
    </div>
  </div>`;

  const saveWithdraw = async (mode = "normal") => {
    const input = host.querySelector("#scannerWithdrawTNumber");
    const tool = findToolByTNumberInput(input?.value || "");
    if (!tool) {
      alert("Werkzeug wurde nicht gefunden.");
      return;
    }
    const freshTool = (await reloadSingleToolFromSupabase(tool.id)) || tool;
    const qty = Number(host.querySelector("#scannerWithdrawQty")?.value || 0);
    if (!Number.isFinite(qty) || qty <= 0) {
      alert("Bitte eine gültige Entnahmemenge eingeben.");
      return;
    }
    if (Number(freshTool.stock || 0) < qty) {
      alert("Nicht genügend Bestand vorhanden.");
      return;
    }

    const oldStock = Number(freshTool.stock || 0);
    const newStock = oldStock - qty;
    const isBorrowing = mode === "borrow";
    const borrowedTo = isBorrowing
      ? String(host.querySelector("#scannerWithdrawBorrowedTo")?.value || "").trim()
      : "";
    if (isBorrowing && !borrowedTo) {
      alert("Bitte Kostenträger/Maschine eintragen.");
      return;
    }
    if (newStock < 0) {
      alert("Nicht genügend Bestand vorhanden.");
      return;
    }

    const payload = {
      stock: newStock,
      ...(isBorrowing ? buildBorrowToolUpdateFields(borrowedTo) : {}),
    };
    const { error } = await supabaseClient
      .from("tools")
      .update(payload)
      .eq("id", freshTool.id);

    if (error) {
      console.error("Fehler bei Werkzeug-Scanner Entnahme:", error);
      alert(`Entnahme konnte nicht gespeichert werden: ${error.message}`);
      return;
    }

    const { data: refreshed, error: reloadError } = await supabaseClient
      .from("tools")
      .select("*")
      .eq("id", freshTool.id)
      .maybeSingle();

    if (reloadError || !refreshed) {
      console.error("Scanner stock reload failed", {
        toolId: freshTool.id,
        reloadError,
      });
      alert(
        "Entnahme konnte nicht geprüft werden. Werkzeug wurde nicht neu geladen.",
      );
      return;
    }

    if (Number(refreshed.stock) !== newStock) {
      console.error("Scanner stock update verification failed", {
        toolId: freshTool.id,
        expectedStock: newStock,
        actualStock: refreshed.stock,
      });
      alert(
        "Entnahme wurde nicht gespeichert. Bitte RLS/Update-Rechte für tools prüfen.",
      );
      return;
    }

    const refreshedTool = normalizeToolFromDb(refreshed);
    const index = state.tools.findIndex((t) => t.id === refreshedTool.id);
    if (index !== -1) {
      state.tools[index] = refreshedTool;
    }

    console.log("Journal wird geschrieben aus openManualToolWithdraw", {
      tool: refreshedTool,
      oldStock,
      newStock,
      qty,
    });
    await addToolJournalEntry({
      toolId: refreshedTool.id,
      toolTNumber: refreshedTool.tNumber,
      toolLabel: refreshedTool.label,
      action: isBorrowing
        ? `Werkzeug ausgeliehen an ${borrowedTo}`
        : `Werkzeug-Scanner Entnahme ${qty}`,
      qty,
      stockBefore: oldStock,
      stockAfter: newStock,
      user: currentUser?.name || "Werkzeug-Scanner",
    });

    closeToolImagePopup();
    await refreshToolsAndRender();
  };

  host.querySelector("#scannerWithdrawSave")?.addEventListener("click", () => {
    saveWithdraw("normal");
  });
  host.querySelector("#scannerBorrowSave")?.addEventListener("click", () => {
    saveWithdraw("borrow");
  });

  host.querySelector("#scannerWithdrawTNumber")?.focus();
  updateScannerWithdrawPreview();
}

function openManualToolRestock(initialTNumber = "") {
  const host = getModalHost();
  stopQrScanner();

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-xl p-4">
      <h3 class="text-lg font-bold mb-3">Werkzeug manuell einlagern</h3>
      <input id="scannerRestockTNumber" inputmode="numeric" class="border rounded p-2 w-full mb-2" placeholder="T-Nummer eingeben, z. B. 133" value="${escapeHtml(initialTNumber)}" oninput="updateScannerRestockPreview()" />
      <div id="scannerRestockToolPreview" class="mb-3">
        <div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>
      </div>
      <label class="block text-sm font-medium mb-4">
        Menge
        <input id="scannerRestockQty" type="number" min="1" class="border rounded p-2 w-full mt-1" value="1" />
      </label>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Abbrechen</button>
        <button id="scannerRestockSave" class="px-3 py-2 rounded bg-blue-700 text-white">Einlagern</button>
      </div>
    </div>
  </div>`;

  host.querySelector("#scannerRestockSave")?.addEventListener("click", async () => {
    const input = host.querySelector("#scannerRestockTNumber");
    const qty = Number(host.querySelector("#scannerRestockQty")?.value || 0);
    const tool = findToolByTNumberInput(input?.value || "");
    if (!tool) {
      alert("Werkzeug wurde nicht gefunden.");
      return;
    }
    const freshTool = (await reloadSingleToolFromSupabase(tool.id)) || tool;
    if (!Number.isFinite(qty) || qty <= 0) {
      alert("Bitte eine gültige Menge eingeben.");
      return;
    }

    const oldStock = Number(freshTool.stock || 0);
    const newStock = oldStock + qty;
    const returnFrom = freshTool.isBorrowed ? freshTool.borrowedTo || "-" : "";
    const payload = {
      stock: newStock,
      ...buildReturnToolUpdateFields(freshTool),
    };
    const { error } = await supabaseClient
      .from("tools")
      .update(payload)
      .eq("id", freshTool.id);

    if (error) {
      console.error("Fehler bei Werkzeug-Scanner Einlagerung:", error);
      alert(`Einlagerung konnte nicht gespeichert werden: ${error.message}`);
      return;
    }

    const { data: refreshed, error: reloadError } = await supabaseClient
      .from("tools")
      .select("*")
      .eq("id", freshTool.id)
      .maybeSingle();

    if (reloadError || !refreshed) {
      console.error("Scanner stock reload failed", {
        toolId: freshTool.id,
        reloadError,
      });
      alert(
        "Einlagerung konnte nicht geprüft werden. Werkzeug wurde nicht neu geladen.",
      );
      return;
    }

    if (Number(refreshed.stock) !== newStock) {
      console.error("Scanner stock update verification failed", {
        toolId: freshTool.id,
        expectedStock: newStock,
        actualStock: refreshed.stock,
      });
      alert(
        "Einlagerung wurde nicht gespeichert. Bitte RLS/Update-Rechte für tools prüfen.",
      );
      return;
    }

    const refreshedTool = normalizeToolFromDb(refreshed);
    const index = state.tools.findIndex((t) => t.id === refreshedTool.id);
    if (index !== -1) {
      state.tools[index] = refreshedTool;
    }

    console.log("Journal wird geschrieben aus openManualToolRestock", {
      tool: refreshedTool,
      oldStock,
      newStock,
      qty,
    });
    await addToolJournalEntry({
      toolId: refreshedTool.id,
      toolTNumber: refreshedTool.tNumber,
      toolLabel: refreshedTool.label,
      action: freshTool.isBorrowed
        ? `Werkzeug zurück von ${returnFrom}`
        : `Werkzeug-Scanner Einlagerung ${qty}`,
      qty,
      stockBefore: oldStock,
      stockAfter: newStock,
      user: currentUser?.name || "Werkzeug-Scanner",
    });

    closeToolImagePopup();
    await refreshToolsAndRender();
  });

  host.querySelector("#scannerRestockTNumber")?.focus();
  updateScannerRestockPreview();
}

function askYesNoCentered(message) {
  const host = getModalHost();
  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">Bestätigung</h3>
        <p class="text-sm text-slate-700 mb-4">${message}</p>
        <div class="flex justify-end gap-2">
          <button id="modalNo" class="px-3 py-1 rounded bg-slate-200">Nein</button>
          <button id="modalYes" class="px-3 py-1 rounded bg-emerald-700 text-white">Ja</button>
        </div>
      </div>
    </div>`;
    host.querySelector("#modalYes")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(true);
    });
    host.querySelector("#modalNo")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(false);
    });
  });
}

function askNumberCentered(message, initialValue = "1") {
  const host = getModalHost();
  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">Eingabe</h3>
        <p class="text-sm text-slate-700 mb-3">${message}</p>
        <input id="modalNumber" type="number" min="1" class="border rounded p-2 w-full mb-4" value="${initialValue}" />
        <div class="flex justify-end gap-2">
          <button id="modalNo" class="px-3 py-1 rounded bg-slate-200">Abbrechen</button>
          <button id="modalOk" class="px-3 py-1 rounded bg-slate-900 text-white">Bestätigen</button>
        </div>
      </div>
    </div>`;
    const input = host.querySelector("#modalNumber");
    input?.focus();
    host.querySelector("#modalOk")?.addEventListener("click", () => {
      const val = Number(input?.value || 0);
      host.innerHTML = "";
      resolve(Number.isFinite(val) ? val : null);
    });
    host.querySelector("#modalNo")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(null);
    });
  });
}

function askTextCentered(message, placeholder = "") {
  const host = getModalHost();
  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">Eingabe</h3>
        <p class="text-sm text-slate-700 mb-3">${escapeHtml(message)}</p>
        <input id="modalText" class="border rounded p-2 w-full mb-4" placeholder="${escapeHtml(placeholder)}" />
        <div class="flex justify-end gap-2">
          <button id="modalNo" class="px-3 py-1 rounded bg-slate-200">Abbrechen</button>
          <button id="modalOk" class="px-3 py-1 rounded bg-slate-900 text-white">Bestätigen</button>
        </div>
      </div>
    </div>`;
    const input = host.querySelector("#modalText");
    input?.focus();
    host.querySelector("#modalOk")?.addEventListener("click", () => {
      const value = String(input?.value || "").trim();
      host.innerHTML = "";
      resolve(value);
    });
    host.querySelector("#modalNo")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(null);
    });
  });
}

async function editTool(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;

  const data = await editToolCentered(tool);
  if (!data) return;

  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  if (
    !data.label ||
    !data.diameter ||
    !isValidToolShelf(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return alert(
      "Bitte Felder korrekt ausfüllen (2-stellige Zahl + Buchstabe für Fach, z. B. 26C oder 02T; Aufnahme HSK 100 oder HSK 63).",
    );
  }

  if (isThreadTool && !data.threadPrefix) {
    return alert("Bitte Gewindekennung wählen.");
  }

  if (isThreadTool && data.threadPrefix === "MF" && !data.threadPitch) {
    return alert("Bitte bei MF die Steigung (P) angeben.");
  }

  if (isRadiusTool && !data.cornerRadius) {
    return alert("Bitte Schneidenradius eingeben.");
  }

  if (!isThreadTool && !Number.isFinite(Number(data.diameter))) {
    return alert("Bitte gültigen numerischen Durchmesser eingeben.");
  }

  if (
    data.insertTool &&
    (!Number.isFinite(Number(data.insertEdges)) ||
      Number(data.insertEdges) <= 0)
  ) {
    return alert("Bitte Anzahl der Schneiden > 0 eingeben.");
  }

  const payload = {
    label: data.label,
    diameter: String(data.diameter),
    overhang_length: data.overhangLength || null,
    thread_prefix: isThreadTool ? data.threadPrefix || null : null,
    thread_pitch:
      isThreadTool && data.threadPrefix === "MF"
        ? data.threadPitch || null
        : null,
    corner_radius: isRadiusTool ? data.cornerRadius || null : null,
    material_id: data.materialId || null,
    shelf: data.shelf,
    article_no: data.articleNo,
    holder: data.holder,
    stock: Number(data.stock || 0),
    min_stock: Number(data.minStock || 0),
    optimal_stock: Number(data.optimalStock || 0),
    manufacturer: data.manufacturer || null,
    insert_tool: !!data.insertTool,
    insert_edges: data.insertTool ? Number(data.insertEdges || 0) : 0,
    insert_radius: data.insertTool ? data.insertRadius || null : null,
  };

  const { error } = await supabaseClient
    .from("tools")
    .update(payload)
    .eq("id", toolId);

  if (error) {
    console.error("Fehler beim Bearbeiten des Werkzeugs:", error);
    return alert(`Werkzeug konnte nicht gespeichert werden: ${error.message}`);
  }

  state.tools = state.tools.map((existingTool) =>
    existingTool.id === toolId
      ? {
          ...existingTool,
          label: data.label,
          diameter: String(data.diameter),
          overhangLength: data.overhangLength || "",
          threadPrefix: isThreadTool ? data.threadPrefix || "" : "",
          threadPitch:
            isThreadTool && data.threadPrefix === "MF"
              ? data.threadPitch || ""
              : "",
          cornerRadius: isRadiusTool ? data.cornerRadius || "" : "",
          materialId: data.materialId || null,
          shelf: data.shelf,
          articleNo: data.articleNo,
          holder: data.holder,
          stock: Number(data.stock || 0),
          minStock: Number(data.minStock || 0),
          optimalStock: Number(data.optimalStock || 0),
          manufacturer: data.manufacturer || "",
          insertTool: !!data.insertTool,
          insertEdges: data.insertTool ? Number(data.insertEdges || 0) : 0,
          insertRadius: data.insertTool ? data.insertRadius || "" : "",
        }
      : existingTool,
  );

  await refreshToolsAndRender();
}

async function editToolCentered(tool) {
  const host = getModalHost();

  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-4xl max-h-[90vh] overflow-auto p-5">
        <div class='flex items-center justify-between mb-4 gap-3'>
          <h3 class="text-lg font-bold">Werkzeug bearbeiten – T ${tool.tNumber}</h3>
          <button id="toolEditCloseTop" class="px-3 py-1 rounded bg-slate-200">Schließen</button>
        </div>

        <div class='grid md:grid-cols-2 gap-4'>
          <label class='text-sm font-medium'>
            T-Nummer
            <input id='editTNumber' class='border rounded p-2 w-full mt-1 bg-slate-100' value="${tool.tNumber}" disabled />
          </label>

          <label class='text-sm font-medium'>
            Bezeichnung
            <select id='editLabel' class='border rounded p-2 w-full mt-1' onchange='updateEditToolTypeFields()'>
              ${getToolLabels()
                .map(
                  (l) =>
                    `<option value="${l}" ${tool.label === l ? "selected" : ""}>${l}</option>`,
                )
                .join("")}
            </select>
          </label>

          <label class='text-sm font-medium'>
            Durchmesser
            <input id='editDiameter' class='border rounded p-2 w-full mt-1' placeholder='Durchmesser' value="${tool.diameter ?? ""}" />
          </label>

          <label class='text-sm font-medium'>
            Ausspannlänge (AL) [mm]
            <input id='editOverhangLength' type='number' step='0.1' min='0' class='border rounded p-2 w-full mt-1' placeholder='AL' value="${getToolOverhangLength(tool)}" />
          </label>

          <label class='text-sm font-medium'>
            Schneidwerkstoff
            <select id='editMaterial' class='border rounded p-2 w-full mt-1'>
              <option value=''>Schneidwerkstoff wählen</option>
              ${toolMaterials
                .map(
                  (m) =>
                    `<option value="${m.id}" ${tool.materialId === m.id ? "selected" : ""}>${m.name}</option>`,
                )
                .join("")}
            </select>
          </label>

          <div id='editThreadPrefixWrap' style='display:none;'>
            <label class='text-sm font-medium block'>
              Gewindekennung
              <select id='editThreadPrefix' class='border rounded p-2 w-full mt-1' onchange='updateEditThreadPitchVisibility()'>
                <option value=''>Kennung wählen</option>
                <option value='M' ${tool.threadPrefix === "M" ? "selected" : ""}>M</option>
                <option value='MF' ${tool.threadPrefix === "MF" ? "selected" : ""}>MF</option>
                <option value='G' ${tool.threadPrefix === "G" ? "selected" : ""}>G</option>
                <option value='UNF' ${tool.threadPrefix === "UNF" ? "selected" : ""}>UNF</option>
                <option value='UNC' ${tool.threadPrefix === "UNC" ? "selected" : ""}>UNC</option>
                <option value='Mx' ${tool.threadPrefix === "Mx" ? "selected" : ""}>Mx</option>
              </select>
            </label>
          </div>

          <div id='editThreadPitchWrap' style='display:none;'>
            <label class='text-sm font-medium block'>
              Steigung
              <input id='editThreadPitch' class='border rounded p-2 w-full mt-1' placeholder='Steigung P (nur MF)' value="${tool.threadPitch || ""}" />
            </label>
          </div>

          <div id='editCornerRadiusWrap' style='display:none;'>
            <label class='text-sm font-medium block'>
              Schneidenradius
              <input id='editCornerRadius' class='border rounded p-2 w-full mt-1' placeholder='Schneidenradius' value="${tool.cornerRadius || ""}" />
            </label>
          </div>

          <label class='text-sm font-medium'>
            Lagerfach
            <input id='editShelf' class='border rounded p-2 w-full mt-1' placeholder='01S' value="${tool.shelf || ""}" />
          </label>

          <label class='text-sm font-medium'>
            Artikelnummer
            <input id='editArticleNo' class='border rounded p-2 w-full mt-1' placeholder='Artikel Nr.' value="${tool.articleNo || ""}" />
          </label>

          <label class='text-sm font-medium'>
            Aufnahme
            <select id='editHolder' class='border rounded p-2 w-full mt-1'>
              <option value='HSK 100' ${tool.holder === "HSK 100" ? "selected" : ""}>HSK 100</option>
              <option value='HSK 63' ${tool.holder === "HSK 63" ? "selected" : ""}>HSK 63</option>
            </select>
          </label>

          <label class='text-sm font-medium'>
            Hersteller
            <select id='editManufacturer' class='border rounded p-2 w-full mt-1'>
              ${getToolManufacturers()
                .map(
                  (m) =>
                    `<option value="${m}" ${tool.manufacturer === m ? "selected" : ""}>${m}</option>`,
                )
                .join("")}
            </select>
          </label>

          <label class='text-sm font-medium'>
            Bestand
            <input id='editStock' type='number' class='border rounded p-2 w-full mt-1' placeholder='Bestand' value="${tool.stock ?? 0}" />
          </label>

          <label class='text-sm font-medium'>
            Mindestbestand
            <input id='editMinStock' type='number' min='0' class='border rounded p-2 w-full mt-1' placeholder='Mindestbestand' value="${tool.minStock ?? 0}" />
          </label>

          <label class='text-sm font-medium'>
            Optimale Stückzahl
            <input id='editOptimalStock' type='number' class='border rounded p-2 w-full mt-1' placeholder='Optimale Stückzahl' value="${tool.optimalStock ?? 0}" />
          </label>

          <div></div>

          <label class='flex items-center gap-2 text-sm md:col-span-2'>
            <input id='editInsertTool' type='checkbox' ${tool.insertTool ? "checked" : ""} onchange='toggleInsertToolFieldsById("editInsertTool","editInsertEdges","editInsertRadius")' />
            Wendeplattenwerkzeug
          </label>

          <label class='text-sm font-medium md:col-span-2'>
            Anzahl Schneiden
            <input id='editInsertEdges' type='number' class='border rounded p-2 w-full mt-1' placeholder='Anzahl Schneiden' value="${tool.insertEdges ?? 0}" ${tool.insertTool ? "" : "disabled"} />
          </label>

          <label id='editInsertRadiusWrap' class='text-sm font-medium md:col-span-2' style='display:${tool.insertTool ? "" : "none"};'>
            Plattenradius optional
            <input id='editInsertRadius' class='border rounded p-2 w-full mt-1' placeholder='Plattenradius optional, z. B. 0.8' value="${escapeHtml(tool.insertRadius || "")}" ${tool.insertTool ? "" : "disabled"} />
          </label>
        </div>

        <div class='flex justify-end gap-2 mt-5'>
          <button id="toolEditCloseBottom" class="px-3 py-2 rounded bg-slate-200">Abbrechen</button>
          <button id="toolEditSave" class="px-3 py-2 rounded bg-slate-900 text-white">Speichern</button>
        </div>
      </div>
    </div>`;

    updateEditToolTypeFields();
    toggleInsertToolFieldsById(
      "editInsertTool",
      "editInsertEdges",
      "editInsertRadius",
    );

    const close = () => {
      host.innerHTML = "";
    };

    host.querySelector("#toolEditCloseTop")?.addEventListener("click", () => {
      close();
      resolve(null);
    });

    host
      .querySelector("#toolEditCloseBottom")
      ?.addEventListener("click", () => {
        close();
        resolve(null);
      });

    host.querySelector("#toolEditSave")?.addEventListener("click", () => {
      const data = {
        label: document.getElementById("editLabel")?.value || "",
        diameter: document.getElementById("editDiameter")?.value?.trim() || "",
        overhangLength:
          document.getElementById("editOverhangLength")?.value?.trim() || "",
        threadPrefix: document.getElementById("editThreadPrefix")?.value || "",
        threadPitch:
          document.getElementById("editThreadPitch")?.value?.trim() || "",
        cornerRadius:
          document.getElementById("editCornerRadius")?.value?.trim() || "",
        materialId: document.getElementById("editMaterial")?.value || "",
        shelf:
          document.getElementById("editShelf")?.value?.trim().toUpperCase() ||
          "",
        articleNo:
          document.getElementById("editArticleNo")?.value?.trim() || "",
        holder: document.getElementById("editHolder")?.value || "",
        manufacturer: document.getElementById("editManufacturer")?.value || "",
        stock: Number(document.getElementById("editStock")?.value || 0),
        minStock: Number(document.getElementById("editMinStock")?.value || 0),
        optimalStock: Number(
          document.getElementById("editOptimalStock")?.value || 0,
        ),
        insertTool: !!document.getElementById("editInsertTool")?.checked,
        insertEdges: Number(
          document.getElementById("editInsertEdges")?.value || 0,
        ),
        insertRadius:
          document.getElementById("editInsertRadius")?.value?.trim() || "",
      };

      close();
      resolve(data);
    });
  });
}

function updateEditToolTypeFields() {
  const label = document.getElementById("editLabel")?.value || "";
  const isThread = isThreadToolLabel(label);
  const isRadius = isRadiusToolLabel(label);

  const threadPrefixWrap = document.getElementById("editThreadPrefixWrap");
  const threadPitchWrap = document.getElementById("editThreadPitchWrap");
  const cornerRadiusWrap = document.getElementById("editCornerRadiusWrap");

  if (threadPrefixWrap) threadPrefixWrap.style.display = isThread ? "" : "none";
  if (cornerRadiusWrap) cornerRadiusWrap.style.display = isRadius ? "" : "none";

  updateEditThreadPitchVisibility();
}

function updateEditThreadPitchVisibility() {
  const label = document.getElementById("editLabel")?.value || "";
  const threadPrefix = document.getElementById("editThreadPrefix")?.value || "";
  const threadPitchWrap = document.getElementById("editThreadPitchWrap");

  const visible = isThreadToolLabel(label) && threadPrefix === "MF";
  if (threadPitchWrap) threadPitchWrap.style.display = visible ? "" : "none";
}

async function editTool(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;

  const data = await editToolCentered(tool);
  if (!data) return;

  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  if (
    !data.label ||
    !data.diameter ||
    !isValidToolShelf(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return alert(
      "Bitte Felder korrekt ausfüllen (2-stellige Zahl + Buchstabe für Fach, z. B. 26C oder 02T; Aufnahme HSK 100 oder HSK 63).",
    );
  }

  if (isThreadTool && !data.threadPrefix) {
    return alert("Bitte Gewindekennung wählen.");
  }

  if (isThreadTool && data.threadPrefix === "MF" && !data.threadPitch) {
    return alert("Bitte bei MF die Steigung (P) angeben.");
  }

  if (isRadiusTool && !data.cornerRadius) {
    return alert("Bitte Schneidenradius eingeben.");
  }

  if (!isThreadTool && !Number.isFinite(Number(data.diameter))) {
    return alert("Bitte gültigen numerischen Durchmesser eingeben.");
  }

  if (
    data.insertTool &&
    (!Number.isFinite(Number(data.insertEdges)) ||
      Number(data.insertEdges) <= 0)
  ) {
    return alert("Bitte Anzahl der Schneiden > 0 eingeben.");
  }

  const payload = {
    label: data.label,
    diameter: String(data.diameter),
    overhang_length: data.overhangLength || null,
    thread_prefix: isThreadTool ? data.threadPrefix || null : null,
    thread_pitch:
      isThreadTool && data.threadPrefix === "MF"
        ? data.threadPitch || null
        : null,
    corner_radius: isRadiusTool ? data.cornerRadius || null : null,
    material_id: data.materialId || null,
    shelf: data.shelf,
    article_no: data.articleNo,
    holder: data.holder,
    stock: Number(data.stock || 0),
    min_stock: Number(data.minStock || 0),
    optimal_stock: Number(data.optimalStock || 0),
    manufacturer: data.manufacturer || null,
    insert_tool: !!data.insertTool,
    insert_edges: data.insertTool ? Number(data.insertEdges || 0) : 0,
    insert_radius: data.insertTool ? data.insertRadius || null : null,
  };

  const { error } = await supabaseClient
    .from("tools")
    .update(payload)
    .eq("id", toolId);

  if (error) {
    console.error("Fehler beim Bearbeiten des Werkzeugs:", error);
    return alert(`Werkzeug konnte nicht gespeichert werden: ${error.message}`);
  }

  state.tools = state.tools.map((existingTool) =>
    existingTool.id === toolId
      ? {
          ...existingTool,
          label: data.label,
          diameter: String(data.diameter),
          overhangLength: data.overhangLength || "",
          threadPrefix: isThreadTool ? data.threadPrefix || "" : "",
          threadPitch:
            isThreadTool && data.threadPrefix === "MF"
              ? data.threadPitch || ""
              : "",
          cornerRadius: isRadiusTool ? data.cornerRadius || "" : "",
          materialId: data.materialId || null,
          shelf: data.shelf,
          articleNo: data.articleNo,
          holder: data.holder,
          stock: Number(data.stock || 0),
          minStock: Number(data.minStock || 0),
          optimalStock: Number(data.optimalStock || 0),
          manufacturer: data.manufacturer || "",
          insertTool: !!data.insertTool,
          insertEdges: data.insertTool ? Number(data.insertEdges || 0) : 0,
          insertRadius: data.insertTool ? data.insertRadius || "" : "",
        }
      : existingTool,
  );

  await refreshToolsAndRender();
}

function openCreateToolModal() {
  if (currentUser?.role !== "admin") return;
  const host = getModalHost();
  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-3xl max-h-[90vh] overflow-auto p-4">
      <div class='flex items-center justify-between mb-4 gap-3'>
        <h3 class="text-lg font-bold">Neues Werkzeug anlegen</h3>
        <button id="toolCreateCloseTop" class="px-3 py-1 rounded bg-slate-200">Schließen</button>
      </div>
      <div class='space-y-4'>
        <div class='border rounded-lg p-3 bg-slate-50'>
          ${renderToolCreateForm("tool")}
        </div>
        <div class='flex justify-between gap-2 flex-wrap'>
          <div class='flex gap-2 flex-wrap'>
            <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='addToolLabel()'>Bezeichnung hinzufügen</button>
            <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='addToolManufacturer()'>Hersteller hinzufügen</button>
          </div>
          <div class='flex gap-2'>
            <button id="toolCreateCloseBottom" class="px-3 py-2 rounded bg-slate-200">Abbrechen</button>
            <button id="toolCreateSave" class="px-3 py-2 rounded bg-slate-900 text-white">Werkzeug speichern</button>
          </div>
        </div>
      </div>
    </div>
  </div>`;

  updateToolTypeFields("tool");

  const close = () => {
    host.innerHTML = "";
  };

  host.querySelector("#toolCreateCloseTop")?.addEventListener("click", close);
  host
    .querySelector("#toolCreateCloseBottom")
    ?.addEventListener("click", close);

  host.querySelector("#toolCreateSave")?.addEventListener("click", async () => {
    if (currentUser.role !== "admin")
      return alert("Nur Admin darf Werkzeuge anlegen.");

    const data = collectToolFormData(document);
    const validation = validateToolData(data);
    if (!validation.ok) return alert(validation.message);

    const normalized = normalizeToolData(data);

    const payload = {
      t_number: String(normalized.tNumber),
      label: normalized.label,
      diameter: String(normalized.diameter),
      overhang_length: normalized.overhangLength || null,
      thread_prefix: normalized.threadPrefix || null,
      thread_pitch: normalized.threadPitch || null,
      corner_radius: normalized.cornerRadius || null,
      material_id: normalized.materialId || null,
      shelf: normalized.shelf,
      article_no: normalized.articleNo,
      holder: normalized.holder,
      stock: Number(normalized.stock || 0),
      min_stock: Number(normalized.minStock || 0),
      optimal_stock: Number(normalized.optimalStock || 0),
      manufacturer: normalized.manufacturer || null,
      ordered: !!normalized.ordered,
      ordered_qty: Number(normalized.orderedQty || 0),
      insert_tool: !!normalized.insertTool,
      insert_edges: Number(normalized.insertEdges || 0),
      insert_radius: normalized.insertRadius || null,
      is_borrowed: false,
      borrowed_to: null,
      borrowed_at: null,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    };

    const { data: inserted, error } = await supabaseClient
      .from("tools")
      .insert(payload)
      .select()
      .single();

    if (error) {
      console.error("Fehler beim Anlegen des Werkzeugs in Supabase:", error);
      return alert(
        `Werkzeug konnte nicht gespeichert werden: ${error.message}`,
      );
    }

    state.tools.push(normalizeToolFromDb(inserted));
    persist();
    close();
    render();
  });
}

function suggestedOrderQty(tool) {
  const stock = Number(tool.stock || 0);
  const optimal = Number(tool.optimalStock || 0);
  const minStock = Number(tool.minStock || 0);

  const baseFromOptimal = Math.max(0, optimal - stock);
  const statSuggestion = getOptimalQtySuggestion(tool);

  if (Number.isFinite(statSuggestion) && statSuggestion > 0) {
    return Math.max(baseFromOptimal, statSuggestion);
  }

  if (baseFromOptimal > 0) return baseFromOptimal;

  if (stock <= minStock) {
    return Math.max(1, minStock - stock + 1);
  }

  return 0;
}

function effectiveOrderQty(tool) {
  const override = Number(state.toolOrderOverrides?.[tool.id]);
  if (Number.isFinite(override) && override >= 0) return override;
  return suggestedOrderQty(tool);
}

function setToolOrderOverride(toolId, value) {
  if (currentUser.role !== "admin") return;
  const qty = Math.max(0, Number(value || 0));
  if (!state.toolOrderOverrides) state.toolOrderOverrides = {};
  state.toolOrderOverrides[toolId] = qty;
  const tool = state.tools.find((t) => t.id === toolId);
  if (tool?.ordered) tool.orderedQty = qty;
  persist();
}

function archiveOrderEvent(tool, qty, action) {
  if (!state.orderArchive) state.orderArchive = [];
  if (!state.orderHistory) state.orderHistory = [];
  const entry = {
    id: `order-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    at: new Date().toISOString(),
    action,
    toolId: tool.id,
    tNumber: tool.tNumber,
    label: tool.label,
    size: formatToolSize(tool),
    manufacturer: tool.manufacturer || "Ohne Hersteller",
    articleNo: tool.articleNo,
    shelf: tool.shelf,
    qty: Number(qty || 0),
    user: currentUser?.name || "System",
  };
  state.orderArchive.unshift(entry);
  state.orderHistory.unshift(entry);
}

function cleanupOrderArchive() {
  if (!state.orderArchive) state.orderArchive = [];
  const now = Date.now();
  const sixWeeksMs = 6 * 7 * 24 * 60 * 60 * 1000;
  state.orderArchive = state.orderArchive.filter(
    (e) => now - new Date(e.at).getTime() <= sixWeeksMs,
  );
}

async function markToolOrdered(toolId, ordered) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;

  const nextOrderedQty = ordered ? Math.max(1, effectiveOrderQty(tool)) : 0;

  const { data: updated, error } = await supabaseClient
    .from("tools")
    .update({
      ordered: !!ordered,
      ordered_qty: Number(nextOrderedQty || 0),
    })
    .eq("id", toolId)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Markieren als bestellt:", error);
    return alert(
      `Bestellstatus konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  const index = state.tools.findIndex((t) => t.id === toolId);
  if (index !== -1) {
    state.tools[index] = normalizeToolFromDb(updated);
  }

  if (ordered) {
    archiveOrderEvent(
      state.tools[index],
      state.tools[index].orderedQty,
      "mark_ordered",
    );
  }

  persist();
  render();
}

async function restockTool(toolId) {
  if (currentUser.role !== "admin") return;

  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const qty = Number(tool.orderedQty || tool.ordered_qty || 0);
  if (!Number.isFinite(qty) || qty <= 0) {
    alert("Keine Bestellmenge vorhanden.");
    return;
  }

  const oldStock = Number(tool.stock || 0);
  const newStock = oldStock + qty;
  const returnFrom = tool.isBorrowed ? tool.borrowedTo || "-" : "";

  const { error } = await supabaseClient
    .from("tools")
    .update({
      stock: newStock,
      ordered: false,
      ordered_qty: 0,
      ...buildReturnToolUpdateFields(tool),
    })
    .eq("id", tool.id);

  if (error) {
    console.error("Fehler beim Einlagern des Werkzeugs:", error);
    return alert(
      `Werkzeug konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  const loadedTools = await loadToolsFromSupabase();
  if (Array.isArray(loadedTools) && loadedTools.length > 0) {
    state.tools = loadedTools;
  }

  let refreshed = state.tools.find((t) => t.id === tool.id);

  if (!refreshed) {
    console.warn("Werkzeug wurde nach dem Speichern nicht neu geladen.", {
      toolId: tool.id,
    });
    state.tools = state.tools.map((t) =>
      t.id === tool.id
        ? {
            ...t,
            stock: newStock,
            ordered: false,
            orderedQty: 0,
            ordered_qty: 0,
            isBorrowed: false,
            borrowedTo: "",
            borrowedAt: "",
          }
        : t,
    );
    refreshed = state.tools.find((t) => t.id === tool.id) || {
      ...tool,
      stock: newStock,
      ordered: false,
      orderedQty: 0,
      ordered_qty: 0,
      isBorrowed: false,
      borrowedTo: "",
      borrowedAt: "",
    };
  }

  if (
    refreshed.ordered ||
    Number(refreshed.orderedQty || refreshed.ordered_qty || 0) > 0
  ) {
    alert(
      "Bestand wurde gespeichert, aber Bestellstatus wurde nicht zurückgesetzt.",
    );
    return;
  }

  archiveOrderEvent(refreshed, qty, "restock");
  console.log("Journal wird geschrieben aus restockTool", {
    tool: refreshed,
    oldStock,
    newStock,
    qty,
  });
  await addToolJournalEntry({
    toolId: refreshed.id,
    toolTNumber: refreshed.tNumber,
    toolLabel: refreshed.label,
    action: tool.isBorrowed
      ? `Werkzeug zurück von ${returnFrom}`
      : `Einlagerung bestellt ${qty}`,
    qty,
    stockBefore: oldStock,
    stockAfter: newStock,
    user: currentUser?.name || "System",
  });
  await refreshToolsAndRender();
}

async function bookToolChange(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const takeOut = await askYesNoCentered("Werkzeug Entnahme?");
  let qty = 0;

  if (takeOut) {
    const defaultQty = String(getDefaultToolBookingQty(tool));

    qty = await askNumberCentered("Entnahmemenge eingeben:", defaultQty);
    if (qty === null) return;

    if (!Number.isFinite(qty) || qty <= 0) {
      return alert("Bitte eine gültige Entnahmemenge eingeben.");
    }

    if (Number(tool.stock || 0) < Number(qty || 0)) {
      return alert("Nicht genügend Bestand vorhanden.");
    }
  }

  let updatedTool = tool;
  let borrowedTo = "";
  let isBorrowing = false;

  if (takeOut && qty > 0) {
    isBorrowing = await askYesNoCentered("Als Ausleihe buchen?");
    if (isBorrowing) {
      borrowedTo = await askTextCentered(
        "Kostenträger/Maschine bei Ausleihe",
        "z. B. 350 / Abteilung Montage",
      );
      if (borrowedTo === null) return;
      if (!borrowedTo) {
        return alert("Bitte Kostenträger/Maschine eintragen.");
      }
    }

    const nextStock = Math.max(0, Number(tool.stock || 0) - Number(qty || 0));

    const { data, error } = await supabaseClient
      .from("tools")
      .update({
        stock: nextStock,
        ...(isBorrowing ? buildBorrowToolUpdateFields(borrowedTo) : {}),
      })
      .eq("id", toolId)
      .select()
      .maybeSingle();

    if (error) {
      console.error("Fehler bei Werkzeug-Entnahme:", error);
      return alert(
        `Entnahme konnte nicht gespeichert werden: ${error.message}`,
      );
    }

    if (!data) {
      console.error("Keine Werkzeugzeile nach Entnahme zurückgegeben.", {
        toolId,
        nextStock,
      });
      return alert(
        "Entnahme wurde nicht bestätigt. Bitte Seite neu laden und erneut versuchen.",
      );
    }

    if (error) {
      console.error("Fehler bei Werkzeug-Entnahme:", error);
      return alert(
        `Entnahme konnte nicht gespeichert werden: ${error.message}`,
      );
    }

    updatedTool = normalizeToolFromDb(data);

    const index = state.tools.findIndex((t) => t.id === toolId);
    if (index !== -1) {
      state.tools[index] = updatedTool;
    }
  }

  console.log("Journal wird geschrieben aus bookToolChange", {
    tool: updatedTool,
    oldStock: Number(tool.stock || 0),
    newStock: takeOut ? Number(updatedTool.stock || 0) : Number(tool.stock || 0),
    qty,
  });
  await addToolJournalEntry({
    toolId: updatedTool.id,
    toolTNumber: updatedTool.tNumber,
    toolLabel: updatedTool.label,
    action: takeOut
      ? isBorrowing
        ? `Werkzeug ausgeliehen an ${borrowedTo}`
        : `Werkzeugwechsel + Entnahme ${qty}`
      : "Werkzeugwechsel ohne Entnahme",
    qty,
    stockBefore: Number(tool.stock || 0),
    stockAfter: takeOut ? Number(updatedTool.stock || 0) : Number(tool.stock || 0),
    user: currentUser.name,
  });

  await refreshToolsAndRender();
}

async function deleteTool(toolId) {
  if (currentUser.role !== "admin") return;

  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const yes = await askYesNoCentered(
    `Werkzeug T ${tool.tNumber} wirklich löschen?`,
  );
  if (!yes) return;

  const { error } = await supabaseClient
    .from("tools")
    .delete()
    .eq("id", toolId);

  if (error) {
    console.error("Fehler beim Löschen des Werkzeugs:", error);
    return alert(`Werkzeug konnte nicht gelöscht werden: ${error.message}`);
  }

  state.tools = state.tools.filter((t) => t.id !== toolId);
  state.toolJournal = state.toolJournal.filter((j) => j.toolId !== toolId);

  persist();
  render();
}

async function undoToolJournalEntry(entryId) {
  const idx = state.toolJournal.findIndex((j) => j.id === entryId);
  if (idx === -1) return;

  const entry = state.toolJournal[idx];
  const tool = state.tools.find((t) => t.id === entry.toolId);

  if (tool && Number(entry.qty) > 0) {
    const nextStock = Number(tool.stock || 0) + Number(entry.qty || 0);

    const { data, error } = await supabaseClient
      .from("tools")
      .update({
        stock: nextStock,
      })
      .eq("id", tool.id)
      .select()
      .single();

    if (error) {
      console.error(
        "Fehler beim Rückgängig machen des Journal-Eintrags:",
        error,
      );
      return alert(
        `Rückgängig konnte nicht gespeichert werden: ${error.message}`,
      );
    }

    const updatedTool = normalizeToolFromDb(data);
    const toolIndex = state.tools.findIndex((t) => t.id === tool.id);
    if (toolIndex !== -1) {
      state.tools[toolIndex] = updatedTool;
    }
  }

  state.toolJournal.splice(idx, 1);
  await refreshToolsAndRender();
}

function resetToolFilters() {
  state.toolFilters = {
    search: "",
    label: "",
    tNumber: "",
    diameter: "",
    holder: "",
    imageStatus: "",
  };
  persist();
  render();
}

function applyToolFilters() {
  const search = document.getElementById("toolSearch")?.value || "";
  const label = document.getElementById("toolFilterLabel")?.value || "";
  const tNumber = document.getElementById("toolFilterT")?.value || "";
  const diameter = document.getElementById("toolFilterD")?.value || "";
  const holder = document.getElementById("toolFilterHolder")?.value || "";
  const imageStatus =
    document.getElementById("toolFilterImageStatus")?.value || "";
  state.toolFilters = {
    search,
    label,
    tNumber,
    diameter,
    holder,
    imageStatus,
  };
  persist();
  render();
}

function getOrderCandidateGroups() {
  return state.tools
    .filter((t) => shouldOrderTool(t) || t.ordered)
    .reduce((acc, tool) => {
      const maker = (tool.manufacturer || "").trim() || "Ohne Hersteller";
      if (!acc[maker]) acc[maker] = [];
      acc[maker].push(tool);
      return acc;
    }, {});
}

function ensureSelectedOrderListManufacturer(groups) {
  const makers = Object.keys(groups).sort((a, b) => a.localeCompare(b));
  if (!makers.length) {
    state.selectedOrderListManufacturer = "";
    return "";
  }
  if (!makers.includes(state.selectedOrderListManufacturer || "")) {
    state.selectedOrderListManufacturer = makers[0];
  }
  return state.selectedOrderListManufacturer;
}

function openOrderListPopup() {
  state.orderListPopupOpen = true;
  const groups = getOrderCandidateGroups();
  ensureSelectedOrderListManufacturer(groups);
  persist();
  render();
}

function closeOrderListPopup() {
  state.orderListPopupOpen = false;
  render();
}

function setOrderListManufacturer(value) {
  state.selectedOrderListManufacturer = value || "";
  persist();
  render();
}

function setOrderStatsView(view) {
  if (!["week", "month", "year"].includes(view)) return;
  state.orderStatsView = view;
  persist();
  render();
}

function getOrderStatsRange(view) {
  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);
  if (view === "week") {
    start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
    end.setDate(start.getDate() + 6);
  } else if (view === "month") {
    start.setDate(1);
    end.setMonth(start.getMonth() + 1, 0);
  } else {
    start.setMonth(0, 1);
    end.setMonth(11, 31);
  }
  return { start, end };
}

function buildOrderFrequency(view) {
  const { start, end } = getOrderStatsRange(view);
  const rows = (state.orderHistory || []).filter((e) => {
    const d = new Date(e.at);
    return (
      d >= start &&
      d <= end &&
      (e.action === "mark_ordered" || e.action === "restock")
    );
  });

  const map = {};
  rows.forEach((e) => {
    const key = `${e.toolId}`;
    if (!map[key]) {
      map[key] = {
        toolId: e.toolId,
        label: e.label,
        size: e.size,
        manufacturer: e.manufacturer,
        articleNo: e.articleNo,
        count: 0,
        qtyTotal: 0,
      };
    }
    map[key].count += 1;
    map[key].qtyTotal += Number(e.qty || 0);
  });

  return Object.values(map).sort((a, b) => b.count - a.count);
}

function getOptimalQtySuggestion(tool) {
  const yearData = buildOrderFrequency("year").find(
    (r) => r.toolId === tool.id,
  );
  if (!yearData || yearData.count < 6) return null;
  const avg = Math.ceil(yearData.qtyTotal / yearData.count);
  return Math.max(1, avg);
}

function shouldShowOptimalQtySuggestion(tool) {
  const suggestion = getOptimalQtySuggestion(tool);
  if (!Number.isFinite(suggestion) || suggestion <= 0) return false;
  const stateEntry = state.orderSuggestionState?.[tool.id];
  if (!stateEntry) return true;
  if (stateEntry.accepted) return false;
  const last = new Date(stateEntry.lastDecisionAt || 0).getTime();
  const days30 = 30 * 24 * 60 * 60 * 1000;
  return Date.now() - last >= days30;
}

function applyOptimalQtySuggestion(toolId) {
  if (currentUser.role !== "admin") return;
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;
  const s = getOptimalQtySuggestion(tool);
  if (!Number.isFinite(s) || s <= 0) return;
  tool.optimalStock = s;
  if (!state.orderSuggestionState) state.orderSuggestionState = {};
  state.orderSuggestionState[toolId] = {
    lastDecisionAt: new Date().toISOString(),
    accepted: true,
  };
  persist();
  render();
}

function rejectOptimalQtySuggestion(toolId) {
  if (currentUser.role !== "admin") return;
  if (!state.orderSuggestionState) state.orderSuggestionState = {};
  state.orderSuggestionState[toolId] = {
    lastDecisionAt: new Date().toISOString(),
    accepted: false,
  };
  persist();
  render();
}

function renderOrderStats() {
  if (currentUser.role !== "admin")
    return `<div class='bg-white rounded-xl shadow p-4'><p>Kein Zugriff.</p></div>`;
  const view = state.orderStatsView || "week";
  const freq = buildOrderFrequency(view);

  const rows = freq
    .map(
      (r) => `<tr class='border-b'>
    <td class='p-2'>${r.label}</td>
    <td class='p-2'>${r.size}</td>
    <td class='p-2'>${r.manufacturer}</td>
    <td class='p-2'>${r.articleNo}</td>
    <td class='p-2'>${r.count}</td>
    <td class='p-2'>${r.qtyTotal}</td>
  </tr>`,
    )
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4 space-y-3'>
    <h2 class='text-lg font-semibold'>Bestell-Statistik</h2>
    <div class='flex gap-2'>
      <button class='px-2 py-1 rounded ${view === "week" ? "bg-slate-900 text-white" : "bg-slate-200"}' onclick="setOrderStatsView('week')">Woche</button>
      <button class='px-2 py-1 rounded ${view === "month" ? "bg-slate-900 text-white" : "bg-slate-200"}' onclick="setOrderStatsView('month')">Monat</button>
      <button class='px-2 py-1 rounded ${view === "year" ? "bg-slate-900 text-white" : "bg-slate-200"}' onclick="setOrderStatsView('year')">Jahr</button>
    </div>
    <div class='border rounded-lg overflow-auto max-h-[60vh]'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Bezeichnung</th>
            <th class='p-2 text-left'>Größe</th>
            <th class='p-2 text-left'>Hersteller</th>
            <th class='p-2 text-left'>Artikelnummer</th>
            <th class='p-2 text-left'>Bestellvorgänge</th>
            <th class='p-2 text-left'>Gesamtmenge</th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td class="p-2" colspan="6">Keine Daten im gewählten Zeitraum.</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;
}

function renderTools(view = "all") {
  const toolsLoaded = state.ui?.toolsInitialLoaded === true;
  const hasTools = Array.isArray(state.tools) && state.tools.length > 0;

  if (!toolsLoaded && !hasTools) {
    setTimeout(() => forceLoadToolPageData(), 0);

    return `
    <div class="bg-white rounded-xl shadow p-6">
      <h2 class="text-xl font-bold mb-2">Werkzeugverwaltung</h2>
      <p class="text-slate-600">Werkzeugdaten werden aus Supabase geladen...</p>
      <button class="mt-4 px-4 py-2 bg-slate-900 text-white rounded"
        onclick="forceLoadToolPageData()">
        Erneut laden
      </button>
    </div>
  `;
  }

  ensureStorageHighlightStyles();

  const labels = getToolLabels();
  const filters = state.toolFilters || {
    search: "",
    label: "",
    tNumber: "",
    diameter: "",
    holder: "",
    imageStatus: "",
  };
  const search = (filters.search || "").toLowerCase();
  const filterLabel = filters.label || "";
  const filterT = filters.tNumber || "";
  const filterD = filters.diameter || "";
  const filterHolder = filters.holder || "";
  const imageStatus = filters.imageStatus || "";
  const isAdmin = currentUser?.role === "admin";
  const toolsLoadingBanner = state.ui?.toolsLoading
    ? `<div class='border border-blue-200 bg-blue-50 text-blue-800 rounded p-3 text-sm'>Werkzeugdaten werden geladen...</div>`
    : "";
  const toolRefreshStatusText = toolRealtimeChannel
    ? "Live-Aktualisierung aktiv"
    : "Auto-Aktualisierung alle 15 Sekunden aktiv";

  const filterLabelOptions = labels
    .map(
      (l) =>
        `<option value="${l}" ${l === filterLabel ? "selected" : ""}>${l}</option>`,
    )
    .join("");

  const tools = state.tools.filter((t) => {
    const bySearch =
      !search ||
      `${t.tNumber} ${t.label} ${t.diameter} ${t.articleNo} ${t.holder || ""} ${t.manufacturer || ""} ${getToolMaterialNameById(t.materialId)}`
        .toLowerCase()
        .includes(search);
    const byLabel = !filterLabel || t.label === filterLabel;
    const byT = !filterT || String(t.tNumber).includes(filterT);
    const byD = !filterD || String(t.diameter).includes(filterD);
    const byHolder = !filterHolder || t.holder === filterHolder;
    const imagePath = getToolImagePath(t);
    const byImageStatus =
      !imageStatus ||
      (imageStatus === "withPath" && !!imagePath) ||
      (imageStatus === "withoutPath" && !imagePath);
    return bySearch && byLabel && byT && byD && byHolder && byImageStatus;
  });

  const toolRows = tools
    .map((t) => {
      const stockWarningLevel = getToolStockWarningLevel(t);
      const warningBadge =
        stockWarningLevel === "critical"
          ? '<span class="inline-flex px-2 py-1 rounded bg-rose-100 text-rose-800 text-xs font-semibold">🚨 Mindestbestand unterschritten</span>'
          : stockWarningLevel === "warning"
            ? '<span class="inline-flex px-2 py-1 rounded bg-orange-100 text-orange-800 text-xs font-semibold">⚠ Mindestbestand erreicht</span>'
            : "";
      const rowWarningClass =
        stockWarningLevel === "critical"
          ? "bg-rose-50 border-l-4 border-rose-500"
          : stockWarningLevel === "warning"
            ? "bg-orange-50 border-l-4 border-orange-400"
            : "";
      const stockWarningClass =
        stockWarningLevel === "critical"
          ? "font-bold text-rose-700"
          : stockWarningLevel === "warning"
            ? "font-bold text-orange-700"
            : "";
      const orderedBadge = t.ordered
        ? '<span class="inline-flex px-2 py-1 rounded bg-emerald-100 text-emerald-800 text-xs font-semibold">Bestellt</span>'
        : "";
      const borrowedBadge = renderToolBorrowedBadge(t);
      const statusText =
        [orderedBadge, warningBadge, borrowedBadge].filter(Boolean).join(" ") ||
        "-";
      const imagePath = getToolImagePath(t);
      const imageCell = imagePath
        ? `<button class='border rounded bg-white p-0.5 hover:bg-slate-50' onclick="openToolImagePopup('${t.id}')" title="Werkzeugbild öffnen">
            <img src="${escapeHtml(imagePath)}" alt="Bild T ${escapeHtml(normalizeToolImageNumber(t.tNumber))}" class="w-9 h-9 object-contain" onerror="this.style.display='none'; this.parentElement.querySelector('[data-tool-image-missing]').classList.remove('hidden')" />
            <span data-tool-image-missing class="hidden text-xs text-rose-700">kein Bild</span>
          </button>`
        : "-";
      const insertRadiusLine =
        t.insertTool && t.insertRadius
          ? `<div class='text-xs text-slate-500'>Plattenradius: R ${escapeHtml(t.insertRadius)}</div>`
          : "";

      return `<tr class='border-b ${rowWarningClass}'>
        <td class='p-2 align-top whitespace-nowrap'>T ${t.tNumber}</td>
        <td class='p-2 align-top'>${imageCell}</td>
        <td class='p-2 align-top'>${t.label}</td>
        <td class='p-2 align-top whitespace-nowrap'>${renderToolSizeCell(t)}${insertRadiusLine}</td>
        ${isAdmin ? `<td class='p-2 align-top'>${getToolMaterialNameById(t.materialId)}</td>` : ""}
        <td class='p-2 align-top whitespace-nowrap'>${t.shelf}</td>
        <td class='p-2 align-top'>${t.articleNo}</td>
        ${isAdmin ? `<td class='p-2 align-top whitespace-nowrap'>${t.holder || "-"}</td>` : ""}
        <td class='p-2 align-top ${stockWarningClass}'>${t.stock}</td>
        <td class='p-2 align-top'>${t.minStock}</td>
        ${isAdmin ? `<td class='p-2 align-top'>${t.manufacturer || "-"}</td>` : ""}
        <td class='p-2 align-top'>${statusText}</td>
        <td class='p-2 align-top'>
          <div class='flex flex-wrap gap-1'>
          <button class='px-2 py-1 rounded bg-emerald-700 text-white text-xs' onclick="bookToolChange('${t.id}')">Wechsel</button>
          ${
            isAdmin
              ? `<button class='px-2 py-1 rounded bg-slate-700 text-white text-xs' onclick="openToolQrPopup('${t.id}')">QR</button>
                 <button class='px-2 py-1 rounded bg-slate-800 text-white text-xs' onclick="openToolHistory('${t.id}')">Historie</button>
                 <button class='px-2 py-1 rounded bg-amber-600 text-white text-xs' onclick="editTool('${t.id}')">Bearbeiten</button>
                 <button class='px-2 py-1 rounded bg-rose-700 text-white text-xs' onclick="deleteTool('${t.id}')">Löschen</button>`
              : ""
          }
          </div>
        </td>
      </tr>`;
    })
    .join("");

  const todoTools = state.tools.filter(
    (t) => isToolBelowMinStock(t) && !t.ordered,
  );
  const orderedTools = state.tools.filter((t) => {
    return (
      t.ordered === true ||
      Number(t.orderedQty || 0) > 0 ||
      Number(t.ordered_qty || 0) > 0
    );
  });
  const orderGroups = getOrderCandidateGroups();
  const availableManufacturers = Object.keys(orderGroups).sort((a, b) =>
    a.localeCompare(b),
  );
  const selectedManufacturer = ensureSelectedOrderListManufacturer(orderGroups);
  const selectedManufacturerTools = selectedManufacturer
    ? orderGroups[selectedManufacturer] || []
    : [];

  const todoRows = todoTools
    .map(
      (t) => {
        const stockWarningLevel = getToolStockWarningLevel(t);
        const stockBadge =
          stockWarningLevel === "critical"
            ? '<span class="inline-flex px-2 py-1 rounded bg-rose-100 text-rose-800 text-xs font-semibold">🚨 unterschritten</span>'
            : '<span class="inline-flex px-2 py-1 rounded bg-orange-100 text-orange-800 text-xs font-semibold">⚠ erreicht</span>';
        const stockClass =
          stockWarningLevel === "critical"
            ? "font-bold text-rose-700"
            : "font-bold text-orange-700";

        return `<tr class='border-b'>
        <td class='p-2'>T ${t.tNumber}</td>
        <td class='p-2'>${t.label}</td>
        <td class='p-2'>${formatToolSize(t)}</td>
        <td class='p-2 ${stockClass}'>${t.stock}</td>
        <td class='p-2'>${t.minStock}</td>
        <td class='p-2'>${t.holder || "-"}</td>
        <td class='p-2'>${t.shelf || "-"}</td>
        <td class='p-2'>${stockBadge}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="markToolOrdered('${t.id}', true)">Bestellen</button>
        </td>
      </tr>`;
      },
    )
    .join("");

  const orderedCards = orderedTools
    .map((t) => {
      const isThreadTool = String(t.label || "").includes("Gewinde");
      const isRadiusTool = String(t.label || "").includes("Radiusfräser");
      const detailLine = isThreadTool
        ? `<div><span class='font-semibold'>Steigung:</span> ${t.threadPitch || t.pitch || "-"}</div>`
        : isRadiusTool
          ? `<div><span class='font-semibold'>Radius:</span> R ${t.cornerRadius || "-"}</div>`
          : "";
      const insertRadiusLine =
        t.insertTool && t.insertRadius
          ? `<div><span class='font-semibold'>Plattenradius:</span> R ${escapeHtml(t.insertRadius)}</div>`
          : "";

      return `<div class='border rounded-lg p-3 bg-emerald-50 space-y-3'>
        <div class='grid md:grid-cols-2 gap-x-6 gap-y-2 text-sm'>
          <div><span class='font-semibold'>T-Nummer:</span> T ${t.tNumber}</div>
          <div><span class='font-semibold'>Bezeichnung:</span> ${t.label}</div>
          <div><span class='font-semibold'>Durchmesser:</span> ${formatToolSize(t)}</div>
          ${renderToolOverhangDetailLine(t)}
          ${renderToolBorrowedDetailLine(t, "text-orange-700")}
          <div><span class='font-semibold'>Schneidwerkstoff:</span> ${getToolMaterialNameById(t.materialId)}</div>
          ${detailLine}
          ${insertRadiusLine}
          <div><span class='font-semibold'>Menge:</span> ${t.orderedQty || effectiveOrderQty(t)}</div>
          <div><span class='font-semibold'>Artikelnummer:</span> ${t.articleNo}</div>
          <div><span class='font-semibold'>Lagerfach:</span> ${t.shelf}</div>
        </div>
        <div class='flex flex-col sm:flex-row gap-2'>
          <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="markToolOrdered('${t.id}', false)">Bestellt</button>
          <button class='px-2 py-1 rounded bg-slate-700 text-white' onclick="restockTool('${t.id}')">Einlagern</button>
        </div>
      </div>`;
    })
    .join("");

  const allJournalEntries = Array.isArray(state.toolJournal)
    ? state.toolJournal
    : [];
  const journalEntries = getFilteredToolJournalEntries();
  const journalFilters = state.toolJournalFilters || {
    search: "",
    action: "",
    range: "",
  };
  const journalFilterBar = `<div class='border rounded p-3 bg-slate-50 mb-3'>
    <div class='grid md:grid-cols-[2fr,1fr,1fr,auto,auto] gap-2 items-end'>
      <label class='text-sm font-medium'>Suche
        <input id='toolJournalSearch' class='border rounded p-2 w-full mt-1' placeholder='T, Werkzeug, Aktion, Benutzer' value='${escapeHtml(journalFilters.search || "")}' />
      </label>
      <label class='text-sm font-medium'>Aktion
        <select id='toolJournalAction' class='border rounded p-2 w-full mt-1'>
          <option value='' ${journalFilters.action ? "" : "selected"}>Alle</option>
          <option value='Entnahme' ${journalFilters.action === "Entnahme" ? "selected" : ""}>Entnahme</option>
          <option value='Einlagerung' ${journalFilters.action === "Einlagerung" ? "selected" : ""}>Einlagerung</option>
          <option value='Wechsel' ${journalFilters.action === "Wechsel" ? "selected" : ""}>Wechsel</option>
        </select>
      </label>
      <label class='text-sm font-medium'>Zeitraum
        <select id='toolJournalRange' class='border rounded p-2 w-full mt-1'>
          <option value='' ${journalFilters.range ? "" : "selected"}>Alle</option>
          <option value='today' ${journalFilters.range === "today" ? "selected" : ""}>Heute</option>
          <option value='7d' ${journalFilters.range === "7d" ? "selected" : ""}>7 Tage</option>
          <option value='30d' ${journalFilters.range === "30d" ? "selected" : ""}>30 Tage</option>
        </select>
      </label>
      <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='applyToolJournalFilters()'>Filter anwenden</button>
      <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='resetToolJournalFilters()'>Zurücksetzen</button>
    </div>
    <div class='text-xs text-slate-500 mt-2'>${journalEntries.length} von ${allJournalEntries.length} Einträgen</div>
  </div>`;

  const journalRows = journalEntries
    .slice(0, 80)
    .map(
      (entry) => `<tr class='border-b'>
        <td class='p-2'>${escapeHtml(entry.createdAt ? new Date(entry.createdAt).toLocaleString() : entry.at || "")}</td>
        <td class='p-2'>${escapeHtml(entry.user || "-")}</td>
        <td class='p-2'>T ${escapeHtml(entry.toolTNumber || "-")} · ${escapeHtml(entry.toolLabel || "-")}</td>
        <td class='p-2'>${escapeHtml(entry.action || "-")}</td>
        <td class='p-2'>${escapeHtml(entry.qty ?? "-")}</td>
        <td class='p-2'>${escapeHtml(entry.stockBefore ?? "-")} / ${escapeHtml(entry.stockAfter ?? "-")}</td>
      </tr>`,
    )
    .join("");

  const toolStatistics = buildToolStatistics().slice(0, 10);
  const toolStatisticsRange = state.toolStatisticsRange || "30d";
  const toolStatisticsRangeLabels = {
    today: "Heute",
    "7d": "7 Tage",
    "30d": "30 Tage",
    all: "Gesamt",
  };
  const toolStatisticsRangeButton = (range, label) => {
    const active = toolStatisticsRange === range;
    return `<button class='px-3 py-1 rounded text-sm ${active ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-700"}' onclick="setToolStatisticsRange('${range}')">${label}</button>`;
  };
  const toolStatisticsRows = toolStatistics
    .map(
      (stat) => `<tr class='border-b'>
        <td class='p-2'>T ${escapeHtml(stat.toolTNumber || "-")}</td>
        <td class='p-2'>${escapeHtml(stat.toolLabel || "-")}</td>
        <td class='p-2 text-rose-700 font-semibold'>${stat.withdrawals}</td>
        <td class='p-2 text-emerald-700 font-semibold'>${stat.restocks}</td>
        <td class='p-2'>${stat.totalActions}</td>
        <td class='p-2'>${stat.currentStock}</td>
        <td class='p-2'>${escapeHtml(stat.lastActionAt ? new Date(stat.lastActionAt).toLocaleString() : "-")}</td>
      </tr>`,
    )
    .join("");

  const toolStatisticsSection = `<div class='border rounded p-3 bg-white'>
    <div class='flex flex-wrap items-center justify-between gap-2 mb-3'>
      <div>
        <h4 class='font-semibold'>Werkzeugstatistik</h4>
        <div class='text-xs text-slate-500'>Zeitraum: ${escapeHtml(toolStatisticsRangeLabels[toolStatisticsRange] || "30 Tage")}</div>
      </div>
      <div class='flex flex-wrap gap-2'>
        ${toolStatisticsRangeButton("today", "Heute")}
        ${toolStatisticsRangeButton("7d", "7 Tage")}
        ${toolStatisticsRangeButton("30d", "30 Tage")}
        ${toolStatisticsRangeButton("all", "Gesamt")}
      </div>
    </div>
    <div class='overflow-auto'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2 text-left'>T-Nummer</th>
            <th class='p-2 text-left'>Werkzeug</th>
            <th class='p-2 text-left'>Entnahmen</th>
            <th class='p-2 text-left'>Einlagerungen</th>
            <th class='p-2 text-left'>Gesamtaktionen</th>
            <th class='p-2 text-left'>Bestand</th>
            <th class='p-2 text-left'>Letzte Aktivität</th>
          </tr>
        </thead>
        <tbody>${toolStatisticsRows || '<tr><td class="p-2" colspan="7">Keine Statistikdaten vorhanden.</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;

  const storageMap = buildToolStorageMap();
  const storageSearchMatches = findStorageSearchMatches();
  const storageSearchLocations = new Set(
    storageSearchMatches.map(getToolStorageLocationKey).filter(Boolean),
  );
  const storageSearchValue = state.storageSearch || "";
  const storageSearchRows = storageSearchMatches
    .map((tool) => {
      const locationKey = getToolStorageLocationKey(tool);
      return `<tr class='border-b'>
        <td class='p-2'>T ${escapeHtml(tool.tNumber || "-")}</td>
        <td class='p-2'>${escapeHtml(tool.label || "-")}</td>
        <td class='p-2'>${escapeHtml(locationKey || "-")}</td>
        <td class='p-2'>
          ${
            locationKey
              ? `<button class='px-2 py-1 rounded bg-slate-900 text-white text-xs' onclick="openStorageLocationModal('${locationKey}')">Fach öffnen</button>`
              : "-"
          }
        </td>
      </tr>`;
    })
    .join("");
  const storageCellClasses = {
    empty: "bg-slate-100 text-slate-500 border-slate-200",
    ok: "bg-green-100 text-green-900 border-green-200",
    warning: "bg-orange-200 text-orange-950 border-orange-300",
    critical: "bg-red-200 text-red-950 border-red-300",
  };
  const storageRackColumns = Array.from({ length: 26 }, (_, index) => {
    const rackNumber = index + 1;
    const cells = [...STORAGE_LETTERS]
      .reverse()
      .map((letter) => {
        const locationKey = `${rackNumber}${letter}`;
        const locationTools = storageMap[locationKey] || [];
        const cellState = getStorageCellState(locationTools);
        const isSearchMatch = storageSearchLocations.has(locationKey);
        const hasBorrowedTool = locationTools.some((tool) => tool.isBorrowed);
        return `<button
          class='border rounded px-2 py-1 text-xs text-left ${storageCellClasses[cellState]} ${isSearchMatch ? "storage-hit-blink" : ""}'
          onclick="openStorageLocationModal('${locationKey}')"
          title='Lagerfach ${locationKey}'
        >
          <span class='font-semibold'>${locationKey}</span>
          ${isSearchMatch ? "<span class='block text-[10px] font-semibold text-blue-700'>Treffer</span>" : ""}
          ${hasBorrowedTool ? "<span class='ml-1 inline-flex px-1 rounded bg-orange-500 text-white text-[10px] font-bold'>A</span>" : ""}
          <span class='block text-[10px]'>${locationTools.length || ""}</span>
        </button>`;
      })
      .join("");

    return `<div class='min-w-[76px] border rounded bg-white p-2'>
      <div class='text-xs font-semibold text-center mb-2'>Regal ${rackNumber}</div>
      <div class='grid gap-1'>${cells}</div>
    </div>`;
  }).join("");

  const storageSearchSection = `<div class='border rounded p-3 bg-white'>
    <h4 class='font-semibold mb-2'>Fachsuche</h4>
    <div class='border rounded bg-slate-50 p-3'>
      <div class='grid md:grid-cols-[1fr,auto,auto] gap-2 items-end'>
        <label class='text-sm font-medium'>Werkzeug suchen
          <input id='storageSearchInput' class='border rounded p-2 w-full mt-1' placeholder='T-Nummer, Bezeichnung, Artikelnummer oder Hersteller suchen' value='${escapeHtml(storageSearchValue)}' />
        </label>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='applyStorageSearch()'>Suchen</button>
        <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='resetStorageSearch()'>Zurücksetzen</button>
      </div>
      <div class='text-xs text-slate-500 mt-2'>${storageSearchMatches.length} Treffer</div>
      ${
        storageSearchMatches.length
          ? `<div class='overflow-auto mt-3 border rounded bg-white max-h-[24vh]'>
              <table class='w-full text-sm'>
                <thead class='bg-slate-100 sticky top-0'>
                  <tr>
                    <th class='p-2 text-left'>T-Nummer</th>
                    <th class='p-2 text-left'>Bezeichnung</th>
                    <th class='p-2 text-left'>Fach</th>
                    <th class='p-2 text-left'>Aktion</th>
                  </tr>
                </thead>
                <tbody>${storageSearchRows}</tbody>
              </table>
            </div>`
          : ""
      }
    </div>
  </div>`;

  const storageRackSection = `<div class='border rounded p-3 bg-white'>
    <h4 class='font-semibold mb-2'>Lagerfach-/Regalansicht</h4>
    <div class='flex gap-2 text-xs text-slate-600 flex-wrap mb-3'>
      <span class='inline-flex items-center gap-1'><span class='inline-block w-3 h-3 rounded bg-slate-100 border'></span> Leer</span>
      <span class='inline-flex items-center gap-1'><span class='inline-block w-3 h-3 rounded bg-green-100 border border-green-200'></span> OK</span>
      <span class='inline-flex items-center gap-1'><span class='inline-block w-3 h-3 rounded bg-orange-200 border border-orange-300'></span> Mindestbestand erreicht</span>
      <span class='inline-flex items-center gap-1'><span class='inline-block w-3 h-3 rounded bg-red-200 border border-red-300'></span> Mindestbestand unterschritten</span>
    </div>
    <div class='overflow-x-auto pb-2'>
      <div class='flex gap-2'>${storageRackColumns}</div>
    </div>
  </div>`;
  const storageOverviewSection =
    view === "storageSearch"
      ? storageSearchSection
      : view === "storageOverview"
        ? storageRackSection
        : `${storageSearchSection}${storageRackSection}`;

  const manufacturerOptionsForPopup = availableManufacturers
    .map(
      (maker) =>
        `<option value="${maker}" ${maker === selectedManufacturer ? "selected" : ""}>${maker}</option>`,
    )
    .join("");

  const selectedManufacturerRows = selectedManufacturerTools
    .map(
      (t) => `<tr class='border-b'>
        <td class='p-2'>T ${t.tNumber}</td>
        <td class='p-2'>${t.label}</td>
        <td class='p-2'>${formatToolSize(t)}</td>
        <td class='p-2'>${getToolMaterialNameById(t.materialId)}</td>
        <td class='p-2'>${t.articleNo || "-"}</td>
        <td class='p-2'>${t.stock}/${t.minStock}</td>
        <td class='p-2'>
          <input
            type='number'
            min='0'
            class='border rounded p-1 w-24'
            value='${effectiveOrderQty(t)}'
            onchange="setToolOrderOverride('${t.id}', this.value)"
          />
        </td>
      </tr>`,
    )
    .join("");

  const orderListPopup = state.orderListPopupOpen
    ? `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-6xl max-h-[88vh] overflow-auto p-4">
          <div class="flex items-center justify-between mb-3 gap-3">
            <h3 class="text-lg font-bold">Bestellliste nach Hersteller</h3>
            <button class="px-2 py-1 rounded bg-slate-200" onclick="closeOrderListPopup()">Schließen</button>
          </div>
          ${
            availableManufacturers.length
              ? `<div class='border rounded p-3 mb-4 bg-slate-50'>
                  <div class='grid md:grid-cols-[280px,1fr] gap-3 items-end'>
                    <label class='text-sm font-medium'>Hersteller auswählen
                      <select class='border rounded p-2 w-full mt-1' onchange="setOrderListManufacturer(this.value)">
                        ${manufacturerOptionsForPopup}
                      </select>
                    </label>
                    <div class='text-sm text-slate-500'>Es wird immer nur ein Hersteller gleichzeitig angezeigt.</div>
                  </div>
                </div>
                <div class='border rounded p-3 bg-white'>
                  <div class='flex items-center justify-between mb-2'>
                    <h4 class='font-semibold text-base'>${selectedManufacturer}</h4>
                    <span class='text-xs text-slate-500'>${selectedManufacturerTools.length} Werkzeug(e)</span>
                  </div>
                  <table class='w-full text-sm'>
                    <thead class='bg-slate-100'>
                      <tr>
                        <th class='p-2 text-left'>T</th>
                        <th class='p-2 text-left'>Bezeichnung</th>
                        <th class='p-2 text-left'>Größe</th>
                        <th class='p-2 text-left'>Werkstoff</th>
                        <th class='p-2 text-left'>Artikelnummer</th>
                        <th class='p-2 text-left'>Bestand</th>
                        <th class='p-2 text-left'>Menge</th>
                      </tr>
                    </thead>
                    <tbody>${selectedManufacturerRows || '<tr><td class="p-2" colspan="7">Keine bestellrelevanten Werkzeuge für diesen Hersteller.</td></tr>'}</tbody>
                  </table>
                </div>`
              : '<div class="text-sm text-slate-500">Keine bestellrelevanten Werkzeuge.</div>'
          }
        </div>
      </div>`
    : "";

  const suggestionTools = state.tools.filter((t) =>
    shouldShowOptimalQtySuggestion(t),
  );
  const suggestionRows = suggestionTools
    .map(
      (t) => `<tr class='border-b'>
        <td class='p-2'>${t.label}</td>
        <td class='p-2'>${formatToolSize(t)}</td>
        <td class='p-2'>${getToolMaterialNameById(t.materialId)}</td>
        <td class='p-2'>${t.optimalStock || 0}</td>
        <td class='p-2'>${getOptimalQtySuggestion(t)}</td>
        <td class='p-2'>
          <div class='flex gap-2'>
            <button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick="applyOptimalQtySuggestion('${t.id}')">Übernehmen</button>
            <button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="rejectOptimalQtySuggestion('${t.id}')">Ablehnen</button>
          </div>
        </td>
      </tr>`,
    )
    .join("");

  const stockTableHeader = isAdmin
    ? `<tr>
        <th class='p-2 text-left'>T</th>
        <th class='p-2 text-left'>Bild</th>
        <th class='p-2 text-left'>Bezeichnung</th>
        <th class='p-2 text-left'>Ø</th>
        <th class='p-2 text-left'>Werkstoff</th>
        <th class='p-2 text-left'>Fach</th>
        <th class='p-2 text-left'>Art.</th>
        <th class='p-2 text-left'>Aufn.</th>
        <th class='p-2 text-left'>Bestand</th>
        <th class='p-2 text-left'>Min</th>
        <th class='p-2 text-left'>Herst.</th>
        <th class='p-2 text-left'>Status</th>
        <th class='p-2 text-left'>Aktionen</th>
      </tr>`
    : `<tr>
        <th class='p-2 text-left'>T</th>
        <th class='p-2 text-left'>Bild</th>
        <th class='p-2 text-left'>Bezeichnung</th>
        <th class='p-2 text-left'>Ø</th>
        <th class='p-2 text-left'>Fach</th>
        <th class='p-2 text-left'>Art.</th>
        <th class='p-2 text-left'>Bestand</th>
        <th class='p-2 text-left'>Min</th>
        <th class='p-2 text-left'>Status</th>
        <th class='p-2 text-left'>Aktionen</th>
      </tr>`;

  const showStockSection = ["all", "list", "stock"].includes(view);
  const showReorderSection = ["all", "reorder"].includes(view);
  const showJournalSection = ["all", "journal"].includes(view);
  const showStorageOverviewSection = ["all", "storageOverview", "storageSearch"].includes(view);
  const showToolStatisticsSection = view === "all";
  const showAdminActionsShell =
    isAdmin && (showReorderSection || showJournalSection || showStorageOverviewSection || showToolStatisticsSection);
  const showEmployeeExtraSections =
    !isAdmin && (showJournalSection || showStorageOverviewSection || showToolStatisticsSection);
  const toolSectionShellTitle = showStorageOverviewSection && !showReorderSection && !showJournalSection
    ? "Lageransicht"
    : "Admin-Bestandsaktionen";

  return `<div class='bg-white rounded-xl shadow p-4 space-y-4 w-full max-w-none'>
    <div class='flex items-center justify-between gap-2 flex-wrap'>
      <div class='flex items-center gap-2'>
        <h2 class='text-xl font-bold mb-1'>Werkzeugverwaltung</h2>
        ${helpButton("werkzeuge")}
      </div>
      <div class='flex items-center gap-2 flex-wrap'>
        <span class='text-xs text-slate-500'>${toolRefreshStatusText}</span>
        <button class='px-3 py-2 rounded bg-blue-700 text-white' onclick='refreshToolPageData()'>Werkzeuge aktualisieren</button>
      </div>
    </div>
    <p class='text-sm text-slate-600 mb-2'>Werkzeugbestand, Stammdaten, Bestellungen und Journal sind getrennt untereinander dargestellt.</p>
    ${toolsLoadingBanner}

    ${
      showStockSection
        ? `<div id='toolManagementStockSection' class='border-2 border-slate-300 rounded-xl p-3 w-full max-w-none'>
      <div class='flex items-center justify-between gap-3 flex-wrap mb-3'>
        <div>
          <h3 class='text-lg font-bold'>Werkzeugbestand</h3>
          <span class='text-sm text-slate-500'>${tools.length} von ${state.tools.length} Werkzeugen angezeigt</span>
        </div>
        ${
          isAdmin
            ? `<button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='openCreateToolModal()'>Neues Werkzeug erfassen</button>`
            : ""
        }
      </div>

      <div class='border rounded-lg bg-slate-50 p-3 mb-3'>
        <div class='grid lg:grid-cols-[2fr,1fr,1fr,1fr,1fr,1fr] md:grid-cols-3 gap-3 mb-3'>
          <input id='toolSearch' class='border rounded p-2 md:col-span-2 lg:col-span-1' placeholder='Suche nach T, Bezeichnung, Artikel, Aufnahme, Hersteller...' value='${escapeHtml(filters.search || "")}' />
          <select id='toolFilterLabel' class='border rounded p-2'>
            <option value='' ${filterLabel ? "" : "selected"}>Alle Bezeichnungen</option>${filterLabelOptions}
          </select>
          <input id='toolFilterT' class='border rounded p-2' placeholder='T-Nummer' value='${escapeHtml(filterT)}' />
          <input id='toolFilterD' class='border rounded p-2' placeholder='Durchmesser' value='${escapeHtml(filterD)}' />
          <select id='toolFilterHolder' class='border rounded p-2'>
            <option value='' ${filterHolder ? "" : "selected"}>Alle Aufnahmen</option>
            <option value='HSK 100' ${filterHolder === "HSK 100" ? "selected" : ""}>HSK 100</option>
            <option value='HSK 63' ${filterHolder === "HSK 63" ? "selected" : ""}>HSK 63</option>
          </select>
          <select id='toolFilterImageStatus' class='border rounded p-2'>
            <option value='' ${imageStatus ? "" : "selected"}>Alle Bildstatus</option>
            <option value='withPath' ${imageStatus === "withPath" ? "selected" : ""}>Mit Bildpfad</option>
            <option value='withoutPath' ${imageStatus === "withoutPath" ? "selected" : ""}>Ohne Bildpfad</option>
          </select>
        </div>
        <div class='flex gap-2 flex-wrap'>
          <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='applyToolFilters()'>Filter anwenden</button>
          <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='resetToolFilters()'>Filter zurücksetzen</button>
        </div>
        <p class='text-xs text-slate-500 mt-2'>Bildstatus prüft nur, ob aus T-Nummer und Aufnahme ein Bildpfad erstellt werden kann.</p>
      </div>

      ${
        isAdmin
          ? `<div class='flex gap-2 flex-wrap mb-3'>
              <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolLabel()'>Bezeichnung hinzufügen</button>
              <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolManufacturer()'>Hersteller hinzufügen</button>
              <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolMaterial()'>Schneidwerkstoff hinzufügen</button>
              <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='openToolMasterDataModal()'>Stammdaten anzeigen</button>
            </div>`
          : ""
      }

      <div id='toolStockTableScroll' class='overflow-x-visible overflow-y-auto max-h-[35vh] border rounded-lg'>
        <table class='w-full table-auto text-sm'>
          <thead class='bg-slate-100 sticky top-0'>
            ${stockTableHeader}
          </thead>
          <tbody>${toolRows || `<tr><td class="p-2" colspan="${isAdmin ? 13 : 10}">Keine Werkzeuge.</td></tr>`}</tbody>
        </table>
      </div>
    </div>`
        : ""
    }

    ${
      showAdminActionsShell
        ? `<div id='toolManagementAdminActionsSection' class='border-2 border-slate-300 rounded-xl p-3 space-y-3'>
            <div class='flex items-center justify-between gap-2 flex-wrap'>
              <h3 class='text-lg font-bold'>${toolSectionShellTitle}</h3>
              ${
                INVENTORY_MODE_ENABLED
                  ? "<button class='px-3 py-2 rounded bg-indigo-700 text-white text-sm' onclick='openInventoryMode()'>Inventurmodus starten</button>"
                  : ""
              }
            </div>

            ${
              showReorderSection
                ? `<div id='toolManagementReorderSection' class='border rounded p-3 bg-white overflow-auto'>
              <h4 class='font-semibold mb-2'>Nachbestellen prüfen</h4>
              <table class='w-full text-sm'>
                <thead class='bg-slate-100'>
                  <tr>
                    <th class='p-2 text-left'>T</th>
                    <th class='p-2 text-left'>Bezeichnung</th>
                    <th class='p-2 text-left'>Größe</th>
                    <th class='p-2 text-left'>Bestand</th>
                    <th class='p-2 text-left'>Mindestbestand</th>
                    <th class='p-2 text-left'>Aufnahme</th>
                    <th class='p-2 text-left'>Fach</th>
                    <th class='p-2 text-left'>Status</th>
                    <th class='p-2'></th>
                  </tr>
                </thead>
                <tbody>${todoRows || '<tr><td class="p-2" colspan="9">Keine Werkzeuge mit erreichtem Mindestbestand.</td></tr>'}</tbody>
              </table>
            </div>

            <div class='border rounded p-3 bg-white'>
              <h4 class='font-semibold mb-2'>Bestellt</h4>
              <div class='space-y-3'>
                ${orderedCards || '<div class="text-sm text-slate-500">Keine bestellten Werkzeuge.</div>'}
              </div>
            </div>

            <div class='border rounded p-3 bg-slate-50'>
              <h4 class='font-semibold mb-2'>Bestellliste nach Hersteller</h4>
              <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick='openOrderListPopup()'>Bestellliste anzeigen</button>
            </div>

            <div class='border rounded p-3 bg-slate-50'>
              <h4 class='font-semibold mb-2'>Vorschläge für optimale Bestellmenge</h4>
              <table class='w-full text-sm'>
                <thead class='bg-slate-100'>
                  <tr>
                    <th class='p-2 text-left'>Bezeichnung</th>
                    <th class='p-2 text-left'>Größe</th>
                    <th class='p-2 text-left'>Werkstoff</th>
                    <th class='p-2 text-left'>Aktuell optimal</th>
                    <th class='p-2 text-left'>Vorschlag</th>
                    <th class='p-2'></th>
                  </tr>
                </thead>
                <tbody>${suggestionRows || '<tr><td class="p-2" colspan="6">Noch keine aussagekräftigen Vorschläge vorhanden.</td></tr>'}</tbody>
              </table>
            </div>`
                : ""
            }

            ${
              showJournalSection
                ? `<div id='toolManagementJournalSection' class='border rounded p-3 bg-white'>
              <h4 class='font-semibold mb-2'>Schichtjournal – Werkzeugwechsel</h4>
              ${journalFilterBar}
              <div class='overflow-auto max-h-[25vh]'>
                <table class='w-full text-sm'>
                  <thead class='bg-slate-100 sticky top-0'>
                    <tr>
                      <th class='p-2 text-left'>Zeit</th>
                      <th class='p-2 text-left'>Benutzer</th>
                      <th class='p-2 text-left'>Werkzeug</th>
                      <th class='p-2 text-left'>Aktion</th>
                      <th class='p-2 text-left'>Menge</th>
                      <th class='p-2 text-left'>Bestand vorher/nachher</th>
                    </tr>
                  </thead>
                  <tbody>${journalEntries.length === 0 ? '<tr><td class="p-2" colspan="6">Keine Einträge.</td></tr>' : journalRows}</tbody>
                </table>
              </div>
            </div>`
                : ""
            }
            ${showToolStatisticsSection ? toolStatisticsSection : ""}
            ${showStorageOverviewSection ? storageOverviewSection : ""}
          </div>`
        : ""
    }

    ${
      showEmployeeExtraSections
        ? `<div class='space-y-3'>
            ${
              showJournalSection
                ? `<div id='toolManagementJournalSection' class='border-2 border-slate-300 rounded-xl p-3'>
              <h3 class='text-lg font-bold mb-2'>Schichtjournal – Werkzeugwechsel</h3>
              ${journalFilterBar}
              <div class='overflow-auto max-h-[25vh]'>
                <table class='w-full text-sm'>
                  <thead class='bg-slate-100 sticky top-0'>
                    <tr>
                      <th class='p-2 text-left'>Zeit</th>
                      <th class='p-2 text-left'>Benutzer</th>
                      <th class='p-2 text-left'>Werkzeug</th>
                      <th class='p-2 text-left'>Aktion</th>
                      <th class='p-2 text-left'>Menge</th>
                      <th class='p-2 text-left'>Bestand vorher/nachher</th>
                    </tr>
                  </thead>
                  <tbody>${journalEntries.length === 0 ? '<tr><td class="p-2" colspan="6">Keine Einträge.</td></tr>' : journalRows}</tbody>
                </table>
              </div>
            </div>`
                : ""
            }
            ${showToolStatisticsSection ? toolStatisticsSection : ""}
            ${showStorageOverviewSection ? storageOverviewSection : ""}
          </div>`
        : ""
    }

    ${showReorderSection ? orderListPopup : ""}
  </div>`;
}

function toggleInsertToolFields() {
  toggleInsertToolFieldsById(
    "toolInsertTool",
    "toolInsertEdges",
    "toolInsertRadius",
  );
}

function toggleInsertToolFieldsById(checkboxId, edgesId, radiusId = "") {
  const checkbox = document.getElementById(checkboxId);
  const edges = document.getElementById(edgesId);
  const radius = radiusId ? document.getElementById(radiusId) : null;
  const radiusWrap = radiusId ? document.getElementById(`${radiusId}Wrap`) : null;
  if (!checkbox || !edges) return;

  const checked = !!checkbox.checked;
  edges.disabled = !checked;
  if (!checked) edges.value = "";

  if (radius) {
    radius.disabled = !checked;
    if (!checked) radius.value = "";
  }
  if (radiusWrap) radiusWrap.style.display = checked ? "" : "none";
}

function taskAssignee(task) {
  return task.assignee || task.assignedTo || task.assigned_to || "";
}

function taskDueDate(task) {
  return task.dueDate || task.due_date || "";
}

function taskCompletedAt(task) {
  return task.completedAt || task.completed_at || null;
}

function normalizeNameForCompare(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function currentUserNameForCompare() {
  return normalizeNameForCompare(
    currentUser?.name ||
      currentEmployeeRecord?.display_name ||
      currentEmployeeRecord?.name ||
      "",
  );
}

function isTaskAssignedToCurrentUser(task) {
  return (
    normalizeNameForCompare(taskAssignee(task)) === currentUserNameForCompare()
  );
}

function renderTodo() {
  const nowIso = new Date().toISOString();
  const isAdmin = currentUser?.role === "admin";
  const allTasks = Array.isArray(state.tasks) ? state.tasks : [];
  console.log("renderTodo allTasks:", allTasks);

  const visibleTasks = isAdmin
    ? allTasks
    : allTasks.filter((task) => isTaskAssignedToCurrentUser(task));

  console.log("To-Do Render Debug:", {
    user: currentUser,
    employeeRecord: currentEmployeeRecord,
    tasks: state.tasks,
    visibleCount: visibleTasks?.length,
  });

  const openTasks = visibleTasks.filter((task) => task.status !== "done");
  const doneTasks = visibleTasks.filter((task) => task.status === "done");

  const openRows = openTasks
    .map((task) => {
      const dueDate = taskDueDate(task);
      const completedAt = taskCompletedAt(task);
      const overdue = dueDate && nowIso > `${dueDate}T23:59:59`;
      return `<tr class='border-b ${overdue ? "bg-rose-50" : ""}'>
      <td class='p-2'>${task.title}</td><td class='p-2'>${taskAssignee(task)}</td><td class='p-2'>${dueDate ? formatDateDisplay(dueDate) : "-"}</td>
      <td class='p-2'>${completedAt ? "Erledigt" : "Offen"}</td>
      <td class='p-2'>
        ${
          isAdmin
            ? `<button class='px-2 py-1 rounded bg-slate-900 text-white mr-1' onclick="deleteTask('${task.id}')">Löschen</button>
          <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="reassignTaskPrompt('${task.id}')">Neu zuweisen</button>`
            : `<button class='px-2 py-1 rounded bg-emerald-600 text-white' onclick="completeTask('${task.id}')">Erledigt</button>`
        }
      </td>
    </tr>`;
    })
    .join("");

  const doneRows = doneTasks
    .map((task) => {
      const dueDate = taskDueDate(task);
      const completedAt = taskCompletedAt(task);
      return `<tr class='border-b bg-emerald-50'>
      <td class='p-2'>✅ ${task.title}</td><td class='p-2'>${taskAssignee(task)}</td><td class='p-2'>${dueDate ? formatDateDisplay(dueDate) : "-"}</td><td class='p-2'>${completedAt || "-"}</td>
      <td class='p-2'>${isAdmin ? `<button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="deleteTask('${task.id}')">Löschen</button>` : "-"}</td>
    </tr>`;
    })
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>To-Do</h2>
    ${isAdmin ? renderTaskCreateBox() : ""}
    <h3 class='font-semibold mt-3 mb-2'>Offene Aufgaben</h3>
    <div class='overflow-auto max-h-[45vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Aufgabe</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Frist</th><th class='p-2 text-left'>Status</th><th class='p-2'></th></tr></thead><tbody>${openRows || '<tr><td class="p-2" colspan="5">Keine offenen Aufgaben.</td></tr>'}</tbody></table>
    </div>
    <h3 class='font-semibold mt-4 mb-2'>Erledigte Aufgaben</h3>
    <div class='overflow-auto max-h-[25vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Aufgabe</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Frist</th><th class='p-2 text-left'>Erledigt am</th><th class='p-2'></th></tr></thead><tbody>${doneRows || '<tr><td class="p-2" colspan="5">Noch keine erledigten Aufgaben.</td></tr>'}</tbody></table>
    </div>
  </div>`;
}

function renderTaskCreateBox() {
  const options = activeUsers()
    .map((u) => `<option value='${u.name}'>${u.name}</option>`)
    .join("");
  return `<div class='border rounded-lg p-3 bg-slate-50 mb-3'>
    <h3 class='font-semibold mb-2'>Neue Aufgabe erstellen</h3>
    <div class='grid md:grid-cols-4 gap-2'>
      <input id='taskTitle' class='border rounded p-2' placeholder='Aufgabe'/>
      <select id='taskAssignee' class='border rounded p-2'>${options}</select>
      <input id='taskDeadline' type='date' class='border rounded p-2'/>
      <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick='createTask()'>Speichern</button>
    </div>
  </div>`;
}

async function createTask() {
  if (!supabaseReady) return;

  const title = document.getElementById("taskTitle")?.value?.trim();
  const assignee = document.getElementById("taskAssignee")?.value;
  const deadline = document.getElementById("taskDeadline")?.value;
  if (!title || !assignee)
    return alert("Bitte Aufgabe und Mitarbeiter angeben.");

  const payload = {
    title,
    description: "",
    assigned_to: assignee,
    due_date: deadline || null,
    status: "open",
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("planner_tasks")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern der Aufgabe:", error);
    return alert(`Aufgabe konnte nicht gespeichert werden: ${error.message}`);
  }

  state.tasks.push(normalizeTaskFromDb(inserted));
  persist();
  render();
}

async function completeTask(taskId) {
  if (!supabaseReady) return;

  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;

  const completedAt = new Date().toISOString();

  const { data: updated, error } = await supabaseClient
    .from("planner_tasks")
    .update({
      status: "done",
      completed_at: completedAt,
      updated_at: completedAt,
    })
    .eq("id", taskId)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Abschließen der Aufgabe:", error);
    return alert(`Aufgabe konnte nicht abgeschlossen werden: ${error.message}`);
  }

  task.status = "done";
  task.completedAt = updated?.completed_at || completedAt;
  persist();
  render();
}

async function deleteTask(taskId) {
  if (!supabaseReady) return;

  const { error } = await supabaseClient
    .from("planner_tasks")
    .delete()
    .eq("id", taskId);

  if (error) {
    console.error("Fehler beim Löschen der Aufgabe:", error);
    return alert(`Aufgabe konnte nicht gelöscht werden: ${error.message}`);
  }

  state.tasks = state.tasks.filter((t) => t.id !== taskId);
  persist();
  render();
}

async function reassignTaskPrompt(taskId) {
  if (!supabaseReady) return;

  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;

  const currentAssignee = taskAssignee(task);
  const options = activeUsers()
    .map((user) => {
      const selected = user.name === currentAssignee ? "selected" : "";
      return `<option value="${escapeHtml(user.name)}" ${selected}>${escapeHtml(user.name)}</option>`;
    })
    .join("");

  const host = getModalHost();
  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
      <h3 class="text-lg font-bold mb-3">Aufgabe neu zuweisen</h3>
      <p class="text-sm text-slate-600 mb-3">${escapeHtml(task.title)}</p>
      <select id="taskReassignSelect" class="border rounded p-2 w-full mb-4">${options}</select>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeTaskReassignModal()">Abbrechen</button>
        <button class="px-3 py-1 rounded bg-slate-900 text-white" onclick="saveTaskReassignment('${taskId}')">Speichern</button>
      </div>
    </div>
  </div>`;
}

function closeTaskReassignModal() {
  getModalHost().innerHTML = "";
}

async function saveTaskReassignment(taskId) {
  if (!supabaseReady) return;

  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;

  const selectedName = document.getElementById("taskReassignSelect")?.value;
  if (!selectedName) return alert("Bitte Mitarbeiter auswählen.");

  const updatedAt = new Date().toISOString();

  const { error } = await supabaseClient
    .from("planner_tasks")
    .update({
      assigned_to: selectedName,
      updated_at: updatedAt,
    })
    .eq("id", taskId);

  if (error) {
    console.error("Fehler beim Neuzuweisen der Aufgabe:", error);
    return alert(
      `Aufgabe konnte nicht neu zugewiesen werden: ${error.message}`,
    );
  }

  task.assignee = selectedName;
  task.assignedTo = selectedName;
  closeTaskReassignModal();
  persist();
  render();
}

function maybeTaskReminder() {
  if (!currentUser || currentUser.role !== "employee") return;
  const now = new Date();
  const today = isoDate(now);
  const hourKey = `${today}:${now.getHours()}:${currentUser.name}`;
  if (state[`_taskReminder_${hourKey}`]) return;
  const dueToday = (state.tasks || []).filter(
    (task) =>
      isTaskAssignedToCurrentUser(task) &&
      task.status !== "done" &&
      taskDueDate(task) === today,
  );
  if (!dueToday.length) return;
  alert(`Erinnerung: ${dueToday.length} Aufgabe(n) heute fällig.`);
  state[`_taskReminder_${hourKey}`] = true;
  persist();
}

function renderConflicts() {
  const rows = Object.entries(state.conflicts)
    .map(
      ([id, c]) => `<tr class='border-b'>
    <td class='p-2'>${formatDateWithWeekday(c.date)}</td><td class='p-2'>${c.user}</td><td class='p-2'>${c.text}</td>
    <td class='p-2'>${c.resolved ? "Ja" : "Nein"}</td>
    <td class='p-2'><button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="setConflictResolved('${id}', ${c.resolved ? "false" : "true"})">${c.resolved ? "Auf Nein" : "Als gelöst markieren"}</button></td>
  </tr>`,
    )
    .join("");
  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Konflikte</h2>
    <div class='overflow-auto max-h-[70vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Konflikt</th><th class='p-2 text-left'>Gelöst</th><th class='p-2'></th></tr></thead>
      <tbody>${rows || '<tr><td class="p-2" colspan="5">Keine Konflikte.</td></tr>'}</tbody></table>
    </div>
  </div>`;
}

function setConflictResolved(id, resolved) {
  if (!state.conflicts[id]) return;
  state.conflicts[id].resolved = resolved;
  persist();
  render();
}

function getPeriodRange(period) {
  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);
  if (period === "week") {
    start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
    end.setDate(start.getDate() + 6);
  } else if (period === "month") {
    start.setDate(1);
    end.setMonth(start.getMonth() + 1, 0);
  } else {
    start.setMonth(0, 1);
    end.setMonth(11, 31);
  }
  return { from: isoDate(start), to: isoDate(end), start, end };
}

function plannedMannedHoursForShift(shift) {
  if (shift.id.includes("-mf-")) return 6;
  if (shift.id.includes("-sa-0")) return 6;
  if (shift.id.includes("-sa-1")) return 6;
  if (shift.id.includes("-su-0")) return 6;
  if (shift.id.includes("-su-1")) return 6;
  return shiftHours(shift.start, shift.end);
}

function parseDurationHours(text) {
  if (!text || !text.includes(":")) return 0;
  const [h, m] = text.split(":").map(Number);
  if (Number.isNaN(h) || Number.isNaN(m)) return 0;
  return h + m / 60;
}

function shiftDateRange(shift) {
  const start = new Date(`${shift.date}T${shift.start}:00`);
  const end = new Date(`${shift.date}T${shift.end}:00`);
  if (end <= start) end.setDate(end.getDate() + 1);
  return { start, end };
}

function getShiftById(shiftId) {
  return generateThreeMonths().find((s) => s.id === shiftId) || null;
}

function isCurrentOrFutureShift(shift) {
  const now = new Date();
  const { start, end } = shiftDateRange(shift);
  return start >= now || (now >= start && now <= end);
}

function maxUnmannedHoursForShift(shiftId) {
  const all = generateThreeMonths()
    .filter((s) => s.assigned)
    .sort((a, b) => {
      const ra = shiftDateRange(a).start;
      const rb = shiftDateRange(b).start;
      return ra - rb;
    });
  const current = all.find((s) => s.id === shiftId);
  if (!current) return 0;
  const currentEnd = shiftDateRange(current).end;
  const next = all.find((s) => shiftDateRange(s).start > currentEnd);
  if (!next) return 0;
  const nextStart = shiftDateRange(next).start;
  const diffHours = Math.max(
    0,
    (nextStart.getTime() - currentEnd.getTime()) / (1000 * 60 * 60),
  );
  return diffHours;
}

function overlapHours(rangeStart, rangeEnd, shiftStart, shiftEnd) {
  const start = Math.max(rangeStart.getTime(), shiftStart.getTime());
  const end = Math.min(rangeEnd.getTime(), shiftEnd.getTime());
  if (end <= start) return 0;
  return (end - start) / (1000 * 60 * 60);
}

function computeStats(period = "week") {
  const { from, to, start, end } = getPeriodRange(period);
  const target = 154;
  const now = new Date();
  const shifts = generateThreeMonths().filter(
    (s) => s.date >= from && s.date <= to && s.assigned,
  );

  const plannedTotal = shifts.reduce(
    (acc, s) => acc + shiftHours(s.start, s.end),
    0,
  );
  const plannedManned = shifts.reduce(
    (acc, s) => acc + plannedMannedHoursForShift(s),
    0,
  );
  const plannedUnmanned = Math.max(0, plannedTotal - plannedManned);

  const elapsedPlanned = shifts.reduce((acc, s) => {
    const { start: sStart, end: sEnd } = shiftDateRange(s);
    return acc + overlapHours(start, now < end ? now : end, sStart, sEnd);
  }, 0);

  const recordedUnmanned = Object.entries(state.unmanned)
    .filter(([shiftId]) => {
      const date = shiftId.slice(0, 10);
      return date >= from && date <= to;
    })
    .reduce((acc, [shiftId, timeText]) => {
      const entered = parseDurationHours(timeText);
      const allowed = maxUnmannedHoursForShift(shiftId);
      return acc + Math.min(entered, allowed);
    }, 0);

  const downtime =
    Object.entries(state.machineDowntime)
      .filter(([k]) => {
        const date = k.split(":")[0];
        return date >= from && date <= to && date <= isoDate(now);
      })
      .reduce((acc, [, v]) => acc + (v.minutes || 0), 0) / 60;

  const istHours = Math.max(0, elapsedPlanned - downtime);
  const deviationPct = elapsedPlanned
    ? (Math.abs(elapsedPlanned - istHours) / elapsedPlanned) * 100
    : 0;
  const downtimePct = elapsedPlanned ? (downtime / elapsedPlanned) * 100 : 0;
  const targetPct = target ? (istHours / target) * 100 : 0;

  return {
    target,
    period,
    plannedTotal: round1(plannedTotal),
    elapsedPlanned: round1(elapsedPlanned),
    plannedManned: round1(plannedManned),
    plannedUnmanned: round1(plannedUnmanned),
    recordedUnmanned: round1(recordedUnmanned),
    downtime: round1(downtime),
    istHours: round1(istHours),
    deviationPct: round1(deviationPct),
    downtimePct: round1(downtimePct),
    targetPct: round1(targetPct),
  };
}

function round1(value) {
  return Math.round(value * 10) / 10;
}

function shiftHours(start, end) {
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let diff = eh * 60 + em - (sh * 60 + sm);
  if (diff <= 0) diff += 24 * 60;
  return diff / 60;
}

function renderStats() {
  const stats = computeStats(statsViewPeriod);
  const color =
    stats.deviationPct <= 8
      ? "text-emerald-600"
      : stats.deviationPct <= 12
        ? "text-orange-500"
        : "text-rose-600";
  const barPlan = Math.min(
    100,
    stats.target ? (stats.plannedTotal / stats.target) * 100 : 0,
  );
  const barIst = Math.min(
    100,
    stats.target ? (stats.istHours / stats.target) * 100 : 0,
  );
  const barStillstand = Math.min(
    100,
    stats.target ? (stats.downtime / stats.target) * 100 : 0,
  );

  return `<div class='bg-white rounded-xl shadow p-4'>
    <div class='flex items-center justify-between mb-3'>
      <h2 class='text-lg font-semibold'>Laufzeitstatistik (${stats.period === "week" ? "Woche" : stats.period === "month" ? "Monat" : "Jahr"})</h2>
      <div class='flex items-center gap-2'>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "week" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('week')">Woche</button>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "month" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('month')">Monat</button>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "year" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('year')">Jahr</button>
        ${helpButton("statistik")}
      </div>
    </div>
    <div class='grid md:grid-cols-4 gap-3 text-center mb-4'>
      <div class='p-3 rounded bg-slate-100'><div class='text-sm'>Marker</div><div class='text-2xl font-bold'>${stats.target} h</div></div>
      <div class='p-3 rounded bg-blue-100'><div class='text-sm'>Geplante Stunden</div><div class='text-2xl font-bold'>${stats.plannedTotal} h</div><div class='text-xs text-slate-500'>${stats.targetPct}% vom Marker</div></div>
      <div class='p-3 rounded bg-emerald-100'><div class='text-sm'>Ist-Stunden (bis jetzt)</div><div class='text-2xl font-bold'>${stats.istHours} h</div><div class='text-xs text-slate-500'>Stillstand ${stats.downtimePct}%</div></div>
      <div class='p-3 rounded bg-white border'><div class='text-sm'>Abweichung Plan/Ist</div><div class='text-2xl font-bold ${color}'>${stats.deviationPct}%</div></div>
    </div>
    <div class='grid md:grid-cols-2 gap-4 mb-4'>
      <div class='p-3 rounded border bg-slate-50'>
        <h3 class='font-semibold mb-2'>Bemannt / Mannlos</h3>
        <div class='text-sm'>Bemannt (Plan): <b>${stats.plannedManned} h</b></div>
        <div class='text-sm'>Mannlos (Plan): <b>${stats.plannedUnmanned} h</b></div>
        <div class='text-sm'>Mannlos (eingetragen): <b>${stats.recordedUnmanned} h</b></div>
      </div>
      <div class='p-3 rounded border bg-slate-50'>
        <h3 class='font-semibold mb-2'>Grafik (h / Marker 154h)</h3>
        <div class='mb-2 text-xs'>Plan: ${stats.plannedTotal} h</div>
        <div class='w-full bg-slate-200 rounded h-4 mb-2'><div class='bg-blue-500 h-4 rounded' style='width:${barPlan}%'></div></div>
        <div class='mb-2 text-xs'>Ist: ${stats.istHours} h</div>
        <div class='w-full bg-slate-200 rounded h-4 mb-2'><div class='bg-emerald-500 h-4 rounded' style='width:${barIst}%'></div></div>
        <div class='mb-2 text-xs'>Stillstand: ${stats.downtime} h</div>
        <div class='w-full bg-slate-200 rounded h-4'><div class='bg-rose-500 h-4 rounded' style='width:${barStillstand}%'></div></div>
      </div>
    </div>
  </div>`;
}

function renderAvailability() {
  const shifts = generateThreeMonths().filter((s) => s.options.length > 1);
  const rows = shifts
    .slice(0, 80)
    .map((s) => {
      const key = `${s.id}:${currentUser.name}`;
      const val = state.availability[key] || "";
      return `<tr class='border-b'><td class='p-2'>${formatDateWithWeekday(s.date)}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.options.join(" / ")}</td>
      <td class='p-2'>
        <select class='border rounded p-1' onchange="setAvailability('${s.id}', this.value)">
          <option value="" ${val === "" ? "selected" : ""}>-</option>
          <option value="yes" ${val === "yes" ? "selected" : ""}>Kann</option>
          <option value="no" ${val === "no" ? "selected" : ""}>Kann nicht</option>
        </select>
      </td></tr>`;
    })
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Unklare Schichten (nur für Admin sichtbar in Planung)</h2>
    <p class='text-sm text-slate-500 mb-3'>Du kannst hier vorab eintragen, ob du bei Entweder/Oder-Schichten könntest.</p>
    <div class='overflow-auto max-h-[70vh]'><table class='w-full text-sm'><thead class='bg-slate-100'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Option</th><th class='p-2 text-left'>Dein Status</th></tr></thead><tbody>${rows}</tbody></table></div>
  </div>`;
}

function setAvailability(shiftId, value) {
  state.availability[`${shiftId}:${currentUser.name}`] = value;
  persist();
}

function renderClosure() {
  const myShifts = generateThreeMonths()
    .filter((s) => s.assigned === currentUser.name)
    .slice(0, 20);
  const rows = myShifts
    .map((s) => {
      const done = state.checklists[s.id] ? "✅" : "⏳";
      return `<tr class='border-b'><td class='p-2'>${formatDateWithWeekday(s.date)}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.start}–${s.end}</td><td class='p-2'>${done}</td>
      <td class='p-2'><button class='px-2 py-1 rounded bg-indigo-700 text-white' onclick="openChecklist('${s.id}')">Checklist starten</button></td></tr>`;
    })
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4 space-y-3'>
    <h2 class='text-lg font-semibold'>Schichtabschluss</h2>
    <p class='text-sm text-slate-500'>Im Livebetrieb soll dieser Dialog 15 Minuten vor Schichtende erscheinen. Im MVP startest du ihn manuell.</p>
    <table class='w-full text-sm'><thead class='bg-slate-100'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Zeit</th><th class='p-2'>Status</th><th></th></tr></thead><tbody>${rows}</tbody></table>
  </div>`;
}

function openChecklist(shiftId) {
  const hourOptions = Array.from(
    { length: 24 },
    (_, i) => `<option>${String(i).padStart(2, "0")}</option>`,
  ).join("");
  const minOptions = Array.from(
    { length: 60 },
    (_, i) => `<option>${String(i).padStart(2, "0")}</option>`,
  ).join("");

  const wrapper = document.createElement("div");
  wrapper.className =
    "fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50";
  wrapper.innerHTML = `<div class='bg-white rounded-xl p-4 max-w-lg w-full space-y-3'>
    <h3 class='text-lg font-semibold'>Pflicht-Checkliste</h3>
    ${checkbox("c1", "Arbeitsplatz ordentlich")}
    ${checkbox("c2", "Informationen für die nächste Schicht vorhanden")}
    ${checkbox("c3", "Material versorgt und Lagerorte eingetragen")}
    ${checkbox("c4", "Alle Aufträge gebucht")}
    <div class='border-t pt-3'>
      <p class='text-sm font-medium mb-2'>Zeit für Mannlosbetrieb erfassen</p>
      <div class='flex gap-2 items-center'>
        <select id='uh'>${hourOptions}</select> : <select id='um'>${minOptions}</select>
      </div>
    </div>
    <div class='flex justify-end gap-2'>
      <button class='px-3 py-2 rounded bg-slate-200' id='cancelBtn'>Abbrechen</button>
      <button class='px-3 py-2 rounded bg-slate-900 text-white' id='saveBtn'>Speichern</button>
    </div>
  </div>`;
  document.body.appendChild(wrapper);

  wrapper.querySelector("#cancelBtn").onclick = () => wrapper.remove();
  wrapper.querySelector("#saveBtn").onclick = () => {
    const checks = ["c1", "c2", "c3", "c4"].map(
      (id) => wrapper.querySelector(`#${id}`).checked,
    );
    if (checks.some((v) => !v)) {
      alert("Bitte alle 4 Punkte bestätigen.");
      return;
    }
    state.checklists[shiftId] = {
      by: currentUser.name,
      at: new Date().toISOString(),
      checks,
    };
    state.shiftEndChecks[shiftId] = {
      by: currentUser.name,
      at: new Date().toISOString(),
      checks,
    };
    const h = wrapper.querySelector("#uh").value;
    const m = wrapper.querySelector("#um").value;
    state.unmanned[shiftId] = `${h}:${m}`;
    persist();
    wrapper.remove();
    render();
  };
}

function getCurrentShiftForUser(name) {
  const now = new Date();
  const today = isoDate(now);
  const shifts = generateThreeMonths().filter(
    (s) => s.assigned === name && s.date === today,
  );
  return (
    shifts.find((s) => {
      const [sh, sm] = s.start.split(":").map(Number);
      const [eh, em] = s.end.split(":").map(Number);
      const start = new Date(now);
      start.setHours(sh, sm, 0, 0);
      const end = new Date(now);
      end.setHours(eh, em, 0, 0);
      if (end <= start) end.setDate(end.getDate() + 1);
      return now >= start && now <= end;
    }) || null
  );
}

function maybeShowMachinePrompt() {
  if (!currentUser || currentUser.role !== "employee") return;
  const today = isoDate(new Date());
  const key = `${today}:${currentUser.name}`;
  if (state.machinePromptSeen[key]) {
    renderDowntimeWidget();
    return;
  }
  state.machinePromptSeen[key] = true;
  const running = confirm("Läuft die Maschine zu Schichtbeginn?");
  if (!running) {
    state.machineDowntime[key] = {
      startAt: new Date().toISOString(),
      minutes: 0,
      running: true,
    };
  }
  persist();
  renderDowntimeWidget();
}

function renderDowntimeWidget() {
  const today = isoDate(new Date());
  const key = `${today}:${currentUser.name}`;
  const downtime = state.machineDowntime[key];
  let box = document.getElementById("downtimeWidget");
  if (!box) {
    box = document.createElement("div");
    box.id = "downtimeWidget";
    box.className = "fixed top-4 right-4 z-50";
    document.body.appendChild(box);
  }
  if (!downtime || !downtime.running) {
    box.innerHTML = "";
    return;
  }
  box.innerHTML = `<button onclick="stopDowntimeTimer()" class="px-4 py-3 rounded-full bg-emerald-600 text-white shadow-lg flex items-center gap-2">
    <span class="w-4 h-4 rounded-full bg-green-300 inline-block"></span>
    Stillstand läuft – Stoppen
  </button>`;
}

function stopDowntimeTimer() {
  const today = isoDate(new Date());
  const key = `${today}:${currentUser.name}`;
  const downtime = state.machineDowntime[key];
  if (!downtime?.running) return;
  const start = new Date(downtime.startAt);
  const minutes = Math.max(
    1,
    Math.round((Date.now() - start.getTime()) / 60000),
  );
  state.machineDowntime[key] = {
    ...downtime,
    running: false,
    minutes: (downtime.minutes || 0) + minutes,
  };
  persist();
  render();
}

function maybeShowShiftEndChecklist() {
  if (!currentUser || currentUser.role !== "employee") return;
  const shift = getCurrentShiftForUser(currentUser.name);
  if (!shift) return;
  const [eh, em] = shift.end.split(":").map(Number);
  const end = new Date();
  end.setHours(eh, em, 0, 0);
  const trigger = new Date(end.getTime() - 15 * 60000);
  const now = new Date();
  if (now >= trigger && !state.shiftEndChecks[shift.id])
    openChecklist(shift.id);
}

function maybeShowShiftStartChecklist() {
  if (!currentUser || currentUser.role !== "employee") return;
  const shift = getCurrentShiftForUser(currentUser.name);
  if (!shift) return;
  const key = `${shift.id}:${currentUser.name}`;
  if (state.shiftStartChecks[key]) return;
  const endData = state.shiftEndChecks[shift.id];
  const answers = [
    "Arbeitsplatz ordentlich",
    "Informationen für die nächste Schicht vorhanden",
    "Material versorgt und Lagerorte eingetragen",
    "Alle Aufträge gebucht",
  ].map((txt) => ({ txt, ok: confirm(`Schichtstart-Check: ${txt}?`) }));
  state.shiftStartChecks[key] = answers;
  answers.forEach((a, idx) => {
    if (endData && endData.checks && endData.checks[idx] !== a.ok) {
      state.conflicts[`${key}:${idx}`] = {
        date: shift.date,
        user: currentUser.name,
        text: `${a.txt} weicht von Schichtende-Angabe ab`,
        resolved: false,
      };
    }
  });
  persist();
}

function checkbox(id, label) {
  return `<label class='flex items-center gap-2'><input id='${id}' type='checkbox' class='w-4 h-4'/> ${label}</label>`;
}

async function markAbsent(shiftId, date, userName) {
  if (!supabaseReady) return;

  const absenceKey = `${date}:${userName}`;

  const { error } = await supabaseClient.from("planner_absences").upsert(
    {
      absence_key: absenceKey,
      absence_date: date,
      user_name: userName,
      absence_type: "abwesend",
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "absence_key" },
  );

  if (error) {
    console.error("Fehler bei Abwesenheit:", error);
    return alert(
      `Abwesenheit konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  // lokal weiterführen (für UI)
  state.absences[absenceKey] = true;

  persist();
  render();
}

async function assignOptionalShift(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  if (!isPrimaryCoreAbsentForShift(shift)) {
    return alert(
      "Einteilung nicht möglich: Der A/B/C-Mitarbeiter ist nicht abwesend gemeldet.",
    );
  }

  const select = document.getElementById(`opt-${shiftId}`);
  if (!select) return;

  const name = select.value;
  if (!name) return alert("Bitte Springer auswählen.");

  if (!isSpringer(name)) {
    return alert("Für diese Funktion dürfen nur Springer eingeteilt werden.");
  }

  if (!canAssignUserToShift(name, shiftId)) return;

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: name,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler bei Springer-Einteilung:", error);
    return alert(
      `Einteilung konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = name;

  persist();
  render();
}

async function assignSuggestedSpringer(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  if (!isPrimaryCoreAbsentForShift(shift)) {
    return alert(
      "Einteilung nicht möglich: Der A/B/C-Mitarbeiter ist nicht abwesend gemeldet.",
    );
  }

  const suggested = getSuggestedSpringerForShift(shift);
  if (!suggested) {
    return alert("Kein verfügbarer Springer gefunden.");
  }

  if (!canAssignUserToShift(suggested, shiftId)) return;

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: suggested,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler bei automatischer Springer-Einteilung:", error);
    return alert(
      `Einteilung konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = suggested;

  persist();
  render();
}

async function approveSaturdayRequest(shiftId, user) {
  if (!supabaseReady) return;
  if (!canAssignUserToShift(user, shiftId)) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: user,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler beim Bestätigen der Samstags-Anfrage:", error);
    return alert(
      `Samstags-Anfrage konnte nicht übernommen werden: ${error.message}`,
    );
  }

  const [deleteRequest, deleteCancellation] = await Promise.all([
    supabaseClient
      .from("planner_saturday_requests")
      .delete()
      .eq("request_key", `${shiftId}:${user}`),
    supabaseClient
      .from("planner_shift_cancellations")
      .delete()
      .eq("shift_id", shiftId),
  ]);

  const firstError = deleteRequest.error || deleteCancellation.error;
  if (firstError) {
    console.error("Fehler beim Abschließen der Samstags-Anfrage:", firstError);
    return alert(
      `Einteilung gespeichert, aber Anfrage/Ausfall konnte nicht bereinigt werden: ${firstError.message}`,
    );
  }

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = user;
  delete state.saturdayEveningRequests[`${shiftId}:${user}`];
  persist();
  render();
}

injectHumbelDesignStyles();
render();
bootSupabase();

window.loginWithSupabase = loginWithSupabase;
window.logoutSupabase = logoutSupabase;
window.fillLogin = fillLogin;
window.loginAs = loginAs;
window.logout = logout;
window.openHelp = openHelp;
window.closeHelp = closeHelp;
window.openVersionLog = openVersionLog;
window.closeVersionLog = closeVersionLog;
window.openToolMasterDataModal = openToolMasterDataModal;
window.closeToolMasterDataModal = closeToolMasterDataModal;
window.openToolHistory = openToolHistory;
window.closeToolHistory = closeToolHistory;
window.openStorageLocationModal = openStorageLocationModal;
window.closeStorageLocationModal = closeStorageLocationModal;
window.openInventoryMode = openInventoryMode;
window.closeInventoryMode = closeInventoryMode;
window.processInventoryLocation = processInventoryLocation;
window.confirmInventoryResult = confirmInventoryResult;
window.openStorageQrModal = openStorageQrModal;
window.openStorageLocationByQr = openStorageLocationByQr;
window.openStorageQrFromInput = openStorageQrFromInput;
window.openStorageLocationFromInput = openStorageLocationFromInput;
window.printStorageQrFromInput = printStorageQrFromInput;
window.printStorageQrLabel = printStorageQrLabel;
window.openMoveToolModal = openMoveToolModal;
window.confirmMoveTool = confirmMoveTool;
window.applyStorageSearch = applyStorageSearch;
window.resetStorageSearch = resetStorageSearch;
window.setTab = setTab;
window.setDashboardSubTab = setDashboardSubTab;
window.setStatsView = setStatsView;
window.setPlanningSubTab = setPlanningSubTab;
window.setPersonnelManagementSubTab = setPersonnelManagementSubTab;
window.setProductionSubTab = setProductionSubTab;
window.setToolManagementSubTab = setToolManagementSubTab;
window.setStorageManagementSubTab = setStorageManagementSubTab;
window.setScannerSubTab = setScannerSubTab;
window.setAnalyticsSubTab = setAnalyticsSubTab;
window.setAdminSystemSubTab = setAdminSystemSubTab;
window.updateProductionOrderDepartmentPreview = updateProductionOrderDepartmentPreview;
window.createProductionOrder = createProductionOrder;
window.updateProductionOrderStatus = updateProductionOrderStatus;
window.pauseProductionOrder = pauseProductionOrder;
window.resumeProductionOrder = resumeProductionOrder;
window.completeProductionOrder = completeProductionOrder;
window.cancelProductionOrder = cancelProductionOrder;
window.selectProductionCounterMachine = selectProductionCounterMachine;
window.selectProductionCounterOrder = selectProductionCounterOrder;
window.resetProductionCounterSelection = resetProductionCounterSelection;
window.createProductionOrderStation = createProductionOrderStation;
window.saveProductionOrderStation = saveProductionOrderStation;
window.deleteProductionOrderStation = deleteProductionOrderStation;
window.addProductionOrderEmployee = addProductionOrderEmployee;
window.deactivateProductionOrderEmployee = deactivateProductionOrderEmployee;
window.reactivateProductionOrderEmployee = reactivateProductionOrderEmployee;
window.adjustProductionStationGoodQty = adjustProductionStationGoodQty;
window.adjustProductionStationAmount = adjustProductionStationAmount;
window.setProductionQaCauseSelection = setProductionQaCauseSelection;
window.closeProductionQaCauseModal = closeProductionQaCauseModal;
window.confirmProductionQaCauseModal = confirmProductionQaCauseModal;
window.openProductionCompletionDialog = openProductionCompletionDialog;
window.closeProductionCompletionDialog = closeProductionCompletionDialog;
window.toggleProductionOrderChecklistItem = toggleProductionOrderChecklistItem;
window.prepareProductionOrderChecklist = prepareProductionOrderChecklist;
window.completeProductionOrderWithChecklist = completeProductionOrderWithChecklist;
window.selectProductionProtocolOrder = selectProductionProtocolOrder;
window.toggleProductionProtocolSort = toggleProductionProtocolSort;
window.openProductionProtocolForOrder = openProductionProtocolForOrder;
window.markAbsent = markAbsent;
window.assignShift = assignShift;
window.cancelShift = cancelShift;
window.assignOptionalShift = assignOptionalShift;
window.approveSaturdayRequest = approveSaturdayRequest;
window.setAvailability = setAvailability;
window.openChecklist = openChecklist;
window.addVacation = addVacation;
window.addSickLeave = addSickLeave;
window.deleteVacation = deleteVacation;
window.deleteSickLeave = deleteSickLeave;
window.planAbsenceReplacement = planAbsenceReplacement;
window.clearAbsenceReplacementPlan = clearAbsenceReplacementPlan;
window.openAbsenceReplacementPlanner = openAbsenceReplacementPlanner;
window.closeReplacementPlanner = closeReplacementPlanner;
window.setReplacementPlannerChoice = setReplacementPlannerChoice;
window.toggleReplacementDay = toggleReplacementDay;
window.applyReplacementForSelectedDays = applyReplacementForSelectedDays;
window.applyReplacementForWeek = applyReplacementForWeek;
window.applyReplacementForAll = applyReplacementForAll;
window.addSwap = addSwap;
window.deleteSwap = deleteSwap;
window.resetActiveSwaps = resetActiveSwaps;
window.resetManualAssignments = resetManualAssignments;
window.resetPlanCurrentFuture = resetPlanCurrentFuture;
window.updateSlotAssignment = updateSlotAssignment;
window.queueEmployeeEdit = queueEmployeeEdit;
window.queueEmployeeActive = queueEmployeeActive;
window.queueSlotAssignment = queueSlotAssignment;
window.saveAllPersonnelChanges = saveAllPersonnelChanges;
window.addEmployee = addEmployee;
window.updateEmployee = updateEmployee;
window.fillEmployeeForm = fillEmployeeForm;
window.deactivateEmployee = deactivateEmployee;
window.activateEmployee = activateEmployee;
window.requestSaturdayEvening = requestSaturdayEvening;
window.setWeekendAvailability = setWeekendAvailability;
window.createTask = createTask;
window.completeTask = completeTask;
window.deleteTask = deleteTask;
window.reassignTaskPrompt = reassignTaskPrompt;
window.closeTaskReassignModal = closeTaskReassignModal;
window.saveTaskReassignment = saveTaskReassignment;
window.setConflictResolved = setConflictResolved;
window.stopDowntimeTimer = stopDowntimeTimer;
window.applyToolFilters = applyToolFilters;
window.resetToolFilters = resetToolFilters;
window.applyToolJournalFilters = applyToolJournalFilters;
window.resetToolJournalFilters = resetToolJournalFilters;
window.setToolStatisticsRange = setToolStatisticsRange;
window.forceLoadToolPageData = forceLoadToolPageData;
window.refreshToolPageData = refreshToolPageData;
window.startToolAutoRefresh = startToolAutoRefresh;
window.stopToolAutoRefresh = stopToolAutoRefresh;
window.startToolRealtimeSubscription = startToolRealtimeSubscription;
window.stopToolRealtimeSubscription = stopToolRealtimeSubscription;
window.bookToolChange = bookToolChange;
window.undoToolJournalEntry = undoToolJournalEntry;
window.openToolImagePopup = openToolImagePopup;
window.closeToolImagePopup = closeToolImagePopup;
window.openManualToolWithdraw = openManualToolWithdraw;
window.openManualToolRestock = openManualToolRestock;
window.openQrScannerPlaceholder = openQrScannerPlaceholder;
window.openQrScanner = openQrScanner;
window.stopQrScanner = stopQrScanner;
window.processScannedToolQr = processScannedToolQr;
window.confirmQrToolWithdraw = confirmQrToolWithdraw;
window.confirmQrToolRestock = confirmQrToolRestock;
window.updateScannerWithdrawPreview = updateScannerWithdrawPreview;
window.updateScannerRestockPreview = updateScannerRestockPreview;
window.openToolQrPopup = openToolQrPopup;
window.copyToolQrPayload = copyToolQrPayload;
window.printToolQrLabel = printToolQrLabel;
window.editTool = editTool;
window.deleteTool = deleteTool;
window.setToolOrderOverride = setToolOrderOverride;
window.setOrderStatsView = setOrderStatsView;
window.openOrderListPopup = openOrderListPopup;
window.closeOrderListPopup = closeOrderListPopup;
window.setOrderListManufacturer = setOrderListManufacturer;
window.applyOptimalQtySuggestion = applyOptimalQtySuggestion;
window.rejectOptimalQtySuggestion = rejectOptimalQtySuggestion;
window.toggleInsertToolFields = toggleInsertToolFields;
window.toggleInsertToolFieldsById = toggleInsertToolFieldsById;
window.markToolOrdered = markToolOrdered;
window.restockTool = restockTool;
window.createTool = createTool;
window.openCreateToolModal = openCreateToolModal;
window.addToolLabel = addToolLabel;
window.addToolManufacturer = addToolManufacturer;
window.updateEditToolTypeFields = updateEditToolTypeFields;
window.updateEditThreadPitchVisibility = updateEditThreadPitchVisibility;
window.addToolMaterial = addToolMaterial;
window.renameToolMaterial = renameToolMaterial;
window.deactivateToolMaterial = deactivateToolMaterial;
window.clearManualShift = clearManualShift;
window.assignSuggestedSpringer = assignSuggestedSpringer;
window.deleteManualAbsence = deleteManualAbsence;
