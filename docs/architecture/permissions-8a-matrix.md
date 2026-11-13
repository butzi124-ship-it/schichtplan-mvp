# Schritt 8A – Rollen- und Berechtigungsmatrix

**Konzept, nicht implementiert.** Ausgangs- und Quellenstand: [Bestandsaufnahme](modular-permissions-8a.md). B = bestätigte Fachregel; P = vorgeschlagene sichere Ausgestaltung; O = Detailentscheidung fehlt; E = bestehende unabhängige Modulberechtigung/Fachspezifikation erhalten. **– bedeutet kein Recht aus dieser Rolle allein**, keine globale Sperre gegenüber einer separat ordnungsgemäß vergebenen zweiten Rolle.

## M1. Organisationsrollen und Datenklassen

| Rolle | Bestätigter Kern | Kein automatischer Anspruch |
|---|---|---|
| Geschäftsleitung (GL) | Übergeordnete geschäftliche Verwaltung; **nur GL ernennt PL** | Kein implizites medizinisches Vollarchiv, kein SQL-/Service-Role-Zugriff; Detailfelder O10 |
| Produktionsleiter (PL) | Übergeordnete operative Rechte im bestätigten Verantwortungsbereich; AL ernennen; Übernahme wenn Leitung und Stellvertretung abwesend sind | Kein HR-/Gesundheitsdatenrecht und keine Ernennung weiterer PL aus PL-Rolle allein |
| Abteilungsleiter (AL) | Operative Leitung aller ausdrücklich zugewiesenen Abteilungen **gleichberechtigt**; nur eigene zulässige Rechte im eigenen Zuständigkeitsbereich weitergeben | Kein pauschaler Admin, keine fremden Abteilungen, keine Personalrollen-/Auth-Zuordnung oder PL-Ernennung |
| Stellvertreter (STV) | Eigene bestätigte Vertretung je Abteilung; bei Leiterabwesenheit dessen operative Rechte | Keine automatische allgemeine Rechtevergabe; Sonderrollenvergabe nur mit entsprechender Freigabe |
| HR | Personalverwaltung nach explizitem Feld-/Verfahrensumfang | Keine automatischen Produktions-/Werkzeugrechte; medizinische Detailrechte separat O10 |
| Mitarbeiter (MA) | Normale bestätigte Mitarbeiter-/Selbstservicerechte, freigegebene Schichtansicht; betriebliche Fachaktionen nach E | Keine Sicherheitsfeld-/Fremdplan-/Rechteänderung; Person sein oder Mitgliedschaft haben ist keine universelle Modulrolle |
| Technischer Administrator (TA) | Technische Systemverwaltung über definierte Systemfunktionen | **Keine automatischen vertraulichen Personal-/Gesundheitsdaten**, keine GL-/HR-/Produktionsvollmacht, kein freies SQL, keine produktiven DB-Credentials/Service-Role-Schlüssel im Appprofil |
| Scannerkonto (bestehender technischer Sonderfall) | Nur ausdrücklich beauftragte Scan-/Werkzeugaktionen E | Keine neue achte Organisationshierarchie; keine Personal-/Planungs-/BA-Rechte durch Scannerstatus |

Beschäftigtenstatus, aktiver Login, Organisationsrolle, Mitgliedschaft, Leitung und Modulgrant sind sechs getrennte Dimensionen. Eine HR-Person oder ein TA erhält Werkzeug-/Produktionserlaubnis nur über eine **zusätzlich ausdrücklich freigegebene** Fachzuweisung. Es gibt keine automatische Vererbung eines globalen employee-Werkzeugpakets an alle anderen Rollen. Vorhandene erlaubte Mitarbeiterbuchungen werden in der späteren Migration ausdrücklich erhalten, nicht kommentarlos entzogen. Rollenpakete/Kombinationen benötigen O06/O12.

TA ist eine **Anwendungsrolle**, nicht der unkontrollierte Infrastruktur-/Datenbankowner. Ein tatsächlicher DBA/Superuser oder Besitzer von Backup-/Signing-/Service-Schlüsseln kann App-RLS technisch umgehen. Diese Infrastrukturgewalt ist getrennt, minimal, kontrolliert und auditiert zu betreiben; das Modell darf keine nicht durchsetzbare Vertraulichkeitsgarantie gegen einen unbeschränkten DB-Owner versprechen.

## M2. Ernennung, Mitgliedschaft und Vergabe

| Aktion | GL | PL | AL | STV | HR | MA | TA |
|---|---|---|---|---|---|---|---|
| PL ernennen | **B: ausschließlich GL**, unabhängig bestätigt | – | – | – | – | – | – |
| GL erstmals zuordnen/weitere GL ernennen | O01: unabhängiger Bootstrap/Nachfolge | – | – | – | – | – | Keine Selbsternennung |
| AL für konkrete Abteilung ernennen | O: ggf. separate PL-Rolle | **B**, im bestätigten Organisationsscope | – | – | – | – | – |
| Eigene Leitungsabteilungen nutzen | Nur mit eigener Fachauthority | B: operative Oberverantwortung | **B: jede zugewiesene Abteilung gleich**, keine Primärpriorität | B: delegierter Bereich/Abwesenheitskontext | – | – | – |
| Abteilungsmitgliedschaft ergänzen | O14/gesonderte Verwaltungsfreigabe | O14 | O14 | Nur ausdrücklich genehmigt | P: Personalprozess O14 | – | Keine fachliche Zuordnung aus Technikrolle |
| Stammabteilung wechseln | O14 | O14 | Nicht aus Leitung allein | – | P: gesicherter Personalprozess O14 | – | – |
| Modulsonderrolle erteilen | Nur über passende explizite Vergabecapability | Nur über eigene zulässige Rechte/Scope | **B: eigene Rechte ∩ Zuständigkeit**, keine Aufwertung | **B: nur mit Freigabe**, dann dieselben Grenzen | Nur passende ausdrücklich gehaltene/modular delegierbare HR-Rechte | Nur wenn separat ausdrücklich delegierbare Funktion zugewiesen, nicht aus MA allein | Keine fachliche Rechtevergabe aus Technikrolle |
| Sonderrollen prüfen | Gemäß expliziter Zuständigkeit | Gemäß Zuständigkeit | **B: neue Leitung prüft gültige übernommene Sonderrollen** | Nur genehmigter Prüfauftrag | Nur passende Zuständigkeit | – | Nur technische Integrität ohne Inhaltsvollmacht |
| Sonderrollen widerrufen | O07: verantwortlicher Rechteinhaber | O07 | Eigener autorisierter Bereich; Reviewzuständigkeit ist nicht automatisch jedes Widerrufsrecht | Nur ausdrückliche Freigabe | Nur passende HR-Zuständigkeit | Keine eigenen Rechte erhöhen/reaktivieren | Kein pauschaler fachlicher Override |
| Austritt/Deaktivierung auslösen | Berechtigter Lifecycleprozess O09 | O09 | O09 | O09 | P: Personal-Lifecycleprozess O09 | – | P: getrennte technische Kontosperre, kein Personalstammzugriff O09 |

GL darf PL nicht über ein editierbares employees.role oder einen allgemeinen Modulgrant ernennen. Reservierte Organisationsaktionen sind keine delegierbaren Modulsonderrechte. Mehrere Rollen einer Person bleiben getrennte, bestätigte Zuweisungen; die Aktion wird mit der tatsächlich tragenden Autorität auditiert. PL allein darf weder sich selbst noch andere zu PL ernennen. Keine höhere Selbstvergabe durch ein zweites Menü, eine REST-Payload, Templateänderung oder Multi-Scope-Batch.

## M3. Domänencapabilities und sensible Felder

**Alle Bezeichnungen sind geplante semantische Aktionen, keine vorhandenen RPCs.** Technische Namen/Granularität O06; keine Aktivierung. Mehrere Fähigkeiten dürfen nicht zu unspezifischem can_manage zusammenfallen.

| Aktionsfamilie | Zulässiger Zielvertrag | Harte Trennung / aktueller Anschluss |
|---|---|---|
| personnel.core.read / write | HR innerhalb bestätigter Personal-Feldmatrix; weitere Rolle nur ausdrücklich, GL-Details O10 | Rolle/Auth-ID/Primärwechsel/Aktivität nicht per allgemeinem Profilupdate. HR kann ggf. inaktive Personalhistorie lesen, obwohl der inaktive Betroffene keinen Zugang mehr hat |
| personnel.contacts.read_own / write_own | Eigene aktive, verifizierte Person und freigegebene Kontaktfelder | Bestehender security_own_contacts als Integrationspunkt; keine fremde Ziel-ID, kein role/auth_user_id/department_id/id/is_active-Payload |
| identity.bind / organization.appoint | Getrennte vertrauenswürdige administrative Prozesse nach O01/O11 | Kein impliziter HR- oder TA-Vollzugriff auf Geschäftsrollen; jede Organisationsaktion mit eigener Voraussetzung |
| schedule.read_published | Mitarbeiter und weitere ausdrücklich freigegebene Leser | Vierfeldprojektion aus Kandidat vorhanden; konkrete Sichtweite/Namen/Felder D1/O10. Kein Krankheits-/Urlaubsdetail über dieselbe Ansicht |
| schedule.assign / cancel / publish | AL zuständige Abteilung, gültige operative Vertretung, berechtigter PL; Publish **separat** | Vorhandene assignment/publish-RPCs erhalten; Storno/Anträge/Ersatz/Tausch sind teils fehlende atomare Kommandos, nicht durch direkten Tabellenfallback ergänzen |
| absence.request / approve / private_read | Eigen-/Genehmiger-/Privatsicht je konkretem Workflow O10 | Kein automatisches Medizinleserecht für AL/PL/GL/TA. Abwesenheitsgrund nicht an allgemeinen Rollen-/Vertretungskontext senden |
| tools.read / book / master.write / journal.read / journal.correct / inventory.* | Bestehende beauftragte Mitarbeiter-/Scannerrollen erhalten; explizite zusätzliche Modulfreigabe pro Bereich | Lesen ≠ Buchen ≠ Stammdaten ≠ fremdes Journal ändern/löschen ≠ Inventur. HR-/TA-Rolle erzeugt keines dieser Rechte |
| articles.read / edit / release | Bestehende Artikel-Scope-/Assignments-/Reviewerbedingungen E | Autor/Reviewertrennung, eingefrorene Revision und Snapshotintegrität bleiben zusätzlich zur Capability |
| production.work / confirm / route / manage | Bestehende Maschinen-/OP-/Abteilungs-/Teilnehmerberechtigungen E; AL/PL-Fachzuständigkeiten | Confirm ≠ Work ≠ Plan. Mitgliedschaft in Abteilung allein ersetzt keine Maschinen-/OP-Autorisierung |
| setup.own / quantity.book / own_correct / case.resolve / lifecycle.close | Bestehende persönliche Teilnahme, OP- bzw. BA-Hauptauthority E | Betriebs-/Mengen-/Abschlussregeln nicht durch übergeordneten Rollentitel umgehen. BA-Hauptverantwortung ≠ operative OP-Verantwortung |
| system.configuration / technical_sessions / sanitized_audit | TA nur definierte technische Aktionen P | Keine vertraulichen Profile, privaten Logs/Backups oder beliebige SQL-/RLS-/Grant-/Featureaktivierung aus dieser Approlle |
| authorization.audit / review | Jeweils zweck- und bereichsbezogene Freigabe | Technischer Fehlercode/ID statt medizinischem Grund oder Vollpayload; niemand erhält Audit-Gesamtdaten allein durch Systemadministration |

## M4. Effektive Entscheidung und Nichtvererbung

Vorgeschlagene serverseitige Grundregel: gültige vertrauenswürdige Identität **UND** aktiver Actor **UND** aktuelle Bindung/Session-/Berechtigungsversion **UND** explizite passende Fähigkeit für konkrete Aktion/Bereich/Objekt **UND** zulässige Feldprojektion **UND** fortgeltende Fachbedingungen. Ein deaktivierter Actor/gesperrte Authority/überschrittener Scope überstimmt Allow. Kontextbedingungen werden pro Aktion geprüft, nicht aus einem einmaligen UI-Login gecacht.

Mehrere positive Rollen können ausdrücklich Rechte kombinieren, aber verbreitern nicht automatisch jede ihrer Datenprojektionen. HR + separat genehmigte Werkzeugrolle ist möglich, sofern O12 keine Kombination sperrt. TA ohne HR darf keine vertraulichen Personalfelder lesen; bei ausdrücklich genehmigter Doppelrolle ist die HR-Autorität der nachzuweisende Grund, nicht die TA-Rolle. Keine numerische Höherstufung und kein Ausdruck GL > PL > AL > alles.

Mitgliedschaft erzeugt keine Leitung, Leitung keine fremde Abteilung, Stammabteilung kein alleiniger Maschinen-/OP-Scope. Mitgliedschaft darf auch nicht pauschal als neue Voraussetzung für **alle** bereits bestätigten abteilungsübergreifenden BA-/OP-Autoritäten eingeführt werden. O08 legt je Aktionsklasse fest, wann Mitgliedschaft erforderlich ist; keine automatische Mitgliedschaft aus einem Modulgrant.

## M5. Konkrete verbindliche Beispiele

- AL führt A und B, Stammabteilung A: darf bei passender Aktionsfähigkeit A und B gleichberechtigt führen, C bleibt gesperrt. Kein A-Primärfilter.
- HR mit Stammabteilung Personal: darf nicht allein deshalb Werkzeuglager A buchen oder Maschinen B umplanen. Eine zusätzlich ausdrücklich genehmigte Werkzeugrolle A gewährt nur deren Aktionen.
- Zusätzliches Mitglied in B: Berechtigungsauswahl ist verpflichtend, auch die ausdrückliche Auswahl **Keine zusätzlichen Rechte**. Unberührtes Defaultformular ist keine stille Rechtevergabe.
- AL besitzt in A Werkzeuglesen, nicht Stammdatenpflege: kann keine Werkzeugpflege A und überhaupt keine Werkzeugrolle B vergeben. Eine Auswahl A+B wird als gesamter Batch zurückgewiesen, wenn ein Scope unzulässig ist; kein unbemerkter Teilerfolg.
- STV in A bei Leiterabwesenheit: operative Übernahme A, keine automatische Sonderrollenvergabe. Freigabe muss das konkrete Vergaberecht/den Zeitraum/Bereich abdecken; Selbstfreigabe ist keine Freigabe.
- Stammwechsel A→B: **alle bisherigen Sonderrollen**, auch für C oder unbefristete, zunächst suspendieren. Rückkehr nach A reaktiviert sie nicht automatisch.
- Wechsel Leitung A: gültige Sonderrollen A bleiben gültig. Die neue Leitung erhält Reviewaufgaben; fehlender Review ist ohne separat entschiedene Regel kein automatischer Ablauf.
- Actor ausgeschieden/deaktiviert: keine neue Serveraktion unter alten JWTs/Grants. Bereits übertragene Daten lassen sich nicht zurückrufen; In-flight-/Cachegrenzen müssen technisch mitgetestet werden.
- Beide Leitungsfunktionen A abwesend: berechtigter PL übernimmt operativ und erhält eine deduplizierte Benachrichtigung; ausführende Person, vertretene Funktion, A und Autoritätsquelle werden protokolliert.
