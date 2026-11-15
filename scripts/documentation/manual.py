"""Generate editable role manual from the SAME help source and real evidence manifest."""
import sys,json,pathlib
from docx import Document
from docx.shared import Inches,Pt,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
out=pathlib.Path(sys.argv[1]);p=json.loads((out/'build-input.json').read_text());doc=Document();sec=doc.sections[0];sec.top_margin=Inches(.7);sec.bottom_margin=Inches(.65)
for style in doc.styles:
 if style.element.pPr is not None:
  for border in list(style.element.pPr.findall(qn('w:pBdr'))):style.element.pPr.remove(border)
for st in ['Title','Subtitle','Heading 1','Heading 2','Heading 3','Normal']:
 doc.styles[st].font.name='Calibri';doc.styles[st].font.color.rgb=RGBColor(0,0,0)
doc.styles['Normal'].font.size=Pt(10)
doc.styles['Normal'].paragraph_format.space_after=Pt(3)
doc.styles['Normal'].paragraph_format.line_spacing=1.05
doc.styles['Heading 3'].paragraph_format.space_before=Pt(5)
doc.styles['Heading 3'].paragraph_format.space_after=Pt(3)
header=sec.header.paragraphs[0];header.text='HUMBEL   |   Benutzerhandbuch   |   Entwicklungsstand '+p['version']
footer=sec.footer.paragraphs[0];footer.text='Nur synthetische Testdaten   ·   Keine Produktivfreigabe   |   Seite '
f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),'PAGE');footer._p.append(f)
doc.add_paragraph('Humbel App Benutzerhandbuch',style='Title');doc.add_paragraph('Bedienung und sicherer Testbetrieb',style='Subtitle')
doc.add_paragraph('Version '+p['version']+' · 14. November 2026')
doc.add_paragraph('Dieses Handbuch führt Mitarbeiter, Abteilungsleitung und Administration durch den aktuellen Entwicklungsstand. Die Bilder stammen aus der tatsächlich laufenden Anwendung mit synthetischen Testdaten. Die Dokumentationsansicht erlaubt keine schreibenden Serveraktionen. Alle Aufnahmen zeigen die synthetische Administrationssicht. Rollenbezogene Bildfolgen für Mitarbeiter und Abteilungsleitung sind noch nicht abgenommen. Ein Screenshot belegt eine sichtbare Oberfläche, keine native Authentifizierungs- oder Datenbankabnahme.')
doc.add_heading('Vor der Präsentation',1);doc.add_paragraph('Der getrennte Testserver läuft nur lokal. Eine Online-Veröffentlichung benötigt ausdrückliche Freigabe, TLS und Zugangsschutz. Neue SQL-Module und modulare Rollen sind nicht produktiv aktiviert. Digitale Qualitätsprüfung und Dokumentenverwaltung sind weiterhin Konzepte. Bei Platzhaltern zeigt die Anwendung diesen Entwicklungsstand ausdrücklich an.')
doc.add_paragraph('Hilfe in der Testinstanz: Taste F1 oder „Hilfe zu dieser Ansicht“. Die Quelle js/help-topics.js erzeugt App-Hilfe und diese Kapitel. Die vollständige Feld- und Modalabdeckung ist noch nicht abgenommen; unbekannte Felder sind nicht durch diese Übersicht freigegeben.')
for role,label in [('employee','Mitarbeiter'),('leader','Abteilungsleitung'),('admin','Administration')]:
 doc.add_page_break();doc.add_heading(label,1)
 for index,c in enumerate([x for x in p['chapters'] if x['role']==role]):
  if index:doc.add_page_break()
  doc.add_heading(c['title'],2);doc.add_paragraph(c['purpose']);doc.add_paragraph('Stand '+p['version']+' · Voraussetzung '+role+' · Status '+c['status'])
  if c.get('image'):
   image=out/c['image'];w,h=Image.open(image).size
   shown=min(h,800);width=5.0 if c['id']=='articles' else 5.7;picture=doc.add_picture(str(image),width=Inches(width),height=Inches(width*shown/w))
   if shown<h:
    rect=OxmlElement('a:srcRect');rect.set('b',str(round((1-shown/h)*100000)));picture._inline.graphic.graphicData.pic.blipFill.insert(1,rect)
   doc.add_paragraph(c['caption']+' Ansichtsausschnitt; die vollständige Aufnahme liegt im Screenshotpaket.',style='Caption')
  else:doc.add_paragraph('BILDNACHWEIS OFFEN · BLOCKED_NOT_TESTED. Für diese Ansicht liegt kein verifizierter Screenshot dieses Builds vor. Geplante Masken werden nicht nachgebildet.')
  doc.add_heading('Eingaben und Voraussetzungen',3)
  for f in c['fields']:doc.add_paragraph(f['name']+': '+f['description'])
  doc.add_heading('Schritte',3)
  for number,step in enumerate(c['steps'],1):doc.add_paragraph(str(number)+'. '+step)
  doc.add_heading('Speichern und Verarbeitung',3);doc.add_paragraph(c['effect'])
  doc.add_heading('Typische Fehler',3)
  for error in c['errors']:doc.add_paragraph(error,style='List Bullet')

doc.add_page_break();doc.add_heading('Freigaben und verbleibende Arbeiten',1)
for text in ['Native PostgreSQL GoTrue PostgREST JWT RLS und echte Parallelität separat abnehmen.','Vollständige Browsermatrix, Telefon und Tablet sowie unveränderte visuelle Referenzen prüfen.','Alle seltenen Eingabemasken und fachlichen Feldhilfen anhand des dynamischen Inventars ergänzen und prüfen.','Hintergrunddienst für Lernangebote, sichere Arbeitszeitkontexte, Inhalte und Ausbildungsfreigaben aus Issue 39 fertigstellen.','Online-Ziel, Verantwortliche, Zugangsschutz, Backup und Wiederherstellungsprobe gesondert freigeben.']:
 doc.add_paragraph(text,style='List Bullet')
doc.add_heading('Störungen und Support',1)
doc.add_paragraph('Fehlermeldung mit Datum, Uhrzeit, betroffenem Modul und reproduzierbaren Schritten erfassen. Keine Passwörter oder sensiblen Produktionsdaten in Screenshots teilen. Der betriebliche Supportkontakt ist noch festzulegen. Nach einem unklaren Buchungsergebnis zuerst den bestätigten Serverstand prüfen.')
doc.save(out/('Humbel_Benutzerhandbuch_'+p['version']+'.docx'));print('DOCX generated')
