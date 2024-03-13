# kontaktSupporter  
### Sorry, german language only!  
  
### Status: Alpha, testbar!

### Kurz-Hilfe mit der F1-Taste!  

  
#### Letzte Änderungen:  
  
Version (&gt;=)| Änderung  
------------ | -------------  
0.045 | Vorschaufenster verbessert, Anhang-Bild ansehen, Makrofunktion  
0.040 | In "ksConfig.ini": externalAppOpenDelay=500 statt 5000, eMailAppOpenDelay=4000 löschen  
0.036 | _lastUsed.txt ersetzt durch _previous.txt, "ksBodyFiles" ersetzt durch "ksInhalte", Neue Einträge mittels "Prototypen" Memo ist jetzt das letzte Feld.  
0.035 | Trennzeichen "|" durch "¦" ersetzt (0x00A6)!  
0.032 | {°&lt;DATEINAME.txt&gt;°} geändert in °&lt;DATEINAME.txt&gt; (Leerzeichen am Ende!)  
0.031 | {°&lt;DATEINAME.txt&gt;°} geändert in {°&lt;DATEINAME.txt&gt;°}  
0.030 | Support für Textbausteine im Subjekt-Feld um im Text (Body-Feld) s.u.  
  
#### Beschreibung  
*kontaktSupporter* dient dem Versenden von eMails,  
die oft mit geringfügig geändertem Inhalt verschickt werden müssen.  
  
Kann man auch alles mit der eMailApp machen,  
*kontaktSupporter* bietet aber bis jetzt folgende besondere Features:  
* Importieren des eMail-Inhalts aus einer vorher festgelegten Datei,  
* Öffnen externer Apps im eMail-Anhang-Ordner,  
* Nutzung von Textbausteinen (nur in Textnachrichten oder RTF-Nachrichten), nicht mit DOCX, ODF oder anderen Formaten.  
  
Als eMail-Inhalte sind möglich:  
* reine Textdateien (Dateiendung ".txt, .rtf")  
* alle anderen Formate, sofern ein Programm (App) zum Öffnen dieses Formats installiert ist,  
  welches die Windows-üblichen Hotkeys (s.u.) unterstützt.  
  Dazu wird das entsprechende Programm kurz geöffnet, um die Daten in die Zwischenablage zu kopieren.  
  Wichtig: Das entsprechende Programm darf vorher nicht geöffnet sein,  
  es wird dann als neuer Prozess geöffnet und *kontaktSupporter* kann es dadurch erkennen!  
  
Zum Senden wird das Standard eMail-Programm verwendet, welches mit "mailto://" verknüpft ist.  
Es muß Windows-üblichen Hotkeys (s.u.) unterstützen.  
Wichtig: Auch das eMail-Programm darf vorher nicht geöffnet seindamit es von *kontaktSupporter* erkannt werden kann!  
  
*kontaktSupporter* funktioniert so, als ob ein Benutzer die Eingaben macht, nur komfortabler.  
Während *kontaktSupporter* arbeitet dürfen daher die Fenster (z. B. des eMailprogramms) nicht geschlossen,  
oder andere Programme gestartet werden, die das Fenster überdecken würden.    
Dadurch wird *kontaktSupporter* sehr universell verwendbar,  
kann aber "bei der Arbeit" durch Benutzereingriffe gestört werden.  
  
#### Hinweise  
  
##### Öffnen von Dateien  
Dateien werden immer mit dem der Dateiendung zugeordneten Programm geöffnet.
Welches das ist, ist eine Windows Einstellung, siehe [Windows 10: Dateiendung einem Programm zuordnen](https://techmixx.de/windows-10-dateiendung-einem-programm-zuordnen/)
  
Hat die Datei keine Endung, wird automatisch die Endung ".txt" angenommen.  
Das gilt für alle verwendeten Dateien, also auch Textbausteine usw.  
Dadurch kann man die Endung in der Tabelle auch weglassen,  
wodurch die Tabelle deutlich übersichtlicher wird!  
  
##### Hinweis zu Memo-Feld  
Da das Memo-Feld mehrzeilig sein kann, werden Zeilenvorschübe speziell kodiert,  
da sie sonst in der Tabellendatei einen neuen Eintrag starten würden.
In der Bearbeiten Anzeige werden sie aber wieder dekodiert angezeigt,  
Die Kodierung (einstellbar) für einen Zeilenvorschub ist default: 0x00AC (das "¬"-Zeichen).  
  
##### Hinweis zur Anzeige der Tabelle  
Der Anhang (Pfad) wird in der Tabelle gekürzt dargestellt, wenn er sehr lang ist.  
   
Bearbeiten und Vorschau werden immer in zentrierten Fenstern angezeigt (TODO).  
Die Größe und Position der Fenster ist änderbar, wird aktuell aber nicht gespeichert (TODO).  
  
#### Menü  
Es gibt zwei Menüs, das Haupt-Menü und das Menü im "Bearbeiten"-Fenster.  
Der Menüpunkt \[🔗 Anhang] ist in beiden Menüs identisch.  
Das Haupt-Menü kann mit Text oder platzsparend nur mit Icons angezeigt werden (Taste [+...] oder [-...])  
Eine Beschreibung der einzelnen Tasten befindet sich in der Kurz-Hilfe,  
welche mit der F1-Taste abrufbar (Datei "kontaktSXQHelp.html") ist.  
  
#### Anhänge  
In der Tabelle kann für jede Zeile der Pfad zu einem Verzeichnis (= Ordner) eingetragen werden, welches die Anhänge enthält.  
In der Datei "ksExternalApps.txt" (aktuell ist der Name fest TODO) kann man externe Apps (=Programme) eintragen,  
die dann im Menü unter \[🔗 Anhang] gestartet werden können.  
Deaktivieren einzelner Einträge mit Kommentarzeichen (";" oder "#" oder "//") am Zeilenanfang.  
Nach Änderungen einmal auf "Refresh" klicken!  
  
Besteht der Anhang aus nur einer einzigen Datei, kann statt dem Pfad des Ordners der Pfad dieser Datei eingetragen werden.  
Beim Zugriff mit einem Dateimanager (Menü -&gt; \[🔗 Anhang] oder Menü -&gt; \[Bearbeiten] -&gt; \[🔗 Anhang])  
wird dann allerdings nicht der Ordner, sondern diese Datei mittels  
dem der Dateiendung zugeordneten Programm geöffnet, was eventuell unerwünscht ist.  
  
Ich muss oft Bildateien verschicken, diese müssen aber meistens noch gedreht und verkleinert werden.  
Dazu verwende ich "ACDSee Pro5", dann muss ich gelegentlich auch noch weitere Dateien hinzufügen,
also einen Dateimanager im Ordner der Anhänge öffnen.  
Ich verwende dazu "Dopus" (nicht kostenlos), 
neben dem Explorer gibt es aber noch einige gute kostenlose Dateimanager, s.u..  
Für Bilddateien verwende ich den speziell dafür geeigneten Dateimanager "ACDSee Pro5" (nicht kostenlos)   
  
Man spart viel Zeit, wenn man nicht überlegen muss, wo denn jetzt wieder diese Anhänge sind ...
(Anhänge sind standardmässig in einem Unterordner von "ksAnhang", der ein Unterordner vom "baseDirectory" ist,  
standardmässig also dem App-Ordner.  
  
Das kann aber alles in der Konfiguration leicht angepasst werden.  
  
Eine durchdachte Verwendung von Unterordnern hilft dabei immer sehr,  
also z.b. "2024" oder "2024/01" oder entwürfe/2024/01 usw.  
  
Verändert das externen Programm irgendwelche Daten, die dann wieder von kontakSupporter verwendet werden,  
kann mit der Menü-Taste "⟳ Refresh" die Anzeige aktualisiert werden (App wird neu gestartet).  
Bei den Funktionen aus dem Menü wird die Anzeige meist automatisch aktualisiert,  
sobald kontakSupporter wieder verwendet wird.  

###### Anhang-Bildvorschau  
Zum schnellen Überprüfen eines Bildes, funktioniert aber nur, wenn das Anhang-Feld auf ein einzelnes Bild zeigt,  
also nicht, wenn das Anhang-Feld auf einen Bild-Ordner zeigt.  
Ausgelöst durch: \[Umschaltaste] + einfacher/doppelter Mausklick.  
  
##### Festes Verzeichnis  
Oft muß man Anhang-Dateien noch verschieben/kopieren,  
dann sind Dateimanager mit 2 Fenstern, also 2 Ordner-Ansichten die bessere Wahl.  
Der zweite Ordner wird dabei in eckigen Klammern hinter dem Dateimanagernamen angegeben.  
  
##### Von *kontaktSupporter* unterstützt Dateimanager, die zwei Ordner nebeneinander anzeigen können  
* [Directory Opus ("Dopus")](https://www.gpsoft.com.au/)  
* [FreeCommander](https://freecommander.com)   
* [MultiCommander](http://multicommander.com)  
    
In der Datei "ksExternalApps.txt" müssen die Pfade zu den jeweiligen \*.exe-Dateien einmalig angepasst werden!  
Die vorhandene Datei "ksExternalApps.txt" kann als Vorlage dienen.  
   
Für FreeCommander und MultiCommander sind dort Pfade zu der jeweiligen portablen Version eingetragen.   
  
Der feste Ordner erscheint jeweils im linken Tab, der des Anhang-Ordners im Rechten.  
  
#### *kontaktSupporter* Hotkeys  
  
Hotkey-Tasten | Funktion  
------------ | -------------  
\[F1] | Kurz-Hilfe  
  
#### Download Zip-file
Github: [kontaktSupporter.zip](https://github.com/jvr-ks/kontaktSupporter/raw/main/kontaktSupporter.zip)  
Die Zip-files enthalten keinen Quellcode!  
  
#### Installation:  
Zip-file entpacken, z.B. in den Ordner: "C:\jvrks\kontaktSupporter\" .  
  
Der Ordner muß von *kontaktSupporter* vom aktuellen Benutzer beschreibbar sein,  
also NICHT in einen Unterordner von "C:\Program Files" oder C:\Program Files (x86)" entpacken!  
  
#### Textbausteine  
Es können in Text-Dateien vorhandene Inhalte eingefügt werden, und zwar in:  
* dem Adresse Feld (eMail-Adressen mit Komma trennen)
* dem Cc Feld (eMail-Adressen mit Komma trennen)
* dem Bcc Feld (eMail-Adressen mit Komma trennen)
* dem Subjekt Feld  
* dem Body-Feld (nur Body-Felder die reine Text-Dateien sind (RTF geht auch)!)  
  
Die Platzhalter haben das Format: °&lt;DATEINAME.txt&gt; (Leerzeichen am Ende!) (ab Version 0.032)  
Der DATEINAME selbst darf KEINE Leerzeichen (oder andere [Whitespace-Zeichen(https://de.wikipedia.org/wiki/Leerraum)) enthalten
Die Datei Extension ".txt" kann auch weggelassen werden!  
Die Dateien müssen sich im Ordner "textModulsFilePath=" (default: "ksTextBausteine") befinden.  
Beispiel:  
Die Datei "Mahnung27896.txt" enthält:  

````
{°absender.txt°}


{°kunde27896.txt°}

{°sgdh.txt°}

he, Leute, her mit der Kohle!

{°mfg.txt°}
````
Die Dateien enthalten:  
"absender.txt" Die Absenderadresse  
"kunde27896.txt" Die Empfängeradresse  
"sgdh.txt" den Text: "Sehr geehrte Damen und Herren,"  
"mfg.txt" den Text: "Mit freundlichen Grüßen..."  
"mahnstufe3.txt" den Text: Letzte Mahnung! 😖😖😖 
  
Tipp: Häufig verwendete Texbaustein-Dateien sollten mit einem Unterstrich beginnen,  
dadurch werden sie in Dateimanagern immer ganz oben angezeigt.  
  
Der Eintrag in "ksTabelle.txt" lautet:  
Mahnung 27896¦schuldner@saeumig.de¦¦¦°mahnstufe3¦Mahnung27896.txt¦¦1¦°mahnstufe3 ist gleichbedeutend mit: °mahnstufe3.txt¬(bei °Dateiname.txt kann das ".txt" weggelassen werden!)¬Im MemoFeld werden Zeilenvorschübe speziell kodiert, da sie sonst die Tabellendatei stören würden.¬Hier wird es aber wieder dekodiert angezeigt!¦
  
Die generierte eMail sieht dann so aus:  
  
An-Feld: schuldner@saeumig.de  
Im Betreff: Letzte Mahnung! 😖😖😖  
  
Im eMail-Body steht:  
````
Irgendwer
Irgendwo
Irgend eine Stadt


Der böse Kunde 27896
Auf dem Holzweg 99
0001 Fast im Gefängnis

Sehr geehrte Damen und Herren,

he, Leute, her mit der Kohle!

Mit freundlichen Grüßen

Generaldirektor Dr. Mustermann
````
#### Makros  
Makros sind eine Buchstabenkombination die eine Textersetzung auslösen.
Im Gegensatz zu Textbausteinen, die das Einfügen von in einer Datei vordefiniertem Text erlauben,  
also quasi ein statischer Vorgang, ermöglichen Makros das Einfügen von dynamisch generierten Inhalten.  
Ein typisches Beispiel ist das Einfügen des aktuelle Datums.  
Es sind nur Makros möglich, die vorher im kontaktSupporter-Programm eingebaut wurden.  
Eine Selbstdefinition wäre zwar möglich, sprengt aber den Rahmen dieses simplen Tools.  
  
Aktuell sind definiert:  

Makro | Funktion
------------ | -------------
\#dl# | Datum lang
\#d# | Datum kurz
\#u# | Uhrzeit  
  
Ein doppeltes Escape-Zeichen hebt die Escape-Funktion auf,
dadurch wird das Makros ausgeführt.
Das Escape-Zeichen wird dann als normales Zeichen angezeigt.
Es werden alle Textelemente die der Makrostruktur entsprechen ausgewertet,
unbekannte Makros werden entfernt.
  
Weitere Makros folgen nach Bedarf!
   
##### Allgemeines Makroformat:
Nach dem einleitenden "#"-Zeichen können 1-3 weitere Zeichen folgen, danach muß ein sogenanntes Whitespace-Zeichen folgen,
das ist z.B ein Leerzeichen.  
Kommt diese Kombination mit dem "#"-Zeichen zufällig auch im Text vor,  
kann die Makrofunktion durch voranstellen eines "\\"-Zeichen (sogenanntes Escape-Zeichen) unterbunden werden,  
also z.B. "#d#" erscheint das Datum, bei "\\#d#" erscheint nur "#d#".  
Siehe auch Datei "makrotest.txt" als Text und in der Vorschau (in der Beispiel Tabelle: 005 Text-Makros)!  
  
#### Gute Texteditoren  
Hier nur solche die STRG+Mauswheel-Zoom unterstützen:  
  
* [Notepad++](https://notepad-plus-plus.org/) Immer noch mein Favorit!  
* [SciTE](https://www.scintilla.org/SciTE.html) Super für Autohotkey,  
* [EmEditor Free](https://www.emeditor.com/) Nutze ich neben notepad++,  
* [CudaText](https://cudatext.github.io/) Vor kurzem erst entdeckt.  
  
Ich hatte auch mal einen programmiert, in Scala, der konnte sogar multible Verschachtelungen von allen Klammern
und gleichzeitig den in Scala verwendeten mehrfach-Hochkommas verarbeiten.  
Allerdings verwendete er die veraltete Swing-Gui-Bibliothek, ich wollte dann auf JavaFX umstellen,
bin aber nicht dazu gekommen. Es gibt aber inzwischen so viele gute Editoren ...
  
Vor nicht allzulanger Zeit habe ich **Aottext** geschrieben, allerdings in Autothotkey Version 1.  
Das ist aber kein richtiger Editor, sondern eine Art Notizblock,  
der sich die Inhalte in automatisch erzeugten Dateien merkt.  
Man kann dann durch die "Vergangenheit" (Dateien) mit der Maus (Alt+) blitzschnell durchscollen.  
Auch "verschwindet" (Hide) er von selbst an den unteren Bildschirmrand oder  
ganz in den Hintergrund auch per Hotkey Alt + A,  
dabei speichert er bei Änderungen jedesmal in eine neue Datei, wenn er verschwindet.  
  
#### Vorlagen für Inhalte  
Unter "Bearbeiten -&gt; Menü -&gt; Inhalt" kann die Inhaltedatei ausgewählt und bearbeitet werden.  

Für Vorlagen gibt es einen eigenen Ordner (intern als "protoFilePath" bezeichnet),  
default ist der Ordner "ksVorlagen" (wenn protoFilePath=ksVorlagen ist),  
so kann man aus den dort vorhandenen Dateien eine auswählen,  
welche dann als Vorlage für eine neue Inhalts-Datei verwendet wird.  
Natürlich muß man der dieser neue Inhalts-Datei auch einen neuen Namen geben.  
Die Dateierweiterung (definiert den Typ der Datei) wird automatisch an die Vorlage angepasst.  
Ohne Dateierweiterung wird immer "\*.txt" angenommen.  
Eine Menütaste "♒ Vorschau" befindet sich auch im Hauptmenü,  
damit man dort schnell die einzelnen Tabelleneinträge durchsehen kann.  
(Vorschau nur für Textdateien).  
  
#### Konfigurations-Datei  
Die Konfigurations-Datei "ksConfig.ini", wird automatisch erstellt, falls sie nicht schon vorhanden ist.  
(Personalisierte Konfiguration s.u.)  
  
Bei anderen Datei-Formaten als reinem Text (\*.txt-Dateien),  
also Office-Dateien usw. muss das entsprechende Programm kurz geöffnet werden,  
um die Daten zu kopieren. Das dauert etwas länger, je nach Programm. 
Es wird bei "externalAppOpenDelay=500" maximal ca. 20 mal 500 Millisekunden, also 10 Sekunden gewartet.  
     
Ist der PC extrem langsam, muß dieser Wert und auch bei "sendDelay=200" erhöht werden.  
Nach manuellen Änderungen eines Wertes muss *kontaktSupporter* neu gestartet werden (Button \[Aktualisieren] im Menü).  
  
Die Einträge im Abschnitt \[gui] werden von der App zum Speichern der Positionsdaten des App-Fensters verwendet. 
  
#### In Windows nicht erlaubte Dateinamen  
Windows verbiete die Verwendung folgender Dateinamen:    
AUX, CON, NUL, PRN, COM1, COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9, LPT1, LPT2, LPT3, LPT4, LPT5, LPT6, LPT7, LPT8 und LPT9..
  
Sonderzeichen in Dateinamen mit Ausnahme von "_" und "-" sind auch nicht erlaubt.   
  
Weiterhin sollten keine Umlaute (ö,ä,ü,ß) verwendet werden,  
diese können Probleme machen, sobald die Dateien irgendwie übertragen werden, z.B. bei der Datensicherung.  
  
#### Zusätzlich in *kontaktSupporter* nicht erlaubte Dateinamen  
Dateinamen ohne Extension sind nicht erlaubt, da sie keiner App zugeordnet werden können,  
auch würde *kontaktSupporter* sie als "\*.txt" Datei einordnen und sie dann nicht finden.  
  
#### Ordner-Struktur und Speicherung der Inhalte  
Alle von der App verwendeten Dateien sind reine Textdateien (UTF-8) (ausser den Dateien für den eMail-Body).  
Das Trennzeichen ist "|", dieses darf dadurch sonst nicht verwendet werden.  
Es wird keine Datenbank eingesetzt.  
  
Die Ordner-Struktur (auch oft Verzeichnis-Struktur genannt) ist nur von Interesse,  
wenn man etwas daran ändern möchte.  
  
Im Ordner (Rootordner) in dem sich die *kontaktSupporter* befindet,  
muß sich auch immer die Konfigurations-Datei "ksConfig.ini" befinden.  
  
Alle anderen Ordner sind normalerweise Unterordner dieses Ordners.  
Sie können jedoch bei Bedarf auch Unterordner eines anderen Ordners sein,  
dieser wird dann als "baseDirectory" bezeichnet.  
Ist das baseDirectory="" leer, so ist es automatisch der Rootordner.  
  
Die Ordner-Struktur für Tabellen und Inhalte ist standardmässig wie folgt:  
  
Ordner-Name (in der Config) \*1) | Typischer Dateiname | Inhalt | Menü Taste  
------------ | ------------- | ------------- | -------------   
ksTabellen (tableFilePath)| ksTabelle.txt  \*2) | Einträge für die Tabelle \*3) | \[Tabelle wählen], \[Zoom/Edit]   
ksBodyfiles (bodyFilePath)| \*.txt, \*.odt usw. | Inhalt der eMail  | \[Zoom/Edit]  
ksAnhang (attachmentsPath)| \*.jpg, \*.pdf usw. | eMail-Dateianhänge  | \[Extern] \*4)  
ksTextBausteine (textModulsFilePath)| \*.txt | Text-Bausteine  | - 
baseDirectory | ksExternalApps.txt | Liste externer Apps  | \[Extern], unten \[Edit]  
baseDirectory | ksConfig.ini | Konfiguration  | \[Konfiguration], \[Konfiguration bearbeiten]  
baseDirectory | ksGesendet.txt | Log gesendete eMails | \[Konfiguration], \[Gesendet ansehen/bearbeiten]  
  
\*1)  
Wird so auch in der Konfiguration verwendet.  

\*2)  
Es können dort mehrere Tabellen gespeichert werden,  
die dann auswählbar sind. Beim Speichern wird jeweils eine Backupdatei erstellt ("XX_previous.txt").    
"Menü -&gt; Tabelle -&gt; Ordner" öffnen.  
  
\*3)
Einträge werden durch ein **"|"-Zeichen** getrennt,  
dadurch ist dieses Zeichen reserviert und darf sonst **nicht verwendet werden**!  
  
\*4) Dort die App zu Öffnen des Ordners der Anhänge wählen, siehe Abschnitt "Anhänge" weiter oben.  
  
Achtung: "baseDirectory=" Änderungen sind momentan noch komplett ungetestet (TODO)!
  
Wird der Eintrag "baseDirectory=" (der Basis-Ordner) geändert,  
können die Standard-Ordner Unterordner eines beliebigen Ordners werden.  
Die Konfigurations-Datei "ksConfig.ini" verbleibt aber immer im App-Ordner.  
  
Nach manuellen Änderungen eines Wertes in der Konfigurations-Datei,  
muss *kontaktSupporter* neu gestartet werden (Button \[Aktualisieren] im Menü).  
  
Bezüge auf Dateien können auch absolute sein (vollständige Pfadangabe),  
das "baseDirectory" und die Standard-Ordner sind dann ohne Bedeutung.  
Die Ordner können auch absolut sein (vollständige Pfadangabe),  
das "baseDirectory" ist dann ohne Bedeutung.  
  
Klingt alles kompliziert, ist es auch. 😀  
und noch nicht fertig ...  
  
#### Personalisierte Konfiguration:  
Alternativ kann als Dateiname der Konfigurations-Datei ein Name verwendet werden, der wie folgt aufgebaut ist:  
"_kontaktsupporter_&lt;NUTZERNAME&gt;.ini".  
(Diese Konfigurations-Datei wird vorrangig ausgewertet.)  
  
* NUTZERNAME ist der Name des aktuellen Windows-Nutzers.  
* Bitte die Unterstriche vor und hinter *kontaktSupporter* beachten.  
* Der Dateityp muß ".ini" sein.  
* Es muß eine ANSI oder eine UTF-8 kodierte Datei sein, also eine reine Textdatei.  
* Als "seletedTableFile" kann ein beliebiger aber gültiger Dateiname verwendet werden,  
  es empfiehlt sich aber dann _ksMainTable_&lt;NUTZERNAME&gt;.txt zu verwenden,  
  gleiches gilt für den Ordnerpfad "bodyFilePath", empfohlen ist dann "_bodyfiles_&lt;NUTZERNAME&gt;".  
  
#### Besonderheiten  
* In ein leeres Subject-Feld wird immer ein "-"-Zeichen eingefügt.  
* Für Text Macros verwende ich [Lintalist](http://lintalist.github.io/)  
  
#### Sourcecode  
Es ist ein Autohotkey V2 Script: "kontaktSupporter.ahk" 

#### Startparameter  
* **remove**, beendet *kontaktSupporter* falls es gerade offen ist.  
  (Wird nur zum Kompilieren benötigt).  
  
#### Bekannte Fehler / Bugs  
Fehler / Bug | Type | behoben in Version  
------------ | ------------- | -------------  
Die Bildgrösse stimmt nicht in der Bildvorschau (Umschalt + Klick) |  Bug | 0.051
~~Bei Inhalt "Datei auswählen" wird ein Unterverzeichnis als absoluter Pfad eingetragen~~ | Bug | 0.043
 
#### Windows-übliche Hotkeys  
Hotkey-Tasten | Funktion  
------------ | ------------- 
\[Strg] + \[a] | Alles markieren  
\[Strg] + \[c] | Kopieren  
\[Strg] + \[v] | Einfügen  
\[Strg] + \[Home] | An den Anfang springen  
  
Steuerung ("Strg") wird im englischen als Control ("Ctrl") bezeichnet.  

#### Antivirus  
Durch das extensive Nutzen der Zwischenablage wird *kontaktSupporter* von manchen Scannern als Virus erkannt.  
Auch der Windows-Standardvirenscanner "Defender" macht des öfteren Probleme, er löscht es einfach.  
Das passierte mit mit etlichen Autohotkey Scripten, daher habe ich "Defender" deaktiviert und mir
"ESET NOD32 Antivirus" gekauft.  
Das ist klasse, NIE WIEDER DEFENDER!  
  
#### License: GNU GENERAL PUBLIC LICENSE  
Please take a look at [license.txt](https://github.com/jvr-ks/kontaktSupporter/raw/main/license.txt)  
(Hold down the \[CTRL]-key to open the file in a new window/tab!)  
  
Copyright (c) 2024 J. v. Roos  
  
  
<a name="virusscan">  


