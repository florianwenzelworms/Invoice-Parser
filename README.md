# Invoice-Parser
Parses Invoices from .xml to .xlsx


## Changelog
### 1.5
- Fix: Civento Zahlungen können nun wie hsh.olav behandelt werden


### 1.4
- Add: Archivefunction. Processed files do not get deleted anymore, they get moved to "files/Archive" instead
- Fix: Bug removed when no "Purose" is available


### 1.3
- Add: Ordnersuche für Pfadauswahl.
- Rearrange: Buttons 
- Fix: Debugging Logging nun auch in der Log.txt Datei zu sehen


### 1.2
- Add: Logfunktion
- Fix: Pfadangabe
- Standardconfig update


### 1.1
- Add: Lockfunktion für Settings


### 1.0
- Add: Erste funktionierende Version

---

# Dokumentation

![Doku.png](https://github.com/florianwenzelworms/Invoice-Parser/blob/main/docs/doc1.png?raw=true)

## Benutzung
### Bedienung
*Quellpfad*: Hier den Pfad einfügen, der auf den Ordner verweist, in dem die .XML Dateien liegen.

*Zielpfad*: Hier den Pfad zu dem Ordner einfügen, in den die bearbeiteten .xlsx Dateien abgespeichert werden sollen.

*USK*: Die vorkommenden Untersachkonten sind im JSON Format abgespeichert. Die richtige Formatierung ist wichtig, sonst können die Rechnungen nicht bearbeitet werden. Die geschweiften Klammern müssen sich am Anfang und Ende befinden.
Um die Untersachkonten zu bearbeiten, muss der Haken bei *Lock-Settings* entfernt werden. Dies verhindert ein versehentliches Löschen oder Bearbeiten des Feldes. 
Wenn eine Änderung an den Pfaden oder den USK vorgenommen wurde muss immer zuerst mit *Speichern* gespeichert werden. Wenn vorher trotzdem *Ausführen* betätigt wird, werden noch die alten Einstellungen angewandt, selbst wenn die neuen Änderungen schon in den Feldern stehen.

Ist alles korrekt eingestellt kann der Bearbeitungsvorgang durch Betätigen des Ausführen-Buttons gestartet werden. In der Regel dauert der Vorgang, abhängig von der Anzahl der zu bearbeitenden Dateien, nur wenige Sekunden. 

Anschließend werden Informationen zum Ausgang des Prozesses angezeigt

### Ergebnis
![Doku.png](https://github.com/florianwenzelworms/Invoice-Parser/blob/main/docs/doc2.png?raw=true)

Dateien wurden ohne Fehler bearbeitet

![Doku.png](https://github.com/florianwenzelworms/Invoice-Parser/blob/main/docs/doc3.png?raw=true)

Quellordner war leer. Falscher Ordner angegeben oder schon alle Dateien bearbeitet?

![Doku.png](https://github.com/florianwenzelworms/Invoice-Parser/blob/main/docs/doc4.png?raw=true)

Es gab fehlerhafte Dateien bei der Bearbeitung. Möglicherweise gibt die Log-Datei Aufschluss über gegebenenfalls aufgetretene Fehler.

### Fehler
Die Log-Datei *log.txt* befindet sich im Installationsordner des Programmes. 

>*2020-11-12 10:23:23,078 - Invoice Parser - run - 39 - DEBUG - Fehler: >>> 'isKFZ' <<< beim Bearbeiten der Datei: 07319000_iKFZ_20201003_XEPAY21FINANZ_STL_PayPal_10b13c84-a282-4d7e-bd8f-27c828adbbcd - Kopie.xml*

*log.txt* Auszug: *Fehler: >>> 'isKFZ' <<<* beschreibt hier den Fehler. Es handelt sich um ein unbekanntes USK.
