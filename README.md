# XAF-Auditfiles Project

Dit project bestaat om XAF-auditfiles te importeren in een Pandas DataFrame en hierna te exporteren naar een .CSV-bestand dat vervolgens gebruikt kan worden. Momenteel kunnen V2 (CLAIR2.00) en V3.2 auditfiles geïmporteerd worden. Versies 3.0 en 3.1 geven mogelijk foutmeldingen aangezien ik geen bestanden of specificaties bezit om dit te testen.

## Hoe te gebruiken
Er zijn 3 Python-scripts aanewzig. XAF V2 importeert alleen versies V2 van de auditfiles. XAF V3.2 importeert alleen V3.2 van de auditfiles. XAF V2-3 importeert zowel versies 2 als 3 van de auditfiles. XAF V2-3 is het script dat bijgewerkt wordt met updates. Hier is een check aanwezig welke versie van de auditfile geselecteerd is en op basis hiervan wordt de import uitgevoerd. Bij het starten van het script komt de windows verkenner tevoorschijn. Hier kan 1 bestand worden geselecteerd of meerdere bestanden. Bij het selecteren van meerdere bestanden wordt elke auditfile apart omgezet en wordt er een "geconsolideerde" auditfile aangemaakt waarbij alle geselecteerde auditfiles zijn samengevoegd. De gecreëerde CSV-bestanden komen op dezelfde locatie als waar de geselecteerde auditfiles zijn opgeslagen.

## Functionaliteit

- Omzetten van de Auditfiles in XAF-formaat naar .CSV-formaat.
- Entiteit en boekjaar kunnen worden toegevoegd door dit in de bestandsnaam mee te geven met een "-". Bijvoorbeeld: Github - 2020. Alle tekst voor het streepje komt in de kolom entiteit, de tekst na het streepje komt in de kolom boekjaar.
- Meerdere auditfiles tegelijk importeren en samenvoegen
- Debet/Credit wordt toegevoegd bij V3 en het totaalbedrag wordt toegevoegd bij V2
- Lege kolommen worden automatisch verwijderd

## Toekomstig

- Handleiding voor het gebruik met screenshots
- Automatische import vanuit het Python-script in Caseware IDEA
- Standaardanalyses voor auditfiles
- Visualisaties op basis van de standaard aanwezige gegevens in een auditfile

## Benodigdheden voor het gebruik

De auditfile moet in .XAF formaat aanwezig zijn.

Om de Python scripts te gebruiken moet Python geïnstalleerd zijn. Deze scripts zijn gebaseerd op V3.5.3

De volgende modules worden in de scripts gebruikt en moeten geïnstalleerd zijn:
- Pandas
- Numpy
- Sys
- Os
- xml.etree.ElementTree
- tkinter

In Caseware IDEA 10.4 zijn deze modules aanwezig en hoeft het niet apart geïnstalleerd te worden als de scripts vanuit IDEA worden gebruikt. Bij de instellingen moet wel worden aangegeven dat Python-scripts gebruikt kunnen worden voordat het werkt.
