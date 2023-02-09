# Instructies installeren Excel automatiserings add-in
**Ondersteunend pakket voor excel automatisering**

## Windows
### Installeren miniconda
- Download geschikt miniconda pakket - https://docs.conda.io/en/latest/miniconda.html
- Gebruik standaard opties

### Installeren benodigde paketten
- Open conda commmand prompt
- In prompt typ ```conda install --file ```
- Sleep het `requirements.txt` bestand naar de prompt en druk vervolgens op enter

### Installeer addin
- In conda command prompt typ
- `xlwings addin install --file `
- Sleep het `programmingengineer.xlam` bestand naar de prompt en druk vervolgens op enter

### Juiste config excel
- Zet pythonpath naar pad folder met python/excel document

### Scripts plaatsen op juiste plek
Plaats de inhoud van de folder met de scripts in de map `programmingengineer` in een map onder de gebruiker. 
Dus als de gebruiker *Jan* heet dan een folder onder `C:\\Users\Jan\programmingengineer`

# Gereed
De extensie is nu gereed om gebruikt te worden

## Troubleshooting

### Bestand is niet te slepen naar de Anaconda prompt

_Oplossing_ : Plaats het bestand eerst op het bureaublad om vervolgens het bestand vanaf het bureaublad naar de Anaconda prompt te slepen.

# Overig referentie
## Windows on Arm specifiek via venv opztie
*Niet aanbevolen*

Waarschijnlijk werkt de Anaconda distributie nog niet wegens python 3.11.1 in beta. 
Daarom venv aanmaken

Om in shell te activeren dient script settings beleid in windows aangepast te worden
`set-executionpolicy RemoteSigned`
Installeer via pip na activate

Set interpreter `C:\Users\hayer\debugmodule\venv\Scripts\pythonw.exe`



