# AutoFrancesco

Legge i PDF dei documenti di FLOTT e compila i registri HACCP aziendali,
aggiungendoli all'archivio in `~/HACCP` (configurabile tramite `HACCP_DIR`).

| Da | A |
|---|---|
| `PDF Vendite` — i DDT verso FLOTT | DR.RPD.00 Riepilogo produzione |
| `PDF Vendite` | DR.CPF.70 pesate, un file per formato |
| `PDF Acquisti` — le fatture da FLOTT | DR.CAM, DR.CAO, DR.CAI e REGISTRO GIACENZE |

## Come si usa

1. Scarica i PDF nuovi e mettili in `PDF Vendite` e `PDF Acquisti`. **I vecchi si tengono**:
   rileggerli costa pochi millisecondi e permette di rifare tutto da capo se serve.
2. Doppio clic su **genera.command**.
3. Se a fine esecuzione compare il riquadro di avviso, apri il log e guarda cosa e' rimasto
   fuori. Se non compare, non c'e' niente da controllare.

Le segnalazioni si vedono **una volta sola**, quando quel documento viene letto la prima
volta. Un caso come il tonno, che si ripete a ogni fattura ma e' voluto, non torna a
disturbare. L'elenco dei documenti gia' letti sta nei file `*-visti.txt` dentro `src/Log`,
uno per programma: cancellando una riga, quel documento torna a essere segnalato.

Il terminale elenca **solo quello che e' stato aggiunto** e **solo i file che sono
cambiati**: se sotto una sezione compare solo la riga di riepilogo, li' non c'era nulla di
nuovo. Un file che non cambia non viene nemmeno riscritto, cosi' la sua data di modifica
resta quella vera.

I registri di accettazione a schermo si chiamano **pesce**, **olio** e **vasi**, non con
le sigle:

```
  fattura 121602 del 26/08/2026
  + pesce: 121602_260826  (1 lotto)
  + olio: 121602_260826  olio di semi di girasole – lotto 26240 kg 972
  + vasi: 121602_260826  VASO 235 G erm. quantita' 3388
  + giacenze  PESCE FAO 34

36 fatture lette, fogli aggiunti: pesce 1, olio 1, vasi 1; 0 anomalie
  file modificati:
    2026/DR.CAM.00- Controllo in accettazione MAT.PRIME e SEMILAVORATI.xlsx
    HACCP/REGISTRO GIACENZE 2026.xlsx
```

Tutto viene **aggiunto**: i documenti gia' inseriti vengono riconosciuti e saltati, quindi
si puo' rilanciare quante volte si vuole senza creare doppioni.

**Prova/prova a vuoto.command** fa le stesse cose dentro `Prova/Generati`, riproducendo la
struttura dell'archivio senza toccarlo: serve a guardare il risultato prima di riversarlo.

## Com'e' fatta la cartella

```
AutoFrancesco/
├── genera.command          scrive nell'archivio HACCP
├── PDF Vendite/            i DDT da lavorare
├── PDF Acquisti/           le fatture da lavorare
├── src/                    i programmi e i modelli
│   ├── Modelli pesate/     una scheda vuota per formato
│   ├── Modelli pesate esterni/   le stesse, con le tolleranze degli esterni
│   ├── Log/                un file per programma, cresce in coda
│   └── Backup/             una copia .bak di ogni file prima di modificarlo
└── Prova/                  prova a vuoto.command e la sua cartella Generati
```

Copie di sicurezza e log **non finiscono mai nell'archivio**, che resta pulito.

I clienti e i fornitori diversi da FLOTT hanno una cartella a parte, **AutoFrancesco
Esterni**, che usa questi stessi programmi e scrive in `HACCP/ESTERNI`. Vedi il suo README:
qui dentro ci sono i modelli `modello_est_*.xlsx`, `acquisti_esterni.py` e l'opzione
`--esterni` del riepilogo.

L'archivio e' organizzato cosi':

```
HACCP/
├── FORMATI E CAPSULE.xlsx          la tabella formato -> capsula e coefficiente
├── REGISTRO GIACENZE 2026.xlsx     l'anno in corso sta nella radice
└── 2026/
    ├── 1 TRIMESTRE/  2 TRIMESTRE/  i trimestri chiusi
    ├── DR.RPD.00, DR.CAM, DR.CAO, DR.CAI    il trimestre in corso
    └── PESATE/                     un file per formato, tutto l'anno
```

## Cosa fare a fine trimestre e a fine anno

**A inizio trimestre**: sposta i quattro registri da `HACCP/2026/` in una cartella
`3 TRIMESTRE`. Al giro dopo il programma ne crea di nuovi, vuoti. Le pesate e il registro
giacenze non si toccano: quelli continuano per tutto l'anno.

**A inizio anno**: prepara il registro giacenze nuovo copiando quello vecchio e
archiviando il precedente. Finche' non c'e', il programma lo dice per nome e non registra
quelle righe. Tutto il resto e' automatico: `HACCP/2027/` e le sue sottocartelle nascono
da sole, e le pesate ripartono dalle schede vuote in `src/Modelli pesate`.

L'anno viene preso **dal documento**, non dalla data di oggi: un DDT del 29 dicembre
lavorato a gennaio finisce comunque in `HACCP/2026`.

## Riepilogo produzione (vendite)

Un foglio per DDT, chiamato `numero_ddmmyy`, con un blocco di 5 colonne per prodotto.

| Campo | Origine |
|---|---|
| PRODOTTO | descrizione ripulita da confezione, lotto e scadenza |
| DATA | il giorno della scadenza riportato sul mese di emissione del DDT; se quel giorno cade dopo l'emissione, il mese precedente. Il lotto non entra nel calcolo |
| LOTTO, SCADENZA | dalla descrizione |
| Vasi, Capsule | numero di confezioni x pezzi per confezione |
| Acciughe, Olio | vasi x coefficiente del formato, arrotondato |
| QUANTITA' PRODOTTE | numero di confezioni e pezzi per confezione |

**ORIGINE** e **LOTTO DI ENTRATA** restano vuote: non sono ricavabili dal DDT in uscita.

## Pesate (DR.CPF.70)

Un file per formato di vaso, che continua per tutto l'anno. Ogni prodotto del DDT diventa
un foglio `numero_ddmmyy` nel file del suo formato; se un DDT ha piu' prodotti dello stesso
formato, il secondo e i successivi prendono `-A`, `-B`. I fogli sono in ordine di numero
di DDT.

Di ogni foglio cambiano **solo tre celle**: B5 la data, E5 il lotto, F5 il nome del
prodotto, uguale a quello del Riepilogo produzione. Tutto il resto viene ricopiato.

**Le pesate vengono fissate come numeri interi.** RANDBETWEEN e' volatile: finche' resta
una formula i valori cambiano a ogni apertura del file, anche nei fogli di mesi prima. Il
programma li sorteggia con la stessa regola — fra il minimo e il massimo del formato — e li
scrive come numeri, diversi per ogni foglio. Medie, somme e verifiche restano formule e si
ricalcolano da se'. La prima volta che tocca un file, fissa anche le pesate dei fogli gia'
presenti ai valori che avevano.

Le celle hanno formato `0`, quindi **una pesata vale quello che si legge**: la formula
originale dava i millesimi (234,631) che si vedevano solo entrando nella cella, mentre il
foglio mostrava 235. Ora nella cella c'e' 235. I file gia' scritti sono stati convertiti
allo stesso modo, arrotondando al numero che gia' mostravano: a video non e' cambiato
niente.

Sui file toccati **viene tolta la protezione del foglio**, e i nomi vengono ripuliti: via i
caratteri di controllo (`_x0001_113_080526`) e i trattini bassi di troppo (`179_270726_`),
e le date di sette cifre riportate a sei (`57_1200326` diventa `57_120326`).

I formati **140 e 190** non hanno la scheda vuota, perche' i file di partenza erano gusci
senza formule ne' parametri. Se un giorno li produci, il programma lo scrive nel log e non
crea la pesata.

## Registri di accettazione e giacenze (acquisti)

Ogni riga della fattura viene classificata e mandata dove le compete:

| Riga della fattura | Registro | Giacenze |
|---|---|---|
| FILETTI ACCIUGHE / SARDE | DR.CAM, una riga per lotto | PESCE FAO xx, SARDE FAO 34, ALTRO PESCE |
| OLIO | DR.CAO, un foglio per lotto | OLIO DI SEMI, OLIO D'OLIVA, OLIO EVO |
| VASO / CAPSULE | DR.CAI, un foglio per articolo | VASI xxx, CAPSULE xx |
| VASSOI, ETICHETTE, SIGILLI, SCATOLE, PELLICOLA | ignorate in silenzio | ignorate |
| TONNO | ignorato ma segnalato nel log | ignorato |

Il tonno viene lasciato fuori **anche dal riepilogo produzione e dalle pesate**, per lo
stesso motivo: l'albero dei materiali e' quello delle acciughe e non gli si adatta. Se un
DDT contiene solo tonno non viene creato nessun foglio e la segnalazione arriva una volta
sola, come per le fatture.

Il riconoscimento parte dalla **prima parola** della descrizione, perche' le parole
chiave compaiono anche dentro descrizioni di tutt'altro genere: `VASO VETRO 160 ETICHETTA
LINEARE PER ONDINA T45` e `VASO VETRO FILETTI ACCIUGHE 450 T82` sono vasi, non etichette
ne' pesce. Per lo stesso motivo i vasi vuoti destinati al tonno si registrano
normalmente.

Se la riga non indica i pezzi per collo ma dice `Unita' singola`, la quantita' e' gia' il
numero di pezzi e viene usata cosi' com'e'.

Kg del pesce = quantita' x peso dell'unita' indicata fra i doppi trattini (CARTONE 16KG,
BUSTA 2KG, SECCHIELLO 15KG). Pezzi degli imballaggi = quantita' x pezzi per collo
(PALLET DA 1694 PZ, CARTONI DA 3300 PZ).
Quando in una riga compaiono piu' lotti si usa il primo.

Quando un documento produce piu' fogli, il secondo e i successivi prendono il suffisso
`-A`, `-B`, `-C`, che compare anche nella cella "Estremi del DDT".

### Il registro giacenze

La scheda **RIMANENZE non viene mai toccata**: il programma aggiunge solo righe alle schede
dei materiali, **in ordine di data**.

Ogni riga porta in **colonna F un identificativo** del tipo `#121560-2`: documento e
posizione della voce dentro il documento. Serve a due cose:

- una riga **corretta a mano** resta riconosciuta, quindi non viene riaggiunta come doppione
- due righe **davvero identiche** — due consegne uguali nello stesso documento — convivono,
  perche' hanno identificativi diversi

Le righe scritte prima che esistessero gli identificativi sono state marchiate
riconoscendole dal contenuto; quelle piu' vecchie, di cui non ci sono i documenti, restano
senza e non danno fastidio.

## Note tecniche

`HACCP/FORMATI E CAPSULE.xlsx` e' la tabella formato -> capsula e coefficiente: si aggiorna
li', senza toccare il programma. Ce n'e' una sola copia, apposta.

Nel DR.CAO la parte variabile di "Estremi del DDT" e "Descrizione prodotto" e'
sottolineata, e nelle pesate le intestazioni hanno il grassetto su una parola sola: e'
testo formattato a tratti dentro la cella, quindi i file vanno sempre aperti con
`rich_text=True`. Su questo openpyxl ha un difetto — scrive gli spazi senza dichiarare di
volerli conservare, ed Excel chiede di riparare il file — a cui il programma rimedia dopo
ogni salvataggio: vedi `salva_libro` in `src/riepilogo.py`.

Da terminale, per lavorare su cartelle diverse:

```
.venv/bin/python src/riepilogo.py "PDF Vendite" -o "PROVA/{anno}/DR.RPD.00 - Riepilogo produzione {anno}.xlsx"
.venv/bin/python src/pesate.py    "PDF Vendite" -o "PROVA/{anno}/PESATE"
.venv/bin/python src/acquisti.py  "PDF Acquisti" -o "PROVA/{anno}" -g "PROVA/REGISTRO GIACENZE {anno}.xlsx"
```

Opzioni comuni: `-l` il file di log, `-b` la cartella delle copie `.bak`, `--sovrascrivi`
per rifare da capo i fogli gia' presenti invece di saltarli.

`pesate.py --sistema` passa su tutti i file di una cartella di pesate anche se oggi non
ricevono niente: toglie la protezione, fissa le pesate rimaste a formula, uniforma il nome
della scheda di partenza e rimette i fogli in ordine di DDT. Serve quando si parte da file
compilati altrove.

Le librerie sono gia' installate in `.venv`. Se un giorno quella cartella sparisce:

```
python3 -m venv .venv && .venv/bin/pip install openpyxl pymupdf
```
