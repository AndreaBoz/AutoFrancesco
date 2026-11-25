# AutoFrancesco Esterni

Come AutoFrancesco, ma per i clienti e i fornitori **diversi da FLOTT**. Scrive
nell'archivio `~/HACCP/ESTERNI` (configurabile tramite `HACCP_DIR`).

| Da | A |
|---|---|
| `EST PDF Vendite` — i DDT verso clienti esterni (Etnatost, Sud Ingross…) | EST-Riepilogo produzione |
| `EST PDF Vendite` | DR.CPF.70 pesate, un file per formato, in `{anno}/PESATE` |
| `EST PDF Acquisti` — le fatture dei fornitori | EST-Controllo MAT.PRIME / OLIO / IMBALLAGGI e EST-REGISTRO GIACENZE |

## Come si usa

1. Scarica i PDF nuovi e mettili in `EST PDF Vendite` e `EST PDF Acquisti`. **I vecchi si
   tengono**: rileggerli costa pochi millisecondi e permette di rifare tutto da capo.
2. Doppio clic su **genera esterni.command**.
3. Se a fine esecuzione compare il riquadro di avviso, apri il log in `Log/` e guarda cosa
   e' rimasto fuori. Se non compare, non c'e' niente da controllare.

Vale tutto quello che vale per FLOTT: si aggiunge soltanto, i documenti gia' inseriti
vengono riconosciuti e saltati, le segnalazioni si vedono una volta sola, il terminale
elenca solo cio' che e' cambiato. **Prova/prova a vuoto.command** rifa' le stesse cose
dentro `Prova/Generati` senza toccare l'archivio.

I programmi sono quelli di **AutoFrancesco**, che deve stare in questa stessa cartella
`Progetti`: qui dentro ci sono solo i PDF, i log e le copie di sicurezza. Anche la tabella
`FORMATI E CAPSULE.xlsx` e' una sola, quella dentro HACCP.

## Cosa cambia rispetto a FLOTT

**I DDT di vendita** sono impaginati allo stesso modo, ma ogni riga comincia col codice
del cliente e quel codice entra nel nome del prodotto:

```
sg) vasi di filetti di acc. osg da gr. 720 conf. da 6 L 26187 sc. 06/01/2028 anonimo zona fao 34
        ->  sg) Vasi Fil. di acc. osg gr. 720 anonimo zona fao 34
```

La parentesi ogni tanto manca (`et vasi di…`) e ogni tanto manca lo spazio
(`pas)Vasi di…`): il programma se ne accorge lo stesso. I formati grandi sono scritti in
chili (`da kg. 1,5`) e valgono 1500 grammi. Nelle **pesate** il codice del cliente non
compare, come nelle schede compilate a mano.

Il **tonno** non viene registrato ne' nel riepilogo ne' nelle pesate: l'albero dei
materiali e' quello delle acciughe e non gli si adatta. Viene segnalato una volta sola e
lo si compila a mano, come si e' sempre fatto con le fatture di acquisto.

**Le fatture di acquisto** sono la parte davvero diversa: ogni fornitore le impagina a
modo suo e il programma le legge tutte.

| Fornitore | Come scrive il lotto | Come scrive la quantita' |
|---|---|---|
| Fornitore tipo A (es. Conserve Ittiche) | `Tipo dato: LOTTO` + `Rif. testo:` | gia' totale, con UM `PZ` o `KG` |
| Fornitore tipo B (layout con `lotto:` sotto l'articolo) | `lotto:` nella riga sotto l'articolo | gia' totale |
| SICIL CONSERVIT | non lo scrive | `- Nr. 2 x 2.550`, in fattura c'e' gia' il totale |
| AGATA PESCA | non lo scrive | chili di alici fresche |

Il **fornitore vero** finisce nella casella `Fornitore:` di ogni scheda, al posto della
dicitura del modello.

Le **alici fresche** di Agata Pesca non portano scritta la zona di pescato: sono pescato
locale e vanno in **FAO 37**.

I numeri di fattura con la barra perdono la parte dopo: `317/25` diventa `317`, `175/A`
diventa `175`.

Le misure dei vasi e delle capsule si leggono anche quando sono scritte alla rovescia:
in `VASO ANFORA GR. 80 TP 45` l'80 e' il vaso e il 45 e' la capsula; in
`CAPS. TO DIAM. 110` il 110 e' la capsula. Il `COPERCHIO PER VASO … DA 1500` e' la capsula
del formato 1500, cioe' la 154, e la tabella dei formati lo sa. L'orcio da 314 ML si
registra nella scheda **VASI 300**, come il vaso da 40 ML si registra nella **VASI 42**.

## Com'e' fatta la cartella

```
AutoFrancesco Esterni/
├── genera esterni.command   scrive nell'archivio HACCP/ESTERNI
├── EST PDF Vendite/         i DDT da lavorare
├── EST PDF Acquisti/        le fatture da lavorare
├── Log/                     un log e una memoria per ogni programma
├── Backup/                  la copia .bak di ogni file prima di modificarlo
├── Esempi/                  come erano compilati a mano, per confronto
└── Prova/prova a vuoto.command
```

E nell'archivio:

```
HACCP/ESTERNI/
├── EST-REGISTRO GIACENZE 2026.xlsx      l'anno in corso sta qui
├── 2025/                                gli anni chiusi, coi loro trimestri
└── 2026/
    ├── EST-Riepilogo produzione 2026.xlsx        il trimestre in corso sta qui
    ├── EST-Controllo in accettazione MAT.PRIME e SEMILAVORATI.xlsx
    ├── EST-Controllo in accettazione OLIO.xlsx
    ├── EST-Controllo imballaggi primari.xlsx
    ├── PESATE/                          un DR.CPF.70 per formato
    └── 1 TRIMESTRE/ 2 TRIMESTRE/ …      i trimestri gia' chiusi
```

## A inizio trimestre e a inizio anno

Esattamente come per FLOTT: a inizio trimestre sposti i quattro registri dentro
`<n> TRIMESTRE` e il programma ne ricrea di nuovi nella radice dell'anno; a inizio anno
sposti `EST-REGISTRO GIACENZE` nella cartella dell'anno appena chiuso e prepari quello
nuovo nella radice di `ESTERNI`. Se un documento di dicembre lo lavori a gennaio, finisce
comunque nel file del suo anno: il registro giacenze dell'anno chiuso viene cercato prima
nella radice e poi nella cartella dell'anno.

## Le pesate

I 16 file `DR.CPF.70 … gr NNN (FORMULA).xlsx` in `2026/PESATE` continuano nel tempo: di
ogni foglio nuovo cambiano solo data, lotto e nome del prodotto, e le pesate vengono
risorteggiate con la stessa formula del foglio da cui si copia. Le pesate dei fogli gia'
scritti vengono **fissate a numero intero**: la `RANDBETWEEN` e' volatile e altrimenti se
le rigenererebbe a ogni apertura di Excel, cambiando i controlli gia' firmati. Le celle
hanno formato `0`, quindi una pesata vale quello che si legge sul foglio: i millesimi della
formula si vedevano solo entrando nella cella. La protezione dei fogli viene tolta e i
fogli restano in ordine di numero di DDT.

Il formato **190** non ha una scheda da cui partire: se un giorno arrivera' un DDT da 190
grammi il programma lo dira' nel log e la scheda andra' creata a mano. Gli altri quindici
hanno il loro modello in `AutoFrancesco/src/Modelli pesate esterni/`.

Se dovessero arrivare file di pesate compilati altrove, ancora protetti o con le formule
vive, si rimettono in riga con:

```bash
"AutoFrancesco/.venv/bin/python" "AutoFrancesco/src/pesate.py" --sistema -o "$HOME/HACCP/ESTERNI/{anno}/PESATE" -m "AutoFrancesco/src/Modelli pesate esterni"
```

## Quello che il programma non compila

Come per FLOTT, `ORIGINE` e `LOTTO DI ENTRATA` del riepilogo produzione restano vuote: dal
DDT non si ricavano, le incroci tu coi registri di accettazione. Vassoi, etichette,
sigilli, scatole, pellicole, interfalde e bancali vengono ignorati in silenzio; il tonno
viene segnalato.
