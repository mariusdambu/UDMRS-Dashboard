# Manual UDMRS Dashboard

Acest manual descrie pachetul public actual. Este scris pentru folosire normala, fara istoria interna a proiectului.

Release stabila: `UDMRS-Hydra-TwoHeads-Stable-v1.0-20260618`.

Build: `UDMRS Build 2026.06.18-H2H-v1.0`.

Validare finala: sintaxa PS5.1/PS7, dashboard, provideri disponibili si fluxuri principale revizuite fara Apply pe galerii reale.

Stare publica actuala: fluxul clasic stabil ramane intrarea principala pentru foldere normale. `Importa galerie` este disponibila pentru `Google Photos / Takeout`, `Apple Photos / iCloud` si `XMP / Sidecar Library`.

## 1. Ce este fiecare componenta

- `PhotoOrganizer.ps1`: motorul PowerShell. Face lucrul real.
- `PhotoOrganizerDashboard.ps1`: interfata grafica. Lanseaza motorul cu parametri si arata progresul.
- `Start-PhotoOrganizer.cmd`: metoda recomandata de pornire.
- `ProcessedFiles.json`: index intern incremental. Pastreaza hash-uri si rute cunoscute.

Dashboard-ul nu dubleaza logica motorului.

## 2. Pornire

Ruleaza:

```text
Start-PhotoOrganizer.cmd
```

Portabilitate: porneste intotdeauna fisierul `.cmd` din folderul curent al UDMRS Dashboard. Un shortcut `.lnk` copiat din alt folder poate continua sa indice ruta veche, deoarece Windows salveaza rute absolute in shortcut. Cand pornesti `Start-PhotoOrganizer.cmd`, lansatorul actualizeaza shortcut-ul local daca exista.

Fluxul oficial este:

```text
Pornire
↓
Alegerea sursei si destinatiei
↓
Alegerea modului
↓
Verificarea rezumatului dinamic
↓
Executare / Anulare / Test scan
↓
Progres
↓
Loguri si rezumat
↓
Rezultate
↓
Setari
```

Pentru utilizare normala lucreaza din `Pornire`.

## Referinta oficiala de comenzi

Comenzile manuale gata de copiat sunt in:

```text
Docs\CommandReference.html
```

Acest fisier foloseste sablon portabil:

```text
Script: <CarpetaUDMRS>\App\PhotoOrganizer.ps1
Source: %USERPROFILE%\OneDrive\Imágenes
Destination ES: %USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas
Destination RO: %USERPROFILE%\OneDrive\Imágenes\Poze_Organizate
Language ES: es
Language RO: ro
```

Nu documenteaza parametri inexistenti precum `-Cleanup`.

## 3. Pornire si rezumat dinamic

In `Pornire` alegi:

- folderul sursa
- folderul destinatie
- `DryRun`
- aplicarea modificarilor reale
- repararea EXIF
- modul de performanta
- numarul maxim de workeri, daca vrei sa limitezi calculatorul

Textul de langa butonul de executare se schimba dupa optiunile selectate. Acela este rezumatul actiunii care va fi lansata.

## 4. Actiuni principale

### Executare

Foloseste configuratia curenta.

- Cu `DryRun`: analizeaza si simuleaza. Nu atinge fisiere.
- Cu modificari reale: muta sau copiaza dupa configuratie.
- Cu reparare EXIF: poate repara metadate daca modificarile reale sunt active.

### Anulare

Opreste procesul PowerShell copil. Motorul scrie anularea in log si pastreaza informatia necesara pentru verificare.

### Test scan

Este sigur.

Face:

- scaneaza fisiere compatibile
- aplica excluderi
- valideaza accesul
- arata fisiere gasite, omise si inaccesibile

Nu face:

- nu citeste EXIF
- nu calculeaza hash-uri
- nu muta
- nu copiaza
- nu modifica EXIF
- nu creeaza structura finala

Foloseste-l inainte de biblioteci mari pentru a confirma sursa.

## 4.1 Importa galerie

`Importa galerie` este o intrare principala pentru exporturi de provider care contin mai mult context decat fisiere media simple: albume, sidecar-uri, JSON, cos sau relatii intre elemente.

Disponibil:

- `Google Photos / Takeout`: analizeaza o exportare Google Takeout aleasa de utilizator, interpreteaza sidecar-uri JSON, albume, cos, incredere metadata si duplicate fizice, apoi copiaza asset-uri curate in destinatia `An\Trimestru`.
- `Apple Photos / iCloud`: analizeaza o exportare iCloud aleasa de utilizator, interpreteaza `Photo Details.csv`, CSV-uri de albume, marcaje de cos, date provider, video si candidati Live Photo, apoi copiaza asset-uri curate in destinatia `An\Trimestru`.
- `XMP / Sidecar Library`: incearca sa interpreteze galerii necunoscute cu sidecar-uri XMP, JSON sau YAML. Daca relatia media/sidecar este clara foloseste metadata; daca este ambigua trimite la revizuire; fara metadata utila foloseste fallback clasic.

Planificate si dezactivate:

- `Samsung Gallery`: blocat pana exista o exportare reala.
- `Immich`: necesita mostre/exporturi potrivite.

Pentru orice provider, trateaza folderul exportat ca sursa temporara: ruleaza mai intai Simulare, verifica raportul HTML/JSON si apoi Apply. Dupa un Apply reusit, dashboard-ul poate cere confirmare pentru stergerea folderului de export selectat. Nu se sterge in Simulare si nu se sterge daca importul esueaza.

Pentru XMP / Sidecar Library, raportul arata media gasite, sidecar-uri gasite/folosite, sidecar-uri orfane, media fara sidecar, incredere High/Medium/Low, conflicte si fallback clasic.

Pentru foldere normale cu poze si video foloseste `Pornire` / `Organize`.

## 4.2 Optiuni avansate / Mod expert

In `Setari`, butonul `Mod expert` deschide instrumente tehnice fara sa le amestece cu fluxul normal din `Pornire`.

Nu dubleaza `Executare`, `Test scan`, `Sincronizeaza indexul` sau `Elimina intrarile disparute` atunci cand acele actiuni au deja buton vizibil.

Instrumente disponibile:

- `Curatare retentie`: curata doar backup-uri EXIF vechi si duplicate confirmate vechi. Nu curata galeria normala, nu curata foldere goale si nu atinge folderul de duplicate in asteptarea revizuirii.
- `Recupereaza duplicate mutate gresit`: incearca sa recupereze fisiere trimise gresit la duplicate folosind loguri si indexul.
- `Tradu foldere existente`: adapteaza foldere trimestriale existente la limba activa.
- `Tradu foldere interne`: redenumeste sau consolideaza containere interne cunoscute in limba activa fara sa citeasca poze sau sa descarce din cloud.
- `Normalizeaza structura`: restructureaza vizual la `An\Trimestru` si curata ramuri goale/junk/reziduale.
- `Curatare duplicate`: analizeaza duplicate exacte dupa hash si in Apply muta confirmatele in carantina.
- `Repara EXIF in-place`: repara EXIF in biblioteca organizata fara sa mute, deduplice sau redenumeasca.
- `Auditeaza datele vizibile`: detecteaza fisiere cu data fiabila cunoscuta, dar metadate vizibile sau date de sistem lipsa/incoerente. Este intotdeauna simulare.
- `Repara datele vizibile`: scrie metadate vizibile si date de sistem atunci cand lipseste o data de captura utila. Necesita Apply si confirmare specifica.
- `Migreaza UDMRS pe alt PC`: creeaza ZIP-urile necesare pentru a muta instalarea partajata si starea utilizatorului curent pe alt calculator. Nu migreaza loguri sau runtime.

Toate respecta `Simulare` implicit. Daca activezi modificari reale, dashboard-ul cere confirmare specifica si arata switch-urile care vor fi lansate.

Cand lansezi un instrument din `Mod expert`, dashboard-ul deschide o consola tehnica persistenta. Consola foloseste PowerShell 7 daca exista si este reutilizata pentru urmatoarele actiuni avansate. Nu se inchide dupa fiecare actiune, ca sa poti verifica iesirea, erorile si progresul fara sa pierzi contextul. Dashboard-ul principal ramane utilizabil.

Consola tehnica apartine sesiunii dashboard-ului. Daca inchizi dashboard-ul, se inchide si consola tehnica controlata, procesele copil si starea temporara din `%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole`. Astfel se evita ferestre PowerShell ramase deschise, lock-uri false si stari `Running` vechi.

## 5. DryRun si Apply

`DryRun` este modul de simulare. Nu scrie schimbari reale.

Aplicarea modificarilor reale permite modificarea bibliotecii dupa o confirmare initiala. Nu exista confirmari pentru fiecare fisier.

`DryRun + reparare EXIF` doar raporteaza posibile reparatii; nu modifica metadate.

Pentru prima trecere reala pe o biblioteca mare sau neprocesata se recomanda:

```text
modificari reale + reparare EXIF
```

Datele EXIF reparate imbunatatesc organizarea de la inceput.

## 6. PowerShell

Dashboard-ul detecteaza automat `pwsh.exe`.

- Daca exista PowerShell 7, il foloseste.
- Daca nu exista, foloseste Windows PowerShell 5.1.

PowerShell 7 este recomandat pentru biblioteci mari.

## 7. Organizarea actuala

Profilul oficial unic este `QuarterlyFolders`.

Structura matura a RC este:

```text
An
↓
Trimestru
```

Modelul structural anterior este retras. Nu mai exista doua filosofii active de organizare.

Exemplu in romana:

```text
Poze_Organizate
  2025
    Ian-Mar
    Apr-Iun
    Iul-Sep
    Oct-Dec
```

In trimestru, numele fisierelor sunt pastrate conservator. `NormalizeExistingFolders` nu trebuie sa inventeze evenimente si nu redenumeste masiv fisiere folosind EXIF.

Responsabilitati:

- `Organize`: ingestie, clasificare, deduplicare, reparare EXIF optionala si index.
- `NormalizeExistingFolders`: restructurare vizuala si curatare.
- `ReconcileProcessedDatabase`: sincronizarea fina a `ProcessedFiles.json`.
- `PurgeMissingFromProcessedDatabase`: curatare avansata a istoricului disparut.
- `DedupeCleanup`: duplicate dupa hash; nu curata foldere goale.

## 8. Foldere interne

Cu limba romana activa, folderele interne principale sunt:

```text
Poze_Organizate
_De_Revizuit
_Duplicate_De_Revizuit
_Carantina_Duplicate_Confirmate
_Backup_Metadate
Logs
MediaMetadataIssues
```

Optiunea de traducere a folderelor interne afecteaza doar containere interne cunoscute ale proiectului. Poate redenumi sau consolida aliasuri istorice precum `Poze_Organizate`, `Fotos_Organizadas`, `Organized_Photos`, `_NeedsReview`, `_NecesitaRevision`, `_Duplicates_To_Review`, `_Duplicados_Para_Revisar`, `_MetadataBackup`, `_CopiaSeguridadMetadatos` sau carantine vechi catre numele oficial al limbii active.

In biblioteci cloud-backed, aceasta redenumire este o operatie structurala de folder. Nu calculeaza hash, nu citeste EXIF, nu enumera profund fisierele interne si nu trebuie sa hidrateze placeholders cloud-only. In `DryRun` arata containerele pe care le-ar redenumi sau consolida; in `Apply` redenumeste containerul cunoscut sau muta doar copiii imediati daca destinatia oficiala exista deja si nu exista conflicte de nume.

`RenameExistingFoldersToCurrentLanguage` se limiteaza la foldere trimestriale existente, de exemplu `Jan-Mar` la `Ian-Mar`, si lucreaza tot la nivel de folder fara sa citeasca continut.

Dupa redenumiri structurale, Explorer si OneDrive pot avea nevoie de cateva minute pentru refresh, reindexare sau sincronizare. Poate parea ca Explorer nu raspunde, mai ales in biblioteci mari, foldere online-only sau dupa mai multe redenumiri. Daca logul confirma ca nu a existat content scan, EXIF, hash sau cloud traversal, acest comportament Explorer/OneDrive nu inseamna corupere, pierdere de date sau hidratare masiva. Asteapta cateva minute inainte de alta operatie. Daca Explorer ramane blocat prea mult timp, reporneste Explorer sau Windows.

Folderele excluse de utilizator sunt sarite inainte de orice citire de continut. `Samsung Gallery` si `Camera Roll` pot fi folosite ca preset-uri sugerate, dar nu sunt excluderi active globale: fiecare utilizator decide din dashboard ce foldere protejeaza.

## 8.1 Foldere excluse / protejate externe

In `Pornire` exista un panou vizibil cu rezumatul folderelor protejate active, gasite si negasite. De acolo poti deschide direct gestionarea excluderilor.

In `Setari` poti gestiona si:

```text
Foldere excluse / Foldere externe protejate
```

Aceste foldere sunt teritoriu interzis pentru aplicatie:

- nu sunt organizate
- nu sunt normalizate
- nu sunt reparate EXIF
- nu sunt deduplicate
- nu sunt purged
- nu sunt curatate ca foldere goale/junk/reziduale
- nu se calculeaza hash
- nu se citeste EXIF
- nu se hidrateaza daca sunt cloud-only

Concepte:

```text
InternalProtectedFolders
= foldere proprii ale aplicatiei

UserExcludedFolders
= foldere pe care utilizatorul decide ca aplicatia nu trebuie sa le atinga

VendorManagedFolders
= preseturi/ajutoare pentru foldere tipice ale aplicatiilor externe
```

Configuratia se salveaza portabil in:

```text
%APPDATA%\PhotoOrganizer\Config\UserExcludedFolders.json
```

Poate folosi rute absolute sau sabloane precum `%USERPROFILE%`, `<SourcePath>` si `<DestinationPath>`. Daca o ruta nu exista, nu este stearsa automat: ramane in lista si este marcata ca negasita.

Daca fisierul de configuratie nu exista, UDMRS il creeaza automat in `%APPDATA%\PhotoOrganizer\Config\UserExcludedFolders.json` cu o lista goala de excluderi active. Instalatia partajata nu defineste excluderi reale pentru toti utilizatorii.

## 9. Repararea indexului: ReconcileProcessedDatabase

In dashboard apare ca sincronizare/reparare a indexului.

Este pentru miscari manuale sau accidentale:

- foldere trase din greseala
- fisiere mutate in afara `Poze_Organizate`
- foldere temporare precum `Primavara 2005`, `De revizuit` sau `Conflicte reparate EXIF`

Comportament actual:

- nu mosteneste excluderile normale gandite doar pentru Organize, dar respecta folderele excluse de utilizator si preset-urile externe protejate
- vede `Poze_Organizate`
- valideaza rutele din `ProcessedFiles.json`
- daca ruta exista, pastreaza intrarea
- daca ruta veche nu exista, cauta acelasi hash in aria actuala
- daca gaseste o singura ruta noua, actualizeaza indexul
- daca fisierul este in afara bibliotecii organizate, il poate marca `ManualMoved/ExternalLocation`
- daca nu apare, il marcheaza `Missing/Stale`
- daca exista mai multe rute cu acelasi hash, raporteaza conflict

Nu muta fotografii. Nu modifica EXIF. Nu reorganizeaza foldere.

In `DryRun` doar raporteaza. In Apply face backup al JSON-ului intern inainte de salvare.

## 10. Curatarea intrarilor disparute

`PurgeMissingFromProcessedDatabase` este actiune avansata, nu flux normal.

Inainte de stergerea intrarilor, revalideaza fiecare intrare disparuta:

- daca ruta exista din nou, o recupereaza
- daca lipseste in continuare, poate sterge intrarea din JSON in Apply
- nu sterge fotografii reale
- nu muta fisiere
- nu modifica EXIF

Foloseste-o doar dupa reconciliere, cand vrei sa cureti istoricul fisierelor care chiar nu mai exista.

Daca nu exista intrari disparute reale, logul trebuie sa spuna clar ca nu este nimic de curatat.

## 11. NormalizeExistingFolders

`NormalizeExistingFolders` este optional pentru normalizarea vizuala a unei biblioteci deja organizate.

Rol nou:

```text
NormalizeExistingFolders -Apply
↓
muta/redenumeste in An\Trimestru
↓
curata foldere goale
↓
curata foldere junk-only
↓
curata ramuri reziduale
↓
intra in validare/reconcile post-operatie
```

Important: Normalize nu mai trebuie inteles ca responsabil principal pentru mentinerea perfecta a `ProcessedFiles.json` dupa fiecare mutare. Adevarul final al indexului apartine lui `ReconcileProcessedDatabase`.

Dupa `Normalize Apply`, motorul intra intr-o faza separata de validare/reconcile pentru a detecta rute stale si a repara indexul. Pentru audit sau repararea miscarilor manuale, foloseste `ReconcileProcessedDatabase`.

Ruleaza intotdeauna mai intai `DryRun` si verifica raportul.

In biblioteci OneDrive sau alti furnizori cloud-backed, `NormalizeExistingFolders -Apply` poate produce multa activitate Explorer/sincronizare deoarece muta si redenumeste structura. Daca Explorer ramane temporar fara refresh sau `Nu raspunde`, de obicei reindexeaza schimbarile. Motorul nu trebuie sa citeasca EXIF sau sa calculeze hash-uri suplimentare doar din cauza acestei asteptari vizuale. Lasa sincronizarea sa se stabilizeze inainte de o alta operatie.


## 11.1 Date vizibile: MetadataAudit si MetadataRepair

`NormalizeExistingFolders` repara structura. `MetadataRepair` repara data vizibila din fisier.

UDMRS poate cunoaste o data fiabila din EXIF, provider, sidecar sau model de nume. Acea data trebuie sa fie coerenta in:

- ruta/folder
- numele final
- metadata incorporata atunci cand formatul permite
- `CreationTime` / `LastWriteTime` cand sunt date accidentale
- `ProcessedFiles.json` si rapoarte

`MetadataAudit` verifica biblioteca organizata si genereaza un CSV cu candidati. Nu modifica nimic, chiar daca dashboard-ul este in Apply. Foloseste-l ca sa vezi cate fisiere au data fiabila cunoscuta, dar nu vizibila pentru Windows, OneDrive sau Microsoft Photos.

`MetadataRepair` actioneaza doar asupra candidatilor siguri. Creeaza backup, scrie metadata incorporata in functie de format, sincronizeaza datele de sistem cand este cazul, recalculeaza hash-ul si actualizeaza indexul.

Formate acoperite de politica de materializare: JPG/JPEG, HEIC/HEIF, MP4/MOV/M4V/3GP, PNG, TIFF, WEBP si GIF atunci cand exista o metoda sigura de scriere a datei utile. Daca formatul nu permite data asteptata sau exista conflict cu metadata valida, logul/raportul trebuie sa marcheze `DateKnownButMetadataNotWritten` sau un warning echivalent.

Reguli importante:

- nu suprascrie metadata valida existenta
- nu rezolva automat conflicte
- nu muta si nu reorganizeaza foldere
- nu deduplica
- nu inlocuieste Reconcile
- poate dura mult in biblioteci mari deoarece citeste metadata in-place

Exemplu DryRun:

``powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" `
  -SourcePath "%USERPROFILE%\OneDrive\Imágenes" `
  -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" `
  -MetadataAudit `
  -Language ro
``

Exemplu Apply:

``powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" `
  -SourcePath "%USERPROFILE%\OneDrive\Imágenes" `
  -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" `
  -MetadataRepair `
  -Language ro `
  -Apply
``
## 12. ProcessedFiles.json

Ruta obisnuita:

```text
%APPDATA%\PhotoOrganizer\ProcessedFiles.json
```

Pastreaza:

- hash
- ruta relativa originala sau noua
- stare
- `lastSeen`
- `missingSince`
- versiunea uneltei

Stari relevante:

- intrare activa: poate alimenta indexul de duplicate
- `ManualMoved/ExternalLocation`: fisier gasit in afara locatiei organizate; nu se foloseste ca duplicat canonic automat
- `Missing/Stale`: ruta inregistrata nu exista acum

Hash-ul este identitatea principala. Ruta este actualizabila.

## 13. Curatare sigura

Motorul poate elimina foldere goale sau foldere care contin doar fisiere auxiliare sigure.

Curatarea sigura se executa in cadrul `NormalizeExistingFolders`. Nu exista un parametru separat `-Cleanup`.

Important:

```text
-KeepEmptyFolders NU curata foldere goale.
-KeepEmptyFolders pastreaza foldere goale.
```

Pentru a permite curatarea in timpul Normalize foloseste:

```powershell
-NormalizeExistingFolders `
-KeepEmptyFolders:$false
```

Exemple:

```text
desktop.ini
Thumbs.db
ehthumbs.db
.DS_Store
.picasa.ini
.dropbox
._*
Icon?
.Trashes
.fseventsd
.Spotlight-V100
*.tmp
```

Poate elimina si foldere cu markere mici sigure, daca nu exista fisiere potential utile.

Nu considera gunoi fisiere cu extensii de fotografii, video, RAW, JSON, XMP, TXT, PDF, DOC, linkuri, HTML, XML sau CSV.

Dupa Normalize, curatarea incearca sa elimine ramuri reziduale goale de jos in sus.

## 13.1 RetentionCleanup

`RetentionCleanup` este curatare de retentie, nu curatare generala a galeriei.

Nu face:

- nu curata foldere goale
- nu ruleaza Normalize
- nu reorganizeaza poze
- nu purge `ProcessedFiles.json`
- nu cauta duplicate noi
- nu atinge folderul de duplicate in asteptarea revizuirii

Poate curata doar continut temporar sau confirmat:

```text
%LOCALAPPDATA%\PhotoOrganizer\_Backup_Metadate
_Carantina_Duplicate_Confirmate
```

Recunoaste si aliasuri din alte limbi, cum ar fi `_CopiaSeguridadMetadatos`, `_MetadataBackup`, `_Cuarentena_Duplicados_Confirmados` sau `_Confirmed_Duplicates_Quarantine`.

Folderul de duplicate in asteptarea revizuirii nu este atins niciodata de `RetentionCleanup`, deoarece poate contine variante, cazuri ambigue sau fisiere care cer verificare umana.

Rezultatul `Nu s-a sters nimic` poate fi complet normal. Implicit backup-urile EXIF se pastreaza 30 de zile si duplicatele confirmate 45 de zile. Daca nu exista inca elemente suficient de vechi, nu vor exista candidati. Poti astepta si rula din nou mai tarziu sau poti sterge manual continut confirmat daca ai nevoie imediata de spatiu si accepti sa pierzi acea fereastra de recuperare.

## 14. Diagnostic EXIF lent

Daca un lot EXIF dureaza prea mult, motorul pastreaza fallback-ul actual:

```text
timeout lot
↓
mod individual
```

In plus scrie diagnostic:

```text
Slow EXIF candidate: Path=...; Extension=...; Size=...; Time=...; Batch=...; Reason=...; Count=...
```

Contoare:

- `slowExifCandidates`: fisiere unice candidate la EXIF lent
- `slowExifDetections`: total detectii

Motive tipice:

- `Batch timeout/fallback`
- `timeout`
- `slow per-file read`
- `metadata warning`
- `fallback`

Un timeout de lot marcheaza candidatii din lot. Fisierul cu adevarat suspect apare de obicei din nou ca timeout individual, citire lenta sau warning de metadata.

Cautare utila:

```powershell
Select-String -Path "%APPDATA%\PhotoOrganizer\Logs\*.log" -Pattern "Slow EXIF candidate"
```

## 15. Loguri si progress.json

Fila `Loguri si rezumat` arata iesirea recenta si permite deschiderea folderului `Logs` chiar in timpul rularii.

Fisiere obisnuite:

```text
PhotoOrganizer-YYYYMMDD-HHMMSS.log
PhotoOrganizer-YYYYMMDD-HHMMSS.progress.json
```

Dashboard-ul ignora silentios citiri sau scrieri blocate temporar, ca sa nu arate exceptii PowerShell brute.

La pornire, dashboard-ul executa o curatare automata si silentioasa a logurilor vechi din `%APPDATA%\PhotoOrganizer\Logs`. Retentia aproximativa este de 7 zile pentru `*.log`, `*.progress.json` si rapoarte HTML. Aceasta curatare nu atinge poze, videoclipuri, `ProcessedFiles.json` real, backup-uri EXIF sau carantine; de asemenea nu atinge `%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole`.

Backup-urile tehnice pentru `ProcessedFiles.json` traiesc separat in `%APPDATA%\PhotoOrganizer\IndexBackups`. Nu sunt loguri normale. UDMRS pastreaza intotdeauna cel mai recent backup, pastreaza backup-uri recente pana la un maxim aproximativ de 10 copii si sterge restul. Daca exista backup-uri vechi in `Logs\JsonBackups`, sunt migrate automat in `IndexBackups`.

Ca masura conservatoare, nu sterge fisiere modificate in ultima ora, protejeaza loguri/progress ale executarilor active cand le poate detecta si pastreaza intotdeauna cel putin cel mai recent backup JSON. Daca un fisier este blocat sau exista probleme de permisiuni, il omite si dashboard-ul continua sa porneasca.

`progress.json` poate include:

- stare
- faza
- PID
- analizate
- mutate
- duplicate
- EXIF reparate
- ETA
- workeri
- coada
- erori
- `slowExifCandidates`
- `slowExifDetections`

## 16. Backup metadate

Ruta obisnuita:

```text
%LOCALAPPDATA%\PhotoOrganizer\_Backup_Metadate\<RunId>\
```

In versiuni anterioare sau in alte limbi acelasi rol poate aparea ca `_MetadataBackup` sau `_CopiaSeguridadMetadatos`. Motorul le trateaza ca aliasuri ale aceluiasi rol intern.

Sunt copii complete ale fisierelor inainte de o reparare EXIF reala. Nu sunt miniaturi si nu sunt doar metadate.

Se creeaza cand exista:

- modificari reale
- reparare EXIF
- ExifTool disponibil
- reparare EXIF necesara
- incredere suficienta

Nu se creeaza la `DryRun` sau `Test scan`.

Daca rularea se termina corect, motorul incearca sa curete backup-ul acelei rulari. Daca exista anulare, oprire sau verificare nesigura, il poate pastra.

Nu migra acest folder intre calculatoare. Starea portabila este in `%APPDATA%\PhotoOrganizer\`.

## 17. Migrarea aplicatiei pe alt calculator

PhotoOrganizer este gandit ca aplicatie portabila sub forma de folder.

Metoda recomandata este `Setari` -> `Mod expert` -> `Migreaza UDMRS pe alt PC`.

Asistentul genereaza o carpeta in:

```text
%USERPROFILE%\Downloads\UDMRS-MigrationPackages\
```

In acea carpeta creeaza:

- un ZIP al instalarii partajate, creat din radacina reala UDMRS
- un ZIP al starii utilizatorului curent, creat din `%APPDATA%\PhotoOrganizer`
- `MigrationGuide.txt`

ZIP-ul de instalare include `App`, `Docs`, `Tools`, `Branding`, `Templates`, `Releases`, `Config`, `README.md` si lansatoarele existente. ZIP-ul de utilizator include `ProcessedFiles.json`, `Config`, `IndexBackups`, setarile dashboard-ului si fisiere JSON de stare necesare. Nu include `Logs`, `Runtime`, `*.progress.json`, cozi, PID/status temporare sau backup-uri EXIF locale.

La final, dashboard-ul deschide automat folderul unde au fost generate pachetele.

Pentru a muta aplicatia:

1. Copiaza folderul complet `<CarpetaUDMRS>` pe noul calculator.
2. Deschide `Start-PhotoOrganizer.cmd` din folderul copiat.
3. Nu reutiliza shortcut-uri `.lnk` vechi. Windows pastreaza rute absolute in shortcut si poate porni o copie veche.
4. Daca vrei shortcut pe desktop, creeaza-l din copia actuala.

Pentru a pastra progresul, hash-urile si starea incrementala, copiaza si:

```text
%APPDATA%\PhotoOrganizer\
```

in aceeasi locatie a profilului Windows de pe noul calculator.

Acest folder contine de obicei:

- `ProcessedFiles.json`
- `settings.json`
- alte fisiere JSON interne de stare

Logurile, rapoartele HTML si progress traiesc in:

```text
%APPDATA%\PhotoOrganizer\Logs\
```

Acest folder trebuie sa poata fi curatat cand nu exista rulari active. Infrastructura runtime pentru Mod expert traieste separat in:

```text
%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole\
```

Nu este log sau raport; contine runner, coada si stare temporara controlata de dashboard. Backup-urile temporare de metadate traiesc in `%LOCALAPPDATA%\PhotoOrganizer\` si nu trebuie migrate decat daca investighezi o rulare recenta.

Lansatorul `.cmd` foloseste `%~dp0` pentru a porni din propriul folder. Scripturile folosesc `$PSScriptRoot` pentru resurse, limbi, manuale si ExifTool in instalatia curenta. Logurile, runtime-ul tehnic, setarile si excluderile utilizatorului traiesc in `%APPDATA%\PhotoOrganizer`. De aceea aplicatia nu trebuie sa depinda de rute private fixe.

Daca biblioteca este in OneDrive sau alt furnizor cloud:

- asteapta finalizarea sincronizarii inainte de procese mari
- marcheaza biblioteca disponibila offline daca vei rula Organize, Normalize sau Repair masiv
- UDMRS Dashboard poate convietui cu continut sincronizat, dar proceseaza doar fisiere verificabile local
- placeholder-ele cloud-only sunt omise si raportate, nu sunt descarcate automat

Dupa migrare, restaurare sau mutari manuale de foldere, ruleaza:

```text
Sincronizeaza indexul / repara modificari manuale
```

Aceasta revalideaza `ProcessedFiles.json`, actualizeaza rute dupa hash cand este posibil si evita duplicate false cauzate de rute vechi.

Regula importanta:

```text
O singura operatie activa a motorului la un moment dat.
```

Nu rula doua dashboard-uri, doua console tehnice sau doua comenzi PowerShell pe aceeasi biblioteca in acelasi timp. Dashboard-ul blocheaza actiunile normale cand `Mod expert` ruleaza o actiune avansata, dar este bine sa eviti si lansarile manuale paralele.

## 18. Ce sa nu faci in timpul unei rulari

Cat timp PhotoOrganizer ruleaza, nu muta si nu sterge:

- folderul sursa
- folderul destinatie
- folderul `Logs`
- folderul de backup metadate
- fisierele procesate

Foloseste `Anulare` din dashboard daca trebuie sa opresti rularea.

## 19. Comenzi oficiale rapide

Referinta completa cu toate comenzile este in `Docs\CommandReference.html`.

### Test scan

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" -TestScan -Language ro
```

### Organize DryRun

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" -Language ro
```

### Organize Apply + RepairExif

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" -Apply -RepairExif -Language ro
```

### NormalizeExistingFolders DryRun cu curatare sigura

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" -NormalizeExistingFolders -KeepEmptyFolders:$false -Language ro
```

### ReconcileProcessedDatabase DryRun

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" -ReconcileProcessedDatabase -Language ro
```

### DedupeCleanup DryRun

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Poze_Organizate" -DedupeCleanup -Language ro
```

### Cautare probleme in loguri

```powershell
Select-String -Path "%APPDATA%\PhotoOrganizer\Logs\*.log" -Pattern "Error","Slow EXIF candidate","DateInfo resolved","Missing/Stale","JSON conflict"
```

Comenzile Apply complete sunt in `Docs\CommandReference.html`. Verifica intotdeauna DryRun si rapoartele HTML inainte de modificari reale.








