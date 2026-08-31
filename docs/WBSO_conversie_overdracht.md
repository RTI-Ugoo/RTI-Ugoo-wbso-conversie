# WBSO-conversie automatisering — Overdrachtsdocument

Dit document bevat alle context, architectuur, werking en openstaande punten van het
WBSO-conversieproject, zodat een andere Claude (of persoon) het project kan overnemen
en voortzetten. Lees dit volledig door voordat je wijzigingen voorstelt.

---

## 1. Wat het project doet

**Doel:** het handmatig overtypen van WBSO-projectteksten uit Word naar het RVO-portaal
automatiseren. Consultants bij Ugoo (Nederlands subsidieadviesbureau, ~25-30 consultants,
500-800 aanvragen/jaar) zetten een Word-aanvraag in een SharePoint-map; een geautomatiseerde
flow zet die per project om naar een ingevuld RVO-projectformulier (PDF), en zet de gevulde
PDF's terug in dezelfde map. De consultant controleert daarna en importeert in het RVO-portaal.
Er blijft dus een mens-in-de-lus.

**Belangrijk principe:** het script neemt de Word-tekst **1-op-1** over (deterministisch,
geen AI-herschrijving). Het herkent de structuur (welke tekst is een knelpunt, oplossing, etc.)
en zet die in het juiste PDF-veld. Het herschrijft nooit inhoud.

---

## 2. Architectuur

Microsoft-stack, volledig binnen de bestaande Ugoo-omgeving:

```
Consultant zet .docx in SharePoint-map
        │
        ▼
Power Automate flow (trigger: bestand aangemaakt)
        │  stuurt .docx als base64 naar de Function
        ▼
Azure Function (Python 3.11, "wbso_convert")
        │  parse .docx → vul RVO-PDF per project → geef gevulde PDF's terug
        ▼
Power Automate schrijft gevulde PDF's terug in SharePoint
        │
        ▼
Teams-notificatie naar de consultant met fouten/aandachtspunten
```

**Kosten:** ~€200-275/jaar totaal. Azure Functions valt binnen de gratis marge (verbruiksplan,
volume is een fractie van 1M executions/mnd). Grootste post is één Power Automate Premium-licentie
voor het serviceaccount (~€170-215/jaar). Schaalt NIET met aantal consultants.

**Schaalbaarheid (30 consultants):** ruim voldoende. Azure Functions schaalt parallel automatisch.
Aandachtspunten: cold start (eerste aanroep na inactiviteit is traag, tot ~1 min) en
trigger-polling (op piek kan verwerking 1-2 min duren). Beide zijn ongemak, geen blokkade;
er gaat niets verloren.

---

## 3. Azure Function — deployment

**Function App URL (HTTP-trigger):**
```
https://wbso-conversie-portal-rvo-g7hjamfvc7g8aeet.westeurope-01.azurewebsites.net/api/wbso_convert?code=<FUNCTION_KEY>
```
(West Europe i.v.m. EU-compliance. De function key staat in de Power Automate HTTP-actie.)

**Projectstructuur (deze bestanden horen bij elkaar):**
```
azure_function/
├── host.json
├── requirements.txt          (azure-functions, python-docx==1.1.2, pypdf==4.3.1, cryptography>=41.0.0)
├── .funcignore
└── wbso_convert/
    ├── __init__.py           HTTP-entrypoint: leest request, kiest formulier, roept converteer() aan
    ├── function.json         httpTrigger, authLevel function, POST
    ├── wbso_core.py          ALLE parsing- en vul-logica (~1000 regels)
    ├── leeg_formulier.pdf         RVO-formulier PROGRAMMATUUR (v1.8, 62 AcroForm-velden)
    └── leeg_formulier_fysiek.pdf  RVO-formulier FYSIEK PRODUCT (v1.81)
```

**Deploy-workflow (VS Code):**
1. Zorg dat alle bestanden in de projectmap de nieuwste versie zijn (vooral `wbso_core.py`
   en `__init__.py` MOETEN synchroon zijn — een mismatch geeft de fout
   "converteer() takes from 2 to 3 positional arguments but 4 were given").
2. **Deploy ALTIJD de hele projectmap**, niet losse bestanden (voorkomt versie-mismatch en
   ontbrekende PDF's). Rechtermuisklik op de projectmap → "Deploy to Function App" → bevestig
   overschrijven.
3. Na deploy soms een **Restart** van de Function App in de Azure-portal nodig (caching).

**Setup-hordes die al opgelost zijn (voor de historie):**
- Resourceproviders Microsoft.Web/Storage/Insights moesten op subscription-niveau
  geregistreerd worden door de IT-partner (buiten Contributor-rechten).
- Python 3.11 vereist (Azure Functions ondersteunt 3.14 niet). Aparte 3.11-installatie nodig.
- `cryptography>=41.0.0` in requirements.txt loste een "AES algorithm"-fout op bij pdf-schrijven.

---

## 4. Power Automate flow

**Naam:** "WBSO conversie". Alle stappen in volgorde:

1. **Trigger:** "When a file is created (properties only)", Site `/regelingen`,
   Library "Gedeelde documenten".
   Trigger conditions:
   - `@endsWith(...['body/{FilenameWithExtension}'], '.docx')`
   - `@not(contains(...['body/{Path}'], 'Archief'))`
   - `@contains(...['body/{Path}'], 'Conversiemap')` (hoofdletter C)

2. **Get file content:** File Identifier via expressie `triggerOutputs()?['body/{Identifier}']`
   (NIET "ID" — geeft File-not-found).

3. **Initialize variable** `Samenvatting` (String, leeg) — voor de Teams-notificatie.

4. **Initialize variable** `FormulierType` (String), Value via expressie:
   ```
   if(contains(triggerOutputs()?['body/{Path}'], 'Productontwikkeling'), 'fysiek', 'programmatuur')
   ```
   (MOET vóór de HTTP-actie staan, op hoofdniveau — niet in een lus/conditie.)

5. **HTTP:** POST naar function-URL, header Content-Type=application/json, body:
   ```json
   {
     "bestandsnaam": "@{triggerOutputs()?['body/{FilenameWithExtension}']}",
     "docx_base64": "@{body('Get_file_content')?['$content']}",
     "formuliertype": "@{variables('FormulierType')}"
   }
   ```

6. **Parse JSON:** schema (let op: `fouten` en `aandachtspunten` als aparte arrays):
   ```json
   {
     "type": "object",
     "properties": {
       "aantal_projecten": { "type": "integer" },
       "projecten": {
         "type": "array",
         "items": {
           "type": "object",
           "properties": {
             "projectnummer": { "type": "string" },
             "projecttitel": { "type": "string" },
             "bestandsnaam": { "type": "string" },
             "pdf_base64": { "type": "string" },
             "fouten": { "type": "array", "items": { "type": "string" } },
             "aandachtspunten": { "type": "array", "items": { "type": "string" } }
           }
         }
       }
     }
   }
   ```

7. **Apply to each** over `projecten`:
   - **Create file:** Site `/regelingen`, GEEN Library-veld; Folder Path via Expression
     `concat('/', triggerOutputs()?['body/{Path}'])`, File Name = Body bestandsnaam,
     File Content = `base64ToBinary(items('Apply_to_each')?['pdf_base64'])`.
   - **Condition:** `add(length(items('Apply_to_each')?['fouten']), length(items('Apply_to_each')?['aandachtspunten']))`
     is greater than `0`.
     (Als de expressie klaagt: gebruik zonder `@{ }`, of de Or-variant met twee length-checks.)
     - **True → Append to string variable** `Samenvatting`, Value (HTML-opmaak, want Teams
       rendert platte-tekst-witregels onbetrouwbaar):
       ```
       <br><br><b>Project @{items('Apply_to_each')?['projectnummer']}</b>@{if(empty(items('Apply_to_each')?['fouten']), '', concat('<br>❌ Fouten: ', join(items('Apply_to_each')?['fouten'], ' | ')))}@{if(empty(items('Apply_to_each')?['aandachtspunten']), '', concat('<br><br>⚠️ Aandachtspunten: ', join(items('Apply_to_each')?['aandachtspunten'], ' | ')))}
       ```
     - **False →** leeg.

8. **Post message in a chat or channel** (Microsoft Teams), NA de Apply to each:
   - Post as: Flow bot; Post in: Chat with Flow bot;
   - Recipient: "Gemaakt door e-mail" (Created By Email) uit de trigger.
   - Message (HTML):
     ```
     <b>Je WBSO-aanvraag is verwerkt</b> ✅<br>Document: @{triggerOutputs()?['body/{FilenameWithExtension}']}<br>Aantal projecten: @{body('Parse_JSON')?['aantal_projecten']}<br>@{if(empty(trim(variables('Samenvatting'))), 'Alle projecten zijn zonder aandachtspunten verwerkt.', variables('Samenvatting'))}<br><br>Controleer de formulieren voordat je ze in het RVO-portaal importeert of pas het Word-document aan en plaats een nieuwe kopie in de Conversiemap.
     ```

**Belangrijke Teams-les:** platte tekst met `\n` / `decodeUriComponent('%0A')` / non-breaking
spaces geeft onbetrouwbare witregels in Teams (vooral de eerste regel wordt niet vet). HTML-opmaak
(`<b>`, `<br>`, `<br><br>`) werkt wél betrouwbaar. Klein risico: `<`, `>`, `&` in titels/teksten
kunnen de HTML verstoren (zeldzaam).

---

## 5. Mapstructuur in SharePoint

```
regelingen (site) / Gedeelde documenten / WBSO / 2. Aanvragen - Handleidingen templates en bijlagen / Conversiemap /
├── ABA/                      ← .docx hier = PROGRAMMATUUR
│   └── Productontwikkeling/  ← .docx hier = FYSIEK PRODUCT
├── BBE/
│   └── Productontwikkeling/
├── RTI/
│   └── Productontwikkeling/
└── ... (submap per consultant, code als naam)
```

Een document direct in de consultant-map → programmatuur (standaard).
Een document in de submap "Productontwikkeling" → fysiek product.
De flow leidt dit af uit het pad (zie `FormulierType`-variabele). **Programmatuur is altijd de
default**; alleen bij "Productontwikkeling" in het pad schakelt hij naar fysiek. Zo kan een
programmatuur-aanvraag nooit per ongeluk als fysiek behandeld worden.

**Let op — toegankelijkheid van de map voor consultants:** de "Synchroniseren"-knop is in deze
tenant uitgeschakeld. Alternatieven (nog niet definitief opgelost): "Snelkoppeling naar OneDrive
toevoegen", IT-partner de sync laten aanzetten, of een mail-ingang als alternatieve trigger.

---

## 6. Hoe het script werkt: mapping van Word naar RVO-velden

### Twee leesstructuren (door elkaar ondersteund)
- **Kopjes-structuur:** TK, TO, TN staan onder eigen kopjes (standaardsjablonen).
- **Inline-structuur:** TK/TO/TR staan als markers ("TK:", "TO:") in lopende tekst onder één
  techniek-kop.

### Kopblok (bovenaan elk project)
| Word-label (NL / EN) | Gaat naar |
|---|---|
| `Project <code>: <titel>` (één regel) OF `Project number:`/`Projectnummer:` + `Project title:`/`Projecttitel:` (gesplitst) | Projectnummer + Projecttitel |
| `Startdatum:` / `Start project:` / `Start date:` | DatumStart |
| `Aantal ontwikkeluren:` / `WBSO-uren:` / `S&O-uren:` / `Number of R&D hours:` | Uren |
| `Kosten/uitgaven:` / `Costs/expenses:` | forfaitair vs werkelijk |
| `Statutaire bedrijfsnaam:` / `Statutory company name:` | referentie (niet op formulier) |

### Inhoudelijke velden — PROGRAMMATUUR-formulier
| RVO-veldnaam | Betekenis | Herkend via kop (o.a.) | Inline-marker |
|---|---|---|---|
| WbsoProjectomschrijving | Projectomschrijving | projectomschrijving, project description | — |
| WbsoToelichtingWijzigingPlanning | Update project | statusupdate, projectupdate, voortgang | — |
| WbsoBeschrijvingTechnischProbleem | 1. Knelpunten | technische knelpunten (tk), technical bottlenecks | TK, TK1, TB, TB1, knelpunt, bottleneck |
| WbsoGekozenOplossingsrichting | 2. Oplossingen | technische oplossingsrichtingen (to), oplossingen, technical solutions, solutions | TO, TO1, TS, TS1, technische oplossingsrichting, technical solution |
| WbsoBestaandeMethodenTechnieken | 3. Talen/tools | programmeertalen, programming languages, languages/development environment | — |
| WbsoZelfOntwikkelenMethoden | 4. Nieuwheid | technische nieuwheid, technical novelty | TR, TR1, TN, TN1, technisch risico, technical risk |
| Ontwikkeling1-10 / DatumGereed1-10 | Fasering (tabel) | fasering, projectfasering, planning werkzaamheden, project planning | — |

Verborgen/vaste velden programmatuur: WbsoProjecttype='B', WbsoProjectzwaartepunt='PRG'.

### Inhoudelijke velden — FYSIEK PRODUCT-formulier (afwijkend!)
| RVO-veldnaam | Betekenis |
|---|---|
| WbsoTWOVraag1 | 1. Technische knelpunten |
| WbsoTWOVraag2 | 2. Technische oplossingsrichtingen |
| WbsoTWOVraag3 | 3. Technische nieuwheid |
| (geen talen/tools-veld) | vervalt bij fysiek |

WbsoProjectzwaartepunt fysiek: 'PDT' (Technisch nieuw product, default) of 'PPS' (Technisch nieuw
productieproces). Script zet standaard PDT + aandachtspunt om te checken.
Het optionele programmatuur-blok binnen het fysiek-formulier wordt BEWUST GENEGEERD (komt in ~2%
van de aanvragen voor; consultant vult dat handmatig bij).

De parsing/validatie draait intern op de programmatuur-veldnamen (getest); bij fysiek worden de
eindvelden vlak voor teruggave omgezet naar WbsoTWOVraag1/2/3 en vervalt het talen/tools-veld.

### Tekenlimieten
1500 tekens: omschrijving, update, TK, TO, tools, nieuwheid. 200: titel. 25: projectnummer.

---

## 7. Belangrijke gedragsregels en features (allemaal geïmplementeerd + getest)

- **Component-nummering:** expliciete nummers (TK1, TB3, cijfer-kop "3.") worden gerespecteerd
  INCLUSIEF gaten (bv. 1,3,4,6 als TK2/TK5 afgerond zijn). Ongenummerde vette component-namen
  (bv. "Asynchrone endpoints") worden automatisch doorgenummerd (auto_comp_teller, reset per project).
- **TO/TR-labeling in output:** oplossingen krijgen `TO<nr>:` en risico's `TR<nr>:` waarbij `<nr>`
  het nummer van het bijbehorende knelpunt is (manier A). Label alleen op de eerste alinea; vervolg-
  alinea's zonder herhaald label. TR/TN gaan naar het nieuwheid-veld (vraag 4 / TWOVraag3).
- **Combi-koppen:** een kop die ≥2 techniek-onderdelen noemt ("Technische knelpunten (TK),
  oplossingen (TO) en risico's (TR)", NL+EN) → inline-structuur. Eén onderdeel ("Oplossingen",
  "Solutions") → losse kop naar eigen veld. Onderscheid = aantal onderdelen in de titel.
- **Ingebedde markers:** " TO:" / " TR:" middenin een doorlopende alinea (spatie ervoor, dubbele
  punt erachter) worden herkend en splitsen de tekst. Ook binnen een TO-alinea wordt een ingebedde
  TR: afgesplitst.
- **Modus-tracking binnen component:** het script onthoudt of het in knelpunt/oplossing/risico zit;
  vervolgzinnen zonder marker gaan naar de HUIDIGE modus (niet standaard naar TK). Reset naar tk bij
  nieuw component. (Dit loste op dat TO-vervolgtekst in TK belandde.)
- **Label-stripping:** inline-marker-labels ("Technisch knelpunt:", "TO:", "Component:", etc.) worden
  uit de tekst gestript; de inhoud blijft.
- **Content controls (w:sdt):** tekst in Word-content-controls (ontstaat bij Google Docs-bewerking)
  wordt uitgelezen via `paragraaf_tekst()` (python-docx' .text slaat sdt over).
- **Faseringstabel:** laatste kolom = datum (ondersteunt 2- én 3-koloms tabellen). Datums worden NIET
  gevalideerd/gesignaleerd (te veel ruis).
- **TK-overloop >1500 (alleen programmatuur):** overschot naar vraag 3 met marker
  `[Vervolg TK's, zie vraag 3]` + prefix `Vervolg TK's: ` + witregel vóór talen/tools. Bij fysiek
  geen overloop (geen talen-veld) → tekenlimiet-fout.
- **Uren-extractie (`extract_uren`):** pakt het eerste zinvolle getal, ook met tekst eromheen
  ("3.000 hours" → 3000). Jaartallen 2020-2099 genegeerd tenzij enige getal. Meerdere verschillende
  getallen → eerste getal + AANDACHTSPUNT. Complex/placeholder → leeg → "Geen uren gevonden".
- **Niet-vette projectkop:** wordt herkend als de code kort/code-achtig is (voorkomt dat gewone
  zinnen matchen).
- **Tekennormalisatie voor PDF:** en-dash/em-dash/krul-apostrofs/non-breaking space → veilige
  equivalenten (é/ë/ï behouden). Projectcode met en-dash ("TPCZ – 01c") genormaliseerd in bestandsnaam
  naar "TPCZ-01c", maar exact overgenomen in het formulierveld.
- **Legenda/skip:** definitieregels ("TK = Technisch knelpunt"), "(max. N tekens)"-titelregels en
  combi-uitleg worden overgeslagen.
- **Engelse documenten** volledig ondersteund (labels, kopjes, markers).

### Signalering: FOUTEN (❌) vs AANDACHTSPUNTEN (⚠️)
- **Fouten** (moet opgelost vóór indienen): leeg verplicht veld (omschrijving/TK/tools/nieuwheid),
  geen uren, geen/ongeldige startdatum, tekenlimiet overschreden, projecttitel ontbreekt.
  TO leeg = fout (met hint "mogelijk oplossingstekst in vraag 1" als er oplossing-taal in TK staat).
- **Aandachtspunten** (even checken): TK doorgeschoven naar vraag 3; oplossing-taal in vraag 1 terwijl
  vraag 2 gevuld; meerdere urengetallen; kosten/uitgaven handmatig; zwaartepunt-check bij fysiek.
- Per project max. één TO-gerelateerde melding (geen dubbelingen).

---

## 8. Bewust NIET gebouwd / afgewezen

- **Kosten/uitgaven-velden** niet geautomatiseerd (klein %, te divers).
- **Opmerkingen** niet naar formulier gemapt.
- **Update-vs-nieuw-project logica** (technische velden niet aanraken bij updates): afgewezen. Reden:
  RVO-import is alles-of-niets, tekst wordt toch 1-op-1 overgenomen (= identiek aan goedgekeurde versie),
  vorige-pdf hergebruiken te veel handwerk. Enige risico is andere indeling bij eerste conversie van een
  bestaand project → eenmalige menselijke check.
- **Optioneel programmatuur-blok in fysiek-formulier** genegeerd (~2%, handmatig).
- **Consultants zelf code laten aanpassen/deployen:** afgeraden (kwaliteitsrisico bij officiële
  aanvragen). Wel besproken als veilige tussenweg: consultants stellen verbeteringen voor → Claude
  beoordeelt/weerlegt → legt voor aan beheerder → beheerder verwerkt met regressietest. En: veilige
  delen (herkende kopjes) configureerbaar maken in een aparte lijst.

---

## 9. Testdocumenten (structuurvarianten die werken)

Het script is getest tegen ~13 documentvarianten. Elk dekt een specifiek geval af:
- Kopjes-structuur (standaardsjablonen), NL en EN.
- Inline-structuur met markers, genummerd en ongenummerd.
- Gesplitste projectkop (Project number / Project title).
- Content controls (Google Docs-bewerkt).
- Combi-koppen (2 en 3 onderdelen).
- 3-koloms faseringstabel + cijfer-component-koppen.
- Volledig Engels document, niet-vette projectkop.
- Complex document met componenten + inline-labels in twee schrijfwijzen + vervolgzinnen +
  ingebedde markers (de zwaarste stresstest).
- Fysiek-product-aanvraag (doorlopende TK/TO/TR in één alinea).

**Regressie-discipline:** bij ELKE wijziging aan `wbso_core.py` moet je alle testdocumenten opnieuw
doorrekenen en vergelijken met de vorige uitkomst. Een kleine wijziging aan de kop-herkenning of
inline-logica breekt gemakkelijk een ander geval. Dit is door de hele bouw het vaste ritueel geweest.

---

## 10. Openstaande punten / mogelijke vervolgstappen

- **Schrijfwijzer voor consultants** (NOG SAMEN TE STELLEN): lichte afspraken zodat de conversie
  betrouwbaarder wordt — markers aan het begin van een regel, één schrijfwijze voor TK/TO/TR,
  componentnamen zonder dubbele punt, faseringsdatums dd-mm-jjjj, projectregel-formaat, kopjes als
  Heading. Dit vermindert randgevallen structureel.
- **Portaalvalidatie:** één keer een gevuld formulier in het RVO-portaal importeren om te bevestigen
  dat tekentelling/regelovergangen exact kloppen.
- **SharePoint-toegang consultants** (sync-knop uit): alternatief kiezen (OneDrive-snelkoppeling,
  IT sync aanzetten, of mail-ingang).
- **Beheerapplicatie** (verkennend besproken, niets gebouwd): Power Apps-pagina voor overzicht +
  feedback ophalen. Feedback → Claude beoordeelt verbetervoorstellen → beheerder verwerkt. Deployen
  bij beheerder houden.
- **Jaarlijkse check** op nieuwe RVO-formulierversies (nu programmatuur v1.8, fysiek v1.81) →
  leeg_formulier(_fysiek).pdf vervangen.
- **Brede testronde** met echte aanvragen na de gestapelde verbeteringen.

---

## 11. Werkinstructies voor de overnemende Claude

- **Taal:** de gebruiker (Rutger, WBSO-adviseur, Amsterdam) werkt in het Nederlands. Antwoord in het
  Nederlands.
- **Toon:** hij is technisch onderlegd maar geen software-engineer; leg keuzes en afwegingen uit,
  wees eerlijk over risico's en betrouwbaarheidsgrenzen, en gok niet op documentstructuur — vraag om
  een (geanonimiseerd) voorbeelddocument als iets niet werkt.
- **Werkwijze bij scriptwijzigingen:** (1) bekijk het echte document eerst, (2) pas gericht aan,
  (3) draai de volledige regressietest tegen alle testdocumenten, (4) lever de hele zip + los
  `wbso_core.py` op, (5) herinner aan "deploy de hele map, niet losse bestanden".
- **Kernbestand:** `wbso_convert/wbso_core.py` bevat vrijwel alle logica. `__init__.py` is dun
  (request → converteer → response). Beide moeten synchroon gedeployed worden.
- **Hoofdzorg van de gebruiker:** het bestaande programmatuur-pad mag NOOIT slechter worden door een
  nieuwe feature. Nieuwe dingen komen er als aparte, opt-in route bij; programmatuur blijft de default.

---

*Dit document beschrijft de staat van het project aan het eind van de bouwsessie. De bijbehorende
code staat in `wbso_azure_function.zip` (volledig project) en `wbso_core.py` (losse laatste versie).*
