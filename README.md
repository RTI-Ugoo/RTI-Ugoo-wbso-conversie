# WBSO-conversie — Azure Function

Zet WBSO-projectteksten uit een Word-document (.docx) automatisch om naar ingevulde
RVO-projectformulieren (PDF). Onderdeel van een Power Automate-flow: consultants zetten
een Word-aanvraag in een SharePoint-map, deze Function vult per project het juiste
RVO-formulier, Power Automate zet de PDF's terug en stuurt een Teams-notificatie.

Volledige achtergrond, architectuur, veldmapping en werkinstructies staan in
[`docs/WBSO_conversie_overdracht.md`](docs/WBSO_conversie_overdracht.md) — lees dat
document eerst voordat je wijzigingen aanbrengt.

## Structuur

```
.
├── host.json
├── requirements.txt
├── .funcignore
└── wbso_convert/
    ├── __init__.py           HTTP-entrypoint
    ├── function.json         httpTrigger, authLevel function, POST
    ├── wbso_core.py          alle parsing- en vullogica
    ├── leeg_formulier.pdf         RVO-formulier programmatuur
    └── leeg_formulier_fysiek.pdf  RVO-formulier fysiek product
```

## Deployen

Deployen gaat via GitHub Actions (`.github/workflows/deploy.yml`): elke push naar
`main` deployt automatisch naar de Azure Function App. Handmatig deployen vanuit VS
Code is niet meer nodig.

Eenmalige setup: voeg het Azure publish-profile van de Function App toe als GitHub
secret `AZURE_FUNCTIONAPP_PUBLISH_PROFILE` (Settings → Secrets and variables →
Actions). Het publish-profile haal je op in de Azure Portal: Function App →
Overview → "Get publish profile" (download als XML, plak de volledige inhoud als
secret-waarde).

## Regressie-discipline

Bij elke wijziging aan `wbso_core.py`: alle testdocumenten opnieuw doorrekenen en
vergelijken met de vorige uitkomst voordat je pusht naar `main` (zie het
overdrachtsdocument, hoofdstuk 9-11).
