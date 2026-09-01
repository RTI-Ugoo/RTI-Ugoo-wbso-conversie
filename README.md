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

**Status: geparkeerd — voorlopig handmatig deployen vanuit VS Code, zoals
voorheen** (rechtermuisklik op de projectmap → "Deploy to Function App").
De automatische GitHub Actions-deploy is voorbereid maar staat uit tot de
eenmalige Entra ID-setup hieronder met de IT-partner is afgerond (daarvoor is
een tenant-niveau app registration nodig, waar de consultant zelf geen
toegang toe heeft). Zodra dat rond is: zet in
`.github/workflows/deploy.yml` de trigger terug naar `push: branches: [main]`
en elke push naar `main` deployt weer automatisch.

De Function App draait op het **Flex Consumption**-plan. Dat plan ondersteunt geen
publish-profile/Kudu-deploy meer; de workflow logt daarom in via een Azure AD
service principal met OIDC (`Azure/login`), zonder wachtwoord in GitHub.

Eenmalige setup (Azure Portal, Entra ID):
1. **App registration** aanmaken (Entra ID → App registrations → New registration).
   Noteer de **Application (client) ID** en **Directory (tenant) ID**.
2. Onder die app registration: **Certificates & secrets → Federated credentials →
   Add credential**, scenario "GitHub Actions deploying Azure resources",
   Organization `RTI-Ugoo`, Repository `RTI-Ugoo-wbso-conversie`, Entity type
   "Branch", Branch name `main`.
3. Rol toekennen: Function App → **Access control (IAM) → Add role assignment** →
   rol "Website Contributor" → toewijzen aan de zojuist aangemaakte app
   registration (service principal).
4. **Subscription ID** noteren (Subscriptions-pagina in Azure Portal).
5. Drie GitHub secrets toevoegen (Settings → Secrets and variables → Actions):
   `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, `AZURE_SUBSCRIPTION_ID`.

`AZURE_FUNCTIONAPP_PUBLISH_PROFILE` is niet meer nodig en mag verwijderd worden.

## Regressie-discipline

Bij elke wijziging aan `wbso_core.py`: alle testdocumenten opnieuw doorrekenen en
vergelijken met de vorige uitkomst voordat je pusht naar `main` (zie het
overdrachtsdocument, hoofdstuk 9-11).
