

import re
import unicodedata
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject


def paragraaf_tekst(par) -> str:
    """Lees de volledige tekst van een paragraaf, INCLUSIEF tekst die in
    content controls (w:sdt) zit. python-docx' .text slaat sdt-inhoud over;
    dit gebeurt vaak bij documenten die via Google Docs zijn bewerkt.
    Leest alle w:t-elementen in documentvolgorde."""
    delen = []
    for node in par._p.iter():
        if node.tag == qn('w:t'):
            delen.append(node.text or "")
        elif node.tag == qn('w:tab'):
            delen.append("\t")
        elif node.tag in (qn('w:br'), qn('w:cr')):
            delen.append("\n")
    return "".join(delen)

# ---------------------------------------------------------------- limieten
LIMIETEN = {
    "Projectnummer": 25,
    "Projectnaam": 200,
    "WbsoProjectomschrijving": 1500,
    "WbsoToelichtingWijzigingPlanning": 1500,
    "WbsoBeschrijvingTechnischProbleem": 1500,
    "WbsoGekozenOplossingsrichting": 1500,
    "WbsoBestaandeMethodenTechnieken": 1500,
    "WbsoZelfOntwikkelenMethoden": 1500,
    "WbsoKostenOmschrijving": 500,
    "WbsoUitgavenOmschrijving": 500,
}

VELD_LABELS = {
    "Projectnummer": "Projectnummer",
    "Projectnaam": "Projecttitel",
    "WbsoProjectomschrijving": "Projectomschrijving",
    "WbsoToelichtingWijzigingPlanning": "Update project",
    "WbsoBeschrijvingTechnischProbleem": "1. Technische knelpunten",
    "WbsoGekozenOplossingsrichting": "2. Technische oplossingsrichtingen",
    "WbsoBestaandeMethodenTechnieken": "3. Programmeertalen en tools",
    "WbsoZelfOntwikkelenMethoden": "4. Technische nieuwheid",
    "WbsoSOUrenBegroot": "Uren",
    "WbsoKostenOmschrijving": "Kosten",
    "WbsoUitgavenOmschrijving": "Uitgaven",
}


def schoon(tekst: str) -> str:
    """Normaliseer whitespace, behoud regeleinden."""
    tekst = unicodedata.normalize("NFC", tekst)
    tekst = tekst.replace("\t", " ")
    tekst = re.sub(r"[ ]{2,}", " ", tekst)
    return tekst.strip()


def body_elementen(doc):
    """Paragrafen en tabellen in documentvolgorde."""
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def is_bold(par) -> bool:
    runs = [r for r in par.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs if r.text.strip())


# ---------------------------------------------------------------- parsing
# Projectregel: "Project <code>: <titel>". De code mag spaties/streepjes bevatten
# (bv. "XX - 13", "XX-13", "XX 13"); die worden daarna genormaliseerd.
RE_PROJECT = re.compile(r"^Project\s+(.+?)\s*:\s*(.+)$", re.IGNORECASE)
# Gesplitste projectkop: "Project number:" / "Projectnummer:" (code) en
# "Project title:" / "Projecttitel:" (titel) op aparte regels.
RE_PROJECT_NUMMER = re.compile(r"^Project\s*(?:number|nummer)\s*:\s*(.+)$", re.IGNORECASE)
RE_PROJECT_TITEL = re.compile(r"^Project\s*(?:title|titel)\s*:\s*(.+)$", re.IGNORECASE)


def normaliseer_projectcode(code: str) -> str:
    """Maak projectcodes consistent: 'XX - 13', 'XX – 13' en 'XX  13' worden 'XX-13'.
    Diverse streepjesvarianten (en-dash, em-dash) worden een gewoon koppelteken;
    spaties rond het streepje verdwijnen; losse spaties worden een streepje."""
    code = code.strip()
    # alle streepjesvarianten -> gewoon koppelteken
    code = re.sub(r"[\u2010\u2011\u2012\u2013\u2014\u2015]", "-", code)
    # spaties rondom een streepje weghalen
    code = re.sub(r"\s*-\s*", "-", code)
    # resterende spatie(s) tussen deel en nummer -> streepje
    code = re.sub(r"\s+", "-", code)
    # dubbele streepjes samenvouwen
    code = re.sub(r"-{2,}", "-", code)
    return code
RE_HEADERVELD = re.compile(r"^(Statutaire bedrijfsnaam|Statutory company name|Periode|Period|Aantal ontwikkeluren|WBSO-uren|WBSO uren|Aantal S&O-uren|S&O-uren|Number of R&D hours|Number of hours|Kosten/uitgaven|Costs/expenses|Startdatum|Start date|Start project|Startdatum project)\s*:\s*(.*)$", re.IGNORECASE)
RE_INTERNE_NOOT = re.compile(r"^<<.*>>$")

# Tussenkopjes/labels die geen inhoudelijke tekst zijn en niet in een veld horen.
_NEGEER_REGELS = re.compile(
    r"^\s*(verdeling activiteiten per entiteit)\s*$",
    re.IGNORECASE,
)
RE_DATUM = re.compile(r"(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})")

KOP_MAP = [
    # statusupdate / voortgang
    ("statusupdate", "update"),
    ("status update", "update"),
    ("status project", "update"),
    ("statusproject", "update"),
    ("projectupdate", "update"),
    ("updateproject", "update"),
    ("update project", "update"),
    ("voortgang", "update"),
    ("stand van zaken", "update"),
    # opmerkingen (max 500) -> apart veld
    ("opmerkingen", "opmerkingen"),
    ("general remark", "opmerkingen"),
    ("remarks", "opmerkingen"),
    # fasering
    ("fasering", "fasering"),
    ("projectfasering", "fasering"),
    ("planning werkzaamheden", "fasering"),
    ("project planning", "fasering"),
    # projectomschrijving
    ("projectomschrijving", "omschrijving"),
    ("project description", "omschrijving"),
    ("omschrijving van het project", "omschrijving"),
    ("algemene omschrijving", "omschrijving"),
    # --- LOSSE SUBKOPJES (standaardsjabloon-structuur) ---
    # Deze staan VOOR de algemene "techniek"-kop zodat ze eerder matchen.
    ("technische knelpunten programmatuur", "tk_sectie"),
    ("omschrijving technische knelpunten", "tk_sectie"),
    ("technische knelpunten (tk)", "tk_sectie"),
    ("technische knelpunten", "tk_sectie"),
    ("technical bottlenecks and technical solutions", "techniek"),  # combi-kop -> inline
    ("technische knelpunten en -oplossingsrichtingen", "techniek"),  # combi TK/TO -> inline
    ("technische knelpunten en oplossingsrichtingen", "techniek"),
    ("technische knelpunten en technische oplossingsrichtingen", "techniek"),
    ("knelpunten en -oplossingsrichtingen", "techniek"),
    ("technical bottlenecks in", "tk_sectie"),
    ("technical bottlenecks", "tk_sectie"),
    ("technical bottleneck", "tk_sectie"),
    ("technische oplossingsrichtingen programmatuur", "to_sectie"),
    ("omschrijving technische oplossingsrichtingen", "to_sectie"),
    ("technische oplossingsrichtingen (to)", "to_sectie"),
    ("technische oplossingsrichtingen", "to_sectie"),
    ("technische oplossingsrichting", "to_sectie"),
    ("technische oplossingen", "to_sectie"),
    ("oplossingsrichtingen", "to_sectie"),
    ("oplossingen", "to_sectie"),
    ("intended technical solutions", "to_sectie"),
    ("technical solutions", "to_sectie"),
    ("technical solution", "to_sectie"),
    ("solutions", "to_sectie"),
    # losse programmeertalen-kop (VOOR nieuwheid, want bevat ook 'technische')
    ("programmeertalen", "tools_sectie"),
    ("programmeertaal", "tools_sectie"),
    ("ontwikkelomgevingen en tools", "tools_sectie"),
    ("tools en ontwikkelomgevingen", "tools_sectie"),
    ("programming languages", "tools_sectie"),
    ("languages, development environment", "tools_sectie"),
    ("languages, development", "tools_sectie"),
    # technische nieuwheid (TN) -> nieuwheid-sectie
    ("omschrijving technische nieuwheid", "nieuwheid_sectie"),
    ("technische nieuwheid programmatuur", "nieuwheid_sectie"),
    ("technische nieuwheid", "nieuwheid_sectie"),
    ("level of technical novelty", "nieuwheid_sectie"),
    ("technical novelty", "nieuwheid_sectie"),
    # technische beschrijving / probleemstelling (inline TK/TO/TR-blok) -> NA de subkopjes
    ("technische probleemstelling", "techniek"),
    ("technische beschrijving", "techniek"),
    ("technical description", "techniek"),
    ("knelpunten en oplossing", "techniek"),
    ("technische risico", "trtn"),
    ("afkortingen", "skip"),
]


def _normaliseer_kop(tekst: str) -> str:
    """Verwijder spaties, koppeltekens en underscores en maak lowercase,
    zodat 'Status update', 'Status-update' en 'Statusupdate' gelijk matchen."""
    return re.sub(r"[\s\-_]+", "", tekst.lower())


def _telt_techniek_onderdelen(tekst: str) -> int:
    """Tel hoeveel verschillende techniek-onderdelen (knelpunt/oplossing/risico)
    in een kop genoemd worden. Gebruikt voor het onderscheid combi-kop (>=2,
    inline) versus losse kop (1, eigen veld)."""
    laag = tekst.lower()
    onderdelen = 0
    # knelpunt
    if re.search(r"knelpunt|bottleneck|\(tk'?s?\)|\(tb'?s?\)", laag):
        onderdelen += 1
    # oplossing
    if re.search(r"oplossing|solution|\(to'?s?\)|\(ts'?s?\)", laag):
        onderdelen += 1
    # risico
    if re.search(r"risico|risk|\(tr'?s?\)|\(tn'?s?\)", laag):
        onderdelen += 1
    return onderdelen


def kop_type(tekst: str):
    """Bepaal het sectietype van een kop. Kiest de LANGSTE matchende sleutel,
    zodat 'technische nieuwheid' niet per ongeluk als 'techniek' wordt gezien.

    Een regel die een inline-marker is ('Technische oplossingsrichting: <zin>')
    wordt NIET als kop gezien, zodat die als inhoud in de techniek-sectie
    verwerkt wordt in plaats van als sectiekop."""
    laag_los = tekst.lower()
    laag_dicht = _normaliseer_kop(tekst)

    # Combi-kop: noemt de titel 2+ techniek-onderdelen (knelpunt + oplossing,
    # evt. + risico), dan is het een gecombineerde kop -> inline-structuur.
    # Bv. "Technische knelpunten (TK), oplossingen (TO) en risico's (TR)".
    # Alleen als het echt een korte kop is (geen lange zin met een dubbele punt).
    if _telt_techniek_onderdelen(tekst) >= 2 and len(tekst) < 90:
        # niet als het een inline-marker-regel is ("TK: <lange zin>")
        m = re.match(r"^[^:]{0,60}:\s*(.+)$", tekst)
        if not (m and len(m.group(1).split()) > 6):
            return "techniek"

    # Kop die begint met "Status" -> update-sectie, ongeacht wat erachter staat
    # (bv. "Status en voortgang februari t/m december 2025", "Status S&O-project
    # en voortgang, april - juni 2026"). De periode/datums erachter verschillen
    # per aanvraag en per consultant, dus die worden hier bewust niet vastgelegd.
    if re.match(r"^status\b", tekst, re.IGNORECASE):
        return "update"

    beste = None
    beste_len = 0
    beste_sleutel = ""
    for sleutel, naam in KOP_MAP:
        sleutel_dicht = _normaliseer_kop(sleutel)
        if laag_dicht.startswith(sleutel_dicht) or laag_los.startswith(sleutel):
            if len(sleutel_dicht) > beste_len:
                beste = naam
                beste_len = len(sleutel_dicht)
                beste_sleutel = sleutel

    # Inline-marker-bescherming: als de match een TK/TO/TR-sectie is en er staat
    # een dubbele punt met een substantiele zin erachter (meer dan ~6 woorden),
    # dan is het geen sectiekop maar inline tekst -> laat de techniek-classificatie
    # het afhandelen.
    if beste in ("tk_sectie", "to_sectie", "trtn"):
        m = re.match(r"^[^:]{0,60}:\s*(.+)$", tekst)
        if m and len(m.group(1).split()) > 6:
            return None

    return beste


def _strip_marker_label(tekst: str) -> str:
    """Verwijder een inline-marker-label aan het begin van de tekst, zoals
    'TK:', 'TO:', 'TR:', 'Technisch knelpunt:', 'Technische oplossingsrichting:',
    'Technisch risico:'. De inhoud zelf blijft staan; alleen het label verdwijnt."""
    patronen = [
        r"^(TK|TB|TO|TS|TR|TN)\s*\d*\s*:\s*",
        r"^Technische?\s+knelpunt(en)?\s*:\s*",
        r"^Technische?\s+oplossingsricht(ing|ingen)?\s*:\s*",
        r"^Technische?\s+risico'?s?\s*:\s*",
        r"^Technical\s+bottleneck(s)?\s*:\s*",
        r"^Technical\s+solution(s)?\s*:\s*",
        r"^Technical\s+risk(s)?\s*:\s*",
        r"^Component\s*:\s*",
    ]
    for pat in patronen:
        nieuw = re.sub(pat, "", tekst, flags=re.IGNORECASE)
        if nieuw != tekst:
            return nieuw.strip()
    return tekst


def _splits_ingebedde_markers(tekst: str):
    """Splits een TK-tekstblok op ingebedde ' TO:' en ' TR:' markers die
    middenin dezelfde alinea staan (niet aan het begin van een regel).
    Geeft (tk_deel, to_deel, tr_deel) terug; to/tr zijn leeg als niet gevonden.

    Alleen markers met een spatie ervoor en dubbele punt erachter worden
    herkend, zodat toevallige lettercombinaties in lopende tekst niet splitsen.
    De TO/TR aan het begin van de tekst worden hier niet behandeld (die vangt
    classificeer_techniek al af)."""
    # zoek de eerste ingebedde TO: of TR: (spatie ervoor, ergens na positie 0)
    m_to = re.search(r"\s(TO|TS)\s*:\s*", tekst)
    m_tr = re.search(r"\s(TR|TN)\s*:\s*", tekst)
    pos_to = m_to.start() if m_to else None
    pos_tr = m_tr.start() if m_tr else None

    if pos_to is None and pos_tr is None:
        return tekst, "", ""

    to_deel = ""
    tr_deel = ""

    if pos_to is not None and (pos_tr is None or pos_to < pos_tr):
        # TO komt eerst: TK loopt tot TO, TO loopt tot TR (indien aanwezig)
        tk_deel = tekst[:pos_to].strip()
        na_to = tekst[m_to.end():]
        m_tr2 = re.search(r"\s(TR|TN)\s*:\s*", na_to)
        if m_tr2:
            to_deel = na_to[:m_tr2.start()].strip()
            tr_deel = na_to[m_tr2.end():].strip()
        else:
            to_deel = na_to.strip()
    else:
        # TR komt eerst (geen TO ervoor): TK loopt tot TR, rest is TR
        tk_deel = tekst[:pos_tr].strip()
        tr_deel = tekst[m_tr.end():].strip()

    return tk_deel, to_deel, tr_deel


def classificeer_techniek(par) -> str:
    """Deel paragrafen binnen de techniek-sectie in bij TK, TO of TR.

    Herkent genummerde en ongenummerde markers, NL en EN:
      knelpunt  : TK / TK1 / TB / TB1 / technisch knelpunt / technical bottleneck
      oplossing : TO / TO1 / TS / TS1 / technische oplossingsrichting / technical solution
      risico    : TR / TR1 / TN / TN1 / technisch risico / technical risk / technical novelty
    """
    t = paragraaf_tekst(par).strip()
    laag = t.lower()

    # Componentkop: begint met TK/TB<nummer> (evt. "TK2 en TO2:") en is vet.
    if re.match(r"^(tk|tb)\s*\d*(\s+en\s+(to|ts)\s*\d*)?\s*[:.]", laag) and is_bold(par):
        return "component"
    # Componentkop met alleen een cijfer, bv "3. (Integrations platform)" of
    # "5. AI development platform", vet, binnen de techniek-sectie. Het cijfer
    # dient als componentnummer voor de bijbehorende TB/TS/TR eronder.
    if re.match(r"^\d+[.\)]\s+\S", t) and is_bold(par):
        return "component"
    if laag.startswith(("onderstaand staat", "technische knelpunten en technische oplossingsricht",
                         "technical bottlenecks")):
        return "skip"
    # Legenda-/definitieregels bovenaan, bv "TK = Technisch knelpunt" of
    # "PP = Propositie Patient    TK = Technisch knelpunt". Bevatten '=' als
    # definitie en zijn geen echte inhoud.
    if re.match(r"^[A-Z]{2,4}\s*=", t) or "\t=" in t or " = technisch" in laag or " = technische" in laag:
        return "skip"
    # Kop met "(max. N tekens)"-suffix is een titelregel, geen inhoud. Bv.
    # "Technische Risico's en Technische Nieuwheid (TR/TN)  (max. 1500 tekens)".
    if re.search(r"\(max\.?\s*[\d.]+\s*tekens\)", laag):
        return "skip"

    # Korte vette subkop die het techniek-type benoemt met een licht afwijkende
    # zinsvolgorde (bv. "Beoogde technische oplossingsrichtingen" i.p.v.
    # "Technische oplossingsrichtingen", "Programmeertechnische probleemstellingen"
    # i.p.v. "Technische knelpunten"). De bestaande tk/to/tr-detectie hieronder
    # verwacht dat de kop MET het kernwoord begint; deze varianten hebben er een
    # bijvoeglijk woord voor staan, waardoor ze anders als los, ongenummerd
    # component gezien worden (en de koptekst zelf als inhoud in TK belandt).
    # Hier wordt alleen de modus gezet; de koptekst zelf is geen inhoud.
    if is_bold(par) and len(t) < 90 and ":" not in t:
        if re.search(r"^(\S+\s+){0,2}oplossingsricht", laag) or re.search(r"^(\S+\s+){0,2}technical\s+solutions?\b", laag):
            return "to_kop"
        if re.search(r"^(\S+\s+){0,2}probleemstelling", laag) or re.search(r"^(\S+\s+){0,2}technisch(e)?\s+knelpunt", laag):
            return "tk_kop"
        if (re.search(r"^(\S+\s+){0,2}technisch(e)?\s+risico", laag)
                or re.search(r"^(\S+\s+){0,2}technische?\s+nieuwheid", laag)
                or re.search(r"^(\S+\s+){0,2}technical\s+(risk|novelty)", laag)):
            return "tr_kop"

    # Ongenummerde component-kop: een korte VETTE naam die zelf geen TK/TO/TR-marker
    # is (bv. "Asynchrone endpoints", "Conditionele logging & tracing"). Dient als
    # component die automatisch doorgenummerd wordt. Alleen als kort en zonder
    # dubbele punt gevolgd door een zin (dat zou inline-tekst zijn).
    if (is_bold(par) and len(t) < 70
            and not re.match(r"^(tk|tb|to|ts|tr|tn)\s*\d*\s*[:.]", laag)
            and not laag.startswith(("technisch", "technical", "programmeer", "programming",
                                     "tools", "knelpunt", "bottleneck", "oplossing", "solution",
                                     "risico", "risk", "nieuwheid", "novelty"))):
        m_zin = re.match(r"^[^:]{0,60}:\s*(.+)$", t)
        if not (m_zin and len(m_zin.group(1).split()) > 4):
            return "component"

    # Technisch risico / novelty / technical risk (TR1, TN1, enz.)
    if (re.match(r"^(het\s+|als\s+)?technische?\s+risico", laag)
            or re.match(r"^technical\s+risk", laag)
            or re.match(r"^technical\s+novelty", laag)
            or re.match(r"^(tr|tn)\s*\d*\s*[:.]", laag)
            or re.match(r"^(tr|tn)\s*\d*\b", laag)):
        return "tr"

    # Technische oplossingsrichting / technical solution (TO1, TS1, enz.)
    if (re.match(r"^(als\s+)?(de\s+)?technische?\s+oplossingsricht", laag)
            or re.match(r"^(als\s+)?technical\s+solution", laag)
            or re.match(r"^(als\s+)?(to|ts)\s*\d*\s*[:.]", laag)
            or re.match(r"^(als\s+)?(to|ts)\s*\d*\b", laag)):
        return "to"

    # Technisch knelpunt / technical bottleneck (TK1, TB1, enz.)
    if (re.match(r"^(het\s+)?technisch(e)?\s+knelpunt", laag)
            or re.match(r"^(het\s+)?technical\s+bottleneck", laag)
            or re.match(r"^(als\s+)?(tk|tb)\s*\d*\s*[:.]", laag)
            or re.match(r"^(als\s+)?(tk|tb)\s*\d*\b", laag)
            or laag.startswith(("knelpunt", "bottleneck"))):
        return "tk"
    if "knelpunt" in laag[:60] or "bottleneck" in laag[:60]:
        return "tk"
    return "context"  # inleidende componenttekst -> hoort bij TK


def parse_docx(pad: str):
    doc = Document(pad)
    projecten = []
    project = None
    sectie = None
    huidige_component = None
    auto_comp_teller = 0
    techniek_modus = "tk"  # waar we zijn binnen een component: tk / to / tr

    def nieuw_project(nummer, titel):
        return {
            "nummer": schoon(nummer),
            "titel": schoon(titel),
            "bedrijf": "",
            "periode": "",
            "uren": "",
            "kosten_uitgaven": "",
            "startdatum": "",
            "update": [],
            "opmerkingen": [],
            "omschrijving": [],
            "fasering": [],
            "tk": [],
            "to": [],
            "tr": [],
            "tools": [],
            "trtn_extra": [],
            "waarschuwingen": [],
            "fouten": [],
            "aandachtspunten": [],
        }

    verwacht_tools = False

    for el in body_elementen(doc):
        if isinstance(el, Table):
            if project is not None and sectie == "fasering":
                for rij in el.rows:
                    cellen = [schoon(c.text) for c in rij.cells]
                    if len(cellen) < 2:
                        continue
                    # Activiteit = eerste cel. Datum = LAATSTE cel.
                    # Sommige tabellen hebben 3 kolommen (activiteit dubbel + datum);
                    # door de laatste cel als datum te nemen werkt zowel 2 als 3 koloms.
                    activiteit = cellen[0]
                    datum = cellen[-1]
                    if activiteit and not activiteit.lower().startswith(("ontwikkeling", "development activity")):
                        project["fasering"].append((activiteit, datum))
            continue

        tekst = schoon(paragraaf_tekst(el))
        if not tekst:
            continue

        eerste_run_bold = bool(el.runs) and bool(el.runs[0].bold)

        # Gesplitste projectkop eerst: "Project number:" start een nieuw project.
        m_nummer = RE_PROJECT_NUMMER.match(tekst)
        if m_nummer and eerste_run_bold:
            if project:
                projecten.append(project)
            project = nieuw_project(m_nummer.group(1).strip(), "")
            sectie = None
            huidige_component = None
            auto_comp_teller = 0
            techniek_modus = "tk"
            continue
        # "Project title:" vult de titel van het lopende project aan.
        m_titel = RE_PROJECT_TITEL.match(tekst)
        if m_titel and project is not None and not project["titel"]:
            project["titel"] = schoon(m_titel.group(1))
            continue

        m = RE_PROJECT.match(tekst)
        if m:
            code_kandidaat = m.group(1).strip()
            # Een echte projectcode is kort en code-achtig (letters/cijfers/streepjes/
            # spaties, geen volzin). Accepteer de kop als hij vet is, OF als de code
            # dat patroon volgt (zodat niet-vette projectkoppen ook herkend worden,
            # zonder dat gewone zinnen die met "Project" beginnen matchen).
            code_achtig = bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 \-\u2010-\u2015.]{0,24}", code_kandidaat))
            if eerste_run_bold or code_achtig:
                if project:
                    projecten.append(project)
                project = nieuw_project(code_kandidaat, m.group(2))
                sectie = None
                huidige_component = None
                auto_comp_teller = 0
                techniek_modus = "tk"
                continue

        if project is None:
            continue  # voorblad

        mh = RE_HEADERVELD.match(tekst)
        if mh and sectie is None:
            sleutel, waarde = mh.group(1).lower(), schoon(mh.group(2))
            if "bedrijfsnaam" in sleutel or "company name" in sleutel:
                project["bedrijf"] = waarde
            elif "periode" in sleutel or "period" in sleutel:
                project["periode"] = waarde
            elif "uren" in sleutel or "hours" in sleutel:
                project["uren"] = waarde
            elif "kosten" in sleutel or "costs" in sleutel:
                project["kosten_uitgaven"] = waarde
            elif "start" in sleutel:
                datum_match = RE_DATUM.search(waarde)
                project["startdatum"] = datum_match.group(1) if datum_match else waarde
            continue

        # Kop-detectie: een echte Heading-stijl, OF een korte vetgedrukte
        # alinea die begint met een bekende koptekst (voor documenten die
        # geen Heading-stijlen gebruiken).
        is_heading_stijl = el.style.name.startswith("Heading")
        # Titelregel met tekenlimiet-suffix "(max. X tekens)" is altijd een kop,
        # ook zonder Heading-stijl. Bevat vaak "TK/TO" e.d.; nooit als inhoud tonen.
        is_maxlimiet_kop = bool(re.search(r"\(\s*max\.?\s*[\d.\s]+\s*tekens\s*\)", tekst, re.IGNORECASE))
        is_vette_kop = (
            eerste_run_bold
            and is_bold(el)
            and len(tekst) < 90
            and not re.match(r"^(TK|TO|TR|TB|TS)\d*\s*[:.]", tekst, re.IGNORECASE)
        )
        # Statuskop zonder opmaak: sommige sjablonen typen "Status en voortgang
        # ..." als gewone, niet-vette alinea (stijl "No Spacing" i.p.v. Heading
        # of bold). Zonder deze uitzondering wordt zo'n regel helemaal niet als
        # kop gezien en blijft de update-tekst ongedetecteerd. Kort en zonder
        # zininterne punt, om te voorkomen dat gewone lopende tekst die met
        # "Status" begint hier per ongeluk in trapt.
        is_statuskop = bool(
            re.match(r"^status\b", tekst, re.IGNORECASE)
            and len(tekst) < 100
            and not re.search(r"[.!?]\s+\S", tekst)
        )
        if is_heading_stijl or is_vette_kop or is_maxlimiet_kop or is_statuskop:
            kt = kop_type(tekst)
            if kt:
                sectie = kt
                verwacht_tools = (kt == "tools_sectie")
                continue
            if is_maxlimiet_kop:
                # Titelregel die we niet aan een sectie konden koppelen:
                # niet als inhoud opnemen, wel doorgaan met de huidige sectie.
                continue

        if RE_INTERNE_NOOT.match(tekst):
            continue

        if _NEGEER_REGELS.match(tekst):
            continue

        if el.style.name in ("List Paragraph",):
            tekst = "- " + tekst

        if sectie == "update":
            project["update"].append(tekst)
        elif sectie == "opmerkingen":
            project["opmerkingen"].append(tekst)
        elif sectie == "omschrijving":
            project["omschrijving"].append(tekst)
        elif sectie == "tk_sectie":
            # Kopjes-structuur: inhoud onder "Technische knelpunten (TK)" -> direct TK
            project["tk"].append(tekst)
        elif sectie == "to_sectie":
            # Kopjes-structuur: inhoud onder "Technische oplossingsrichtingen (TO)" -> direct TO
            project["to"].append(tekst)
        elif sectie == "techniek":
            soort = classificeer_techniek(el)
            if soort == "skip":
                continue
            if soort in ("to_kop", "tk_kop", "tr_kop"):
                # Subkop die alleen het techniek-type benoemt (geen marker-
                # inhoud erachter): alleen de modus zetten, koptekst zelf niet
                # als inhoud opslaan.
                techniek_modus = {"to_kop": "to", "tk_kop": "tk", "tr_kop": "tr"}[soort]
                continue
            if soort == "component":
                # Expliciet nummer (TK3, TB3, of cijfer-kop "3.") -> gebruik dat
                # en zet de auto-teller erop. Geen nummer -> tel automatisch door.
                m_comp = re.match(r"^(?:TK|TB)\s*(\d+)", tekst, re.IGNORECASE)
                m_num = re.match(r"^(\d+)[.\)]", tekst)
                if m_comp:
                    auto_comp_teller = int(m_comp.group(1))
                elif m_num:
                    auto_comp_teller = int(m_num.group(1))
                else:
                    auto_comp_teller += 1
                huidige_component = f"TK{auto_comp_teller}"
                techniek_modus = "tk"  # nieuw component -> terug naar knelpunt
                kop = tekst
                project["tk"].append(kop)
                continue
            # Componentnummer afleiden (bv. 'TK2' -> '2') voor TO2:/TR2:-labels.
            comp_nr = ""
            if huidige_component:
                mnr = re.search(r"(\d+)", huidige_component)
                comp_nr = mnr.group(1) if mnr else ""
            to_prefix = f"TO{comp_nr}: " if comp_nr else ""
            tr_prefix = f"TR{comp_nr}: " if comp_nr else ""

            if soort == "to":
                # Nieuwe oplossing begint: label alleen op de eerste alinea.
                eerste = techniek_modus != "to"
                techniek_modus = "to"
                schoon_to = _strip_marker_label(tekst)
                m_tr_in = re.search(r"\s(TR|TN)\s*:\s*", schoon_to)
                if m_tr_in:
                    project["to"].append((to_prefix if eerste else "") + schoon_to[:m_tr_in.start()].strip())
                    techniek_modus = "tr"
                    project["tr"].append(tr_prefix + schoon_to[m_tr_in.end():].strip())
                else:
                    project["to"].append((to_prefix if eerste else "") + schoon_to)
            elif soort == "tr":
                eerste = techniek_modus != "tr"
                techniek_modus = "tr"
                project["tr"].append((tr_prefix if eerste else "") + _strip_marker_label(tekst))
            elif soort == "tk":
                techniek_modus = "tk"
                # Splits op ingebedde markers binnen dezelfde alinea.
                tk_deel, to_deel, tr_deel = _splits_ingebedde_markers(tekst)
                project["tk"].append(_strip_marker_label(tk_deel))
                if to_deel:
                    techniek_modus = "to"
                    project["to"].append(to_prefix + _strip_marker_label(to_deel))
                if tr_deel:
                    techniek_modus = "tr"
                    project["tr"].append(tr_prefix + _strip_marker_label(tr_deel))
            else:  # context: vervolgtekst -> volgt de huidige modus, zonder herhaald label
                if techniek_modus == "to":
                    project["to"].append(tekst)
                elif techniek_modus == "tr":
                    project["tr"].append(tekst)
                else:
                    # nog in knelpunt-modus; splits alsnog op ingebedde markers
                    tk_deel, to_deel, tr_deel = _splits_ingebedde_markers(tekst)
                    project["tk"].append(_strip_marker_label(tk_deel))
                    if to_deel:
                        techniek_modus = "to"
                        project["to"].append(to_prefix + _strip_marker_label(to_deel))
                    if tr_deel:
                        techniek_modus = "tr"
                        project["tr"].append(tr_prefix + _strip_marker_label(tr_deel))
        elif sectie == "tools_sectie":
            project["tools"].append(tekst)
        elif sectie == "nieuwheid_sectie":
            if tekst.lower().startswith("programmeertalen"):
                verwacht_tools = True
                continue
            if verwacht_tools:
                project["tools"].append(tekst)
            else:
                project["trtn_extra"].append(tekst)
        elif sectie == "trtn":
            project["trtn_extra"].append(tekst)

    if project:
        projecten.append(project)
    return projecten


# ---------------------------------------------------------------- validatie
STANDAARDTEKST_TRTN = re.compile(r"technische risico'?s zijn toegelicht", re.IGNORECASE)

# Zinsdelen die op een oplossingsrichting wijzen, gebruikt voor zachte signalering
# als ze in het knelpunt-veld opduiken zonder aparte TO-marker.
_RUIKT_NAAR_OPLOSSING = re.compile(
    r"(als\s+oplossingsricht|"
    r"om\s+dit\s+op\s+te\s+lossen|"
    r"ter\s+oplossing|"
    r"de\s+oplossing(srichting)?\s+(is|bestaat|onderzoekt|betreft)|"
    r"as\s+(a\s+)?technical\s+solution|"
    r"to\s+solve\s+this|"
    r"oplossingsrichting\s+onderzoekt\s+men)",
    re.IGNORECASE,
)


def valideer_datum(d: str) -> bool:
    try:
        datetime.strptime(d, "%d-%m-%Y")
        return True
    except ValueError:
        return False



def samenvoegen(alineas):
    """Voeg alinea's samen met een enkele regelovergang, zodat de telling
    overeenkomt met wat er in het RVO-veld staat (en het portaal telt)."""
    return "\n".join(alineas)

def extract_uren(tekst: str):
    """Haal het urengetal uit de uren-tekst. Geeft (uren, meerdere) terug.
    - Pakt het eerste zinvolle getal, ook met tekst eromheen ('3.000 hours',
      '3000 S&O-uren', '3000 R&D hours').
    - 'meerdere' is True als er meerdere verschillende getallen staan (bv. uren
      per meerdere B.V.'s), zodat er een aandachtspunt gegenereerd kan worden.
    - Een jaartal (2000-2099) telt alleen mee als er geen ander getal is."""
    if not tekst:
        return "", False
    # alle getallen in volgorde (met punt/spatie als duizendscheiding)
    ruwe = re.findall(r"\d[\d.\s]*\d|\d", tekst)
    getallen = []
    for g in ruwe:
        cijfers = re.sub(r"[^\d]", "", g)
        if cijfers:
            getallen.append(cijfers)
    if not getallen:
        return "", False
    # jaartallen apart houden; alleen gebruiken als er geen andere getallen zijn.
    # Range 2020-2099: dat zijn de periode-jaren in huidige aanvragen. 2000 uur
    # is aannemelijker als urenaantal dan als jaartal, dus die telt als getal.
    niet_jaar = [g for g in getallen if not re.fullmatch(r"20[2-9]\d", g)]
    kandidaten = niet_jaar if niet_jaar else getallen
    uren = kandidaten[0]
    meerdere = len(set(kandidaten)) > 1
    return uren, meerdere


def bouw_veldwaarden(p: dict, formuliertype: str = "programmatuur"):
    """Map geparste projectdata naar PDF-veldnamen + verzamel fouten en aandachtspunten.
    formuliertype: 'programmatuur' (standaard) of 'fysiek' (fysiek product/productieproces)."""
    fouten = p["fouten"]
    aandacht = p["aandachtspunten"]

    nieuwheid_delen = []
    if p["tr"]:
        nieuwheid_delen.append("\n".join(p["tr"]))
    for extra in p["trtn_extra"]:
        if not STANDAARDTEKST_TRTN.search(extra):
            nieuwheid_delen.append(extra)

    uren_waarde, uren_meerdere = extract_uren(p["uren"])

    velden = {
        "Bedrijfsnaam": p["bedrijf"],
        "Projectnummer": p["nummer"],
        "Projectnaam": p["titel"],
        "DatumStart": p["startdatum"],
        "WbsoProjectomschrijving": samenvoegen(p["omschrijving"]),
        "WbsoToelichtingWijzigingPlanning": "\n".join(p["update"]),
        "WbsoBeschrijvingTechnischProbleem": "\n".join(p["tk"]),
        "WbsoGekozenOplossingsrichting": "\n".join(p["to"]),
        "WbsoBestaandeMethodenTechnieken": "\n".join(p["tools"]),
        "WbsoZelfOntwikkelenMethoden": "\n".join(nieuwheid_delen),
        "WbsoSOUrenBegroot": uren_waarde,
    }

    # Fasering (max 10 regels). Faseringsdatums worden NIET gesignaleerd
    # (te veel ruis); de consultant controleert de fasering sowieso.
    for i, (activiteit, datum) in enumerate(p["fasering"][:10], start=1):
        velden[f"Ontwikkeling{i}"] = activiteit
        velden[f"DatumGereed{i}"] = datum

    # Kosten/uitgaven
    if "forfait" in p["kosten_uitgaven"].lower():
        pass  # velden bewust leeg bij forfaitaire keuze
    elif p["kosten_uitgaven"]:
        aandacht.append(f"Kosten/uitgaven staat op '{p['kosten_uitgaven']}'; kostenvelden zijn NIET automatisch gevuld. Handmatig aanvullen.")

    # Overloop TK -> vraag 3 (talen/tools). Alleen bij programmatuur, want het
    # fysiek-product-formulier heeft geen apart talen/tools-veld. Bij fysiek
    # leidt een te lange TK gewoon tot een tekenlimiet-fout hieronder.
    if formuliertype != "fysiek":
        verwerk_tk_overloop(velden, fouten, aandacht)

    # Tekenlimieten -> FOUT (portaal weigert/kapt af)
    for veld, limiet in LIMIETEN.items():
        waarde = velden.get(veld, "")
        if len(waarde) > limiet:
            fouten.append(f"{VELD_LABELS.get(veld, veld)}: {len(waarde)} tekens, limiet is {limiet} ({len(waarde) - limiet} te veel). Inkorten vereist.")

    # Verplichte inhoud leeg -> FOUT (behalve TO, die heeft eigen logica hieronder).
    # Het talen/tools-veld is alleen verplicht bij programmatuur.
    verplicht = ["WbsoProjectomschrijving", "WbsoBeschrijvingTechnischProbleem",
                 "WbsoZelfOntwikkelenMethoden"]
    if formuliertype != "fysiek":
        verplicht.append("WbsoBestaandeMethodenTechnieken")
    for veld in verplicht:
        if not velden.get(veld, "").strip():
            fouten.append(f"{VELD_LABELS[veld]}: leeg. Controleer of deze sectie in het Word-document staat.")

    # Update project leeg -> AANDACHTSPUNT (geen fout: bij een nieuw project
    # hoeft dit veld niet gevuld te zijn, maar bij een lopend project vaak wel).
    if not velden.get("WbsoToelichtingWijzigingPlanning", "").strip():
        aandacht.append("Update project is leeg. Bij een nieuw project kan dat kloppen; "
                        "bij een lopend project controleren of de statusupdate in het "
                        "Word-document is meegenomen.")

    if not velden["WbsoSOUrenBegroot"]:
        fouten.append("Geen uren gevonden.")
    elif uren_meerdere:
        aandacht.append(f"Meerdere urenaantallen gevonden; het eerste ({uren_waarde}) is ingevuld. "
                        "Controleer of dit klopt (bv. bij uren per meerdere B.V.'s).")

    # --- TO-logica (vraag 2), opgeschoond zodat er nooit dubbele meldingen zijn ---
    tk_tekst = velden.get("WbsoBeschrijvingTechnischProbleem", "")
    to_tekst = velden.get("WbsoGekozenOplossingsrichting", "")
    oplossing_taal_in_tk = bool(tk_tekst and _RUIKT_NAAR_OPLOSSING.search(tk_tekst))
    if not to_tekst.strip():
        # Vraag 2 leeg = FOUT. Met hint als er oplossing-taal in vraag 1 staat.
        if oplossing_taal_in_tk:
            fouten.append("Oplossingsrichtingen (vraag 2) is leeg. Mogelijk staat de "
                          "oplossingstekst nog in vraag 1; controleer de scheiding.")
        else:
            fouten.append("Oplossingsrichtingen (vraag 2) is leeg. Controleer of deze "
                          "sectie in het Word-document staat.")
    elif oplossing_taal_in_tk:
        # Vraag 2 gevuld, maar oplossing-taal ook in vraag 1 = AANDACHTSPUNT.
        aandacht.append("Oplossing-taal in vraag 1; controleer of knelpunten en "
                        "oplossingen goed gescheiden zijn.")

    # Startdatum -> FOUT als ontbreekt of ongeldig formaat
    if not p["startdatum"]:
        fouten.append("Startdatum niet gevonden; handmatig invullen.")
    elif not valideer_datum(p["startdatum"]):
        fouten.append(f"Startdatum '{p['startdatum']}' heeft niet het formaat dd-mm-jjjj.")

    # Projecttitel ontbreekt (bv. gesplitste kop zonder title-regel) -> FOUT
    if not p["titel"].strip():
        fouten.append("Projecttitel niet gevonden; handmatig invullen.")

    # Fysiek-product-formulier gebruikt andere veldnamen voor de inhoudelijke
    # vragen en heeft geen apart talen/tools-veld. De parsing/validatie hierboven
    # draait op de programmatuur-namen (getest); hier zetten we alleen de
    # eindvelden om naar de fysiek-namen.
    if formuliertype == "fysiek":
        velden["WbsoTWOVraag1"] = velden.pop("WbsoBeschrijvingTechnischProbleem", "")
        velden["WbsoTWOVraag2"] = velden.pop("WbsoGekozenOplossingsrichting", "")
        velden["WbsoTWOVraag3"] = velden.pop("WbsoZelfOntwikkelenMethoden", "")
        velden.pop("WbsoBestaandeMethodenTechnieken", None)  # geen talen/tools-veld

    return velden



# ---------------------------------------------------------------- overloop
TK_MARKER = "[Vervolg TK's, zie vraag 3]"
TK_VERVOLG_PREFIX = "Vervolg TK's: "


def splits_op_zin(tekst: str, budget: int):
    """Splits tekst zo dicht mogelijk bij budget, bij voorkeur op een
    alineagrens, anders op een zinseinde, anders op een spatie."""
    if len(tekst) <= budget:
        return tekst, ""
    kop = tekst[:budget]
    # 1) alineagrens
    idx = kop.rfind("\n\n")
    if idx > budget * 0.4:
        return tekst[:idx].rstrip(), tekst[idx:].lstrip()
    # 2) zinseinde (punt/vraagteken/uitroepteken gevolgd door spatie of regeleinde)
    kandidaten = [m.end() for m in re.finditer(r"[.!?](?=\s)", kop)]
    if kandidaten and kandidaten[-1] > budget * 0.4:
        idx = kandidaten[-1]
        return tekst[:idx].rstrip(), tekst[idx:].lstrip()
    # 3) laatste spatie
    idx = kop.rfind(" ")
    return tekst[:idx].rstrip(), tekst[idx:].lstrip()


def verwerk_tk_overloop(velden: dict, fouten: list, aandacht: list):
    """Als het TK-veld boven de limiet zit: kap af op een logisch punt en
    plaats het vervolg in vraag 3 (Programmeertalen en tools), voor de talen/tools."""
    limiet = LIMIETEN["WbsoBeschrijvingTechnischProbleem"]
    tk = velden.get("WbsoBeschrijvingTechnischProbleem", "")
    if len(tk) <= limiet:
        return
    budget = limiet - len(TK_MARKER) - 1  # ruimte voor regelovergang + marker
    kop, rest = splits_op_zin(tk, budget)
    velden["WbsoBeschrijvingTechnischProbleem"] = kop + "\n" + TK_MARKER

    tools = velden.get("WbsoBestaandeMethodenTechnieken", "")
    vervolg = TK_VERVOLG_PREFIX + rest
    # Witregel tussen het doorgeschoven TK-vervolg en de talen/tools (leesbaarheid).
    velden["WbsoBestaandeMethodenTechnieken"] = vervolg + ("\n\n" + tools if tools else "")

    aandacht.append(f"TK-veld was {len(tk)} tekens (limiet {limiet}). Laatste {len(rest)} tekens "
                    f"doorgeschoven naar vraag 3. Controleer of de afkapping logisch is.")
    if len(velden["WbsoBestaandeMethodenTechnieken"]) > LIMIETEN["WbsoBestaandeMethodenTechnieken"]:
        fouten.append("Vraag 3 zit na het TK-vervolg zelf boven de 1500 tekens. Handmatig inkorten vereist.")


# ---------------------------------------------------------------- pdf vullen
def vul_pdf(bron: str, velden: dict, doel: Path):
    writer = PdfWriter(clone_from=bron)
    try:
        writer.set_need_appearances_writer(True)
    except Exception:
        acro = writer._root_object["/AcroForm"]
        acro[NameObject("/NeedAppearances")] = BooleanObject(True)

    # Radiobutton samenwerking -> Nee
    velden = dict(velden)
    velden["Samenwerking"] = "/SamenwerkingNee"

    for pagina in writer.pages:
        writer.update_page_form_field_values(pagina, velden, auto_regenerate=False)

    with open(doel, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------- rapport
def maak_rapport(projecten_uit: list, doel: Path):
    regels = ["# Controlerapport WBSO-formulieren", ""]
    regels.append(f"Gegenereerd: {datetime.now().strftime('%d-%m-%Y %H:%M')}. "
                  "Controleer elk formulier volledig voordat het wordt geüpload in het RVO-portaal.")
    regels.append("")
    for p, velden, pdfnaam in projecten_uit:
        regels.append(f"## Project {p['nummer']}: {p['titel']}")
        regels.append(f"Bestand: `{pdfnaam}` | Periode: {p['periode']} | Uren: {velden.get('WbsoSOUrenBegroot','')} | Kosten/uitgaven: {p['kosten_uitgaven']}")
        regels.append("")
        regels.append("| Veld | Tekens | Limiet | Status |")
        regels.append("|---|---|---|---|")
        for veld, limiet in LIMIETEN.items():
            if veld in ("Projectnummer", "Projectnaam"):
                continue
            lengte = len(velden.get(veld, ""))
            status = "OK" if lengte <= limiet else f"**{lengte - limiet} TE VEEL**"
            if lengte == 0:
                status = "leeg"
            regels.append(f"| {VELD_LABELS.get(veld, veld)} | {lengte} | {limiet} | {status} |")
        regels.append("")
        regels.append(f"Fasering: {len(p['fasering'])} regel(s).")
        regels.append("")
        if p.get("fouten"):
            regels.append("**❌ Fouten:**")
            for fo in p["fouten"]:
                regels.append(f"- {fo}")
        if p.get("aandachtspunten"):
            regels.append("**⚠️ Aandachtspunten:**")
            for aa in p["aandachtspunten"]:
                regels.append(f"- {aa}")
        regels.append("")
    doel.write_text("\n".join(regels), encoding="utf-8")


# Tekens die buiten de standaard WinAnsi-tekenset van PDF-formuliervelden vallen
# en in Adobe Reader als 'rare tekens' verschijnen. Vervangen door hun visuele
# equivalent; de inhoud/leesbaarheid blijft identiek. Accenten (é, ë, ï, enz.)
# zitten WEL in WinAnsi en blijven dus ongemoeid.
_PDF_TEKEN_VERVANGING = {
    "\u2013": "-",    # en-dash –  -> koppelteken
    "\u2014": "-",    # em-dash —  -> koppelteken
    "\u2011": "-",    # non-breaking hyphen ‑ -> koppelteken
    "\u2018": "'",    # left single quote ' -> rechte apostrof
    "\u2019": "'",    # right single quote ' -> rechte apostrof
    "\u201a": "'",    # single low-9 quote ‚
    "\u201c": '"',    # left double quote " -> recht aanhalingsteken
    "\u201d": '"',    # right double quote " -> recht aanhalingsteken
    "\u201e": '"',    # double low-9 quote „
    "\u00a0": " ",    # non-breaking space -> gewone spatie
    "\u2026": "...",  # ellipsis … -> drie punten
    "\u2022": "-",    # bullet • -> koppelteken
    "\u00ad": "",     # soft hyphen (onzichtbaar) -> weg
    "\u200b": "",     # zero-width space -> weg
}


def normaliseer_voor_pdf(tekst: str) -> str:
    """Vervang tekens die in PDF-formuliervelden weergaveproblemen geven.
    Wijzigt alleen de weergavevorm van leestekens, niet de inhoud."""
    if not tekst:
        return tekst
    for bron, doel in _PDF_TEKEN_VERVANGING.items():
        tekst = tekst.replace(bron, doel)
    return tekst


def vul_pdf_bytes(bron_pdf_bytes: bytes, velden: dict) -> bytes:
    """Zelfde als vul_pdf maar volledig in-memory: lege-formulier-bytes in,
    gevulde-pdf-bytes uit. Geschikt voor serverless gebruik."""
    import io
    writer = PdfWriter(clone_from=io.BytesIO(bron_pdf_bytes))
    try:
        writer.set_need_appearances_writer(True)
    except Exception:
        acro = writer._root_object["/AcroForm"]
        acro[NameObject("/NeedAppearances")] = BooleanObject(True)

    # Normaliseer alle tekstwaarden zodat ze correct in Adobe Reader tonen.
    velden = {
        k: (normaliseer_voor_pdf(v) if isinstance(v, str) and not v.startswith("/") else v)
        for k, v in velden.items()
    }
    velden["Samenwerking"] = "/SamenwerkingNee"
    for pagina in writer.pages:
        writer.update_page_form_field_values(pagina, velden, auto_regenerate=False)

    uit = io.BytesIO()
    writer.write(uit)
    return uit.getvalue()


def veilige_bestandsnaam(tekst: str) -> str:
    return re.sub(r"[^\w\-. ]", "_", tekst) or "project"


def converteer(docx_bytes: bytes, leeg_formulier_bytes: bytes, brondocument_naam: str = "",
               formuliertype: str = "programmatuur") -> dict:
    """Hoofdentree voor de Function.
    In:  bytes van het Word-document + bytes van het lege RVO-projectformulier
         + optioneel de bestandsnaam van het Word-document (voor de output-naamgeving)
         + formuliertype: 'programmatuur' (standaard) of 'fysiek'.
    Uit: dict met per project de gevulde pdf-bytes, bestandsnaam en signalering.
    """
    import io
    projecten = parse_docx(io.BytesIO(docx_bytes))

    # Basisnaam = naam van het Word-bestand zonder extensie, of "aanvraag" als fallback
    basis = re.sub(r"\.docx?$", "", brondocument_naam, flags=re.IGNORECASE).strip()
    basis = veilige_bestandsnaam(basis) if basis else "aanvraag"

    resultaten = []
    for p in projecten:
        velden = bouw_veldwaarden(p, formuliertype)
        if formuliertype == "fysiek":
            # Zwaartepunt staat standaard op 'Technisch nieuw product' (PDT).
            # Is het een productieproces, dan past de consultant dat handmatig aan.
            p["aandachtspunten"].append(
                "Zwaartepunt staat op 'Technisch nieuw product'. Is dit een "
                "productieproces, pas het zwaartepunt dan handmatig aan naar PPS.")
        pdf_bytes = vul_pdf_bytes(leeg_formulier_bytes, velden)
        naam = f"Projectformulier_{basis}_{veilige_bestandsnaam(normaliseer_projectcode(p['nummer']))}.pdf"
        resultaten.append({
            "projectnummer": p["nummer"],
            "projecttitel": p["titel"],
            "bestandsnaam": naam,
            "pdf_bytes": pdf_bytes,
            "fouten": p["fouten"],
            "aandachtspunten": p["aandachtspunten"],
        })
    return {"aantal_projecten": len(resultaten), "projecten": resultaten}
