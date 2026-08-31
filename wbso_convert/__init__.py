"""
Azure Function: WBSO Word-aanvraag -> gevulde RVO-projectformulieren.

HTTP-trigger. Ontvangt een Word-document, vult per project het lege
RVO-projectformulier (v1.8) en geeft de gevulde pdf's terug.

Request (JSON):
{
    "bestandsnaam": "aanvraag_klantX.docx",   # optioneel, voor nette naamgeving
    "docx_base64": "<base64 van het .docx-bestand>"
}

Response (JSON):
{
    "aantal_projecten": 2,
    "projecten": [
        {
            "projectnummer": "XX-01",
            "projecttitel": "...",
            "bestandsnaam": "WBSO_XX-01.pdf",
            "pdf_base64": "<base64 van de gevulde pdf>",
            "waarschuwingen": ["..."]
        }
    ]
}

Het lege projectformulier zit als vast bestand naast deze code
(leeg_formulier.pdf), zodat de Function zelfstandig werkt.
"""

import base64
import json
import logging
import os

import azure.functions as func

from .wbso_core import converteer

FORMULIER_PAD = os.path.join(os.path.dirname(__file__), "leeg_formulier.pdf")
FORMULIER_PAD_FYSIEK = os.path.join(os.path.dirname(__file__), "leeg_formulier_fysiek.pdf")


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("WBSO-conversie aangeroepen.")

    try:
        body = req.get_json()
    except ValueError:
        return _fout("Geen geldige JSON in de request body.", 400)

    docx_b64 = body.get("docx_base64")
    if not docx_b64:
        return _fout("Veld 'docx_base64' ontbreekt in de request.", 400)

    try:
        docx_bytes = base64.b64decode(docx_b64)
    except Exception:
        return _fout("Kon 'docx_base64' niet decoderen.", 400)

    # Formuliertype: 'programmatuur' (standaard) of 'fysiek'. De flow geeft dit
    # mee op basis van de map waarin het document is geplaatst.
    formuliertype = (body.get("formuliertype") or "programmatuur").strip().lower()
    if formuliertype not in ("programmatuur", "fysiek"):
        formuliertype = "programmatuur"
    pad = FORMULIER_PAD_FYSIEK if formuliertype == "fysiek" else FORMULIER_PAD

    try:
        with open(pad, "rb") as f:
            formulier_bytes = f.read()
    except FileNotFoundError:
        return _fout("Leeg projectformulier niet gevonden op de server.", 500)

    brondocument_naam = body.get("bestandsnaam", "")

    try:
        resultaat = converteer(docx_bytes, formulier_bytes, brondocument_naam, formuliertype)
    except Exception as e:
        logging.exception("Conversie mislukt.")
        return _fout(f"Conversie mislukt: {e}", 500)

    if resultaat["aantal_projecten"] == 0:
        return _fout("Geen projecten gevonden in het Word-document. "
                     "Controleer of het document de juiste structuur heeft.", 422)

    # pdf-bytes omzetten naar base64 voor transport
    for p in resultaat["projecten"]:
        p["pdf_base64"] = base64.b64encode(p.pop("pdf_bytes")).decode("ascii")

    return func.HttpResponse(
        json.dumps(resultaat, ensure_ascii=False),
        status_code=200,
        mimetype="application/json",
    )


def _fout(bericht: str, code: int) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps({"fout": bericht}, ensure_ascii=False),
        status_code=code,
        mimetype="application/json",
    )
