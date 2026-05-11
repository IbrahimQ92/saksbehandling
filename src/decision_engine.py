def evaluate_application(data):

    avstand_ok = data["avstand"] >= 4
    areal_ok = data["areal"] <= 50

    summary = f"""
Søknaden gjelder oppføring av {data["byggtype"]}.
Byggets størrelse er {data["areal"]} m².
Avstand til nabogrense er {data["avstand"]} meter.
"""

    if avstand_ok and areal_ok:
        decision = "Forslag: Søknaden kan godkjennes etter plan- og bygningsloven."
    elif not avstand_ok:
        decision = "For liten avstand til nabogrense."
    else:
        decision = "Forslag: Vurdering nødvendig – bygget overstiger vanlig størrelse."

    indicators = {
        "struktur": "OK",
        "regelreferanse": "Plan- og bygningsloven §29-4",
        
    }

    return {
        "summary": summary,
        "decision": decision,
        "indicators": indicators
    }
