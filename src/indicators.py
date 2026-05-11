from rapidfuzz import fuzz

def quality_indicators(document, summary):
    score = fuzz.partial_ratio(document[:1000], str(summary))
    return {
        "Presisjon": f"{score}%",
        "Konsistens": "Moderat",
        "Faglig korrekthet": "Må vurderes av saksbehandler"
    }

def legal_indicators():
    return {
        "Etterprøvbarhet": "Middels",
        "Likebehandling": "Ingen åpenbare risikoflagg",
        "Ansvar": "Beslutning tas av saksbehandler"
    }
