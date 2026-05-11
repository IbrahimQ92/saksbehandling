def generate_summary(document):
    return {
        "Sammendrag": "Dette er et KI-generert sammendrag basert på dokumentet.",
        "Nøkkelfakta": [
            "Saken gjelder en søknad",
            "Dokumentasjon er delvis vedlagt"
        ],
        "Uklart / mangler": [
            "Manglende opplysninger om inntekt"
        ]
    }

def generate_decision(document):
    return {
        "Vurdering": "Basert på dokumentet fremstår saken som ufullstendig.",
        "Forslag": "Det foreslås å be om tilleggsdokumentasjon.",
        "Må verifiseres": [
            "Kontroller regelverk",
            "Bekreft faktagrunnlag"
        ]
    }
