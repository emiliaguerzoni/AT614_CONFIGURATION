"""Modulo per la gestione centralizzata dei tooltip nell'interfaccia."""

from __future__ import annotations

from at614_editor.domain.schemas import get_test_schema_catalog
from at614_editor.domain.test_schema import TestParameterDescriptor

# Descrizioni predefinite per i campi comuni
COMMON_FIELD_TOOLTIPS = {
    "nome_test": "Nome descrittivo del test. Usato per identificare rapidamente il test nella sequenza.",
    "test_id": "Identificatore del tipo di test. Determina i parametri specifici che potranno essere configurati.",
    "indice": "Indice numerico o espressione di incremento (i++, ++i). Determina l'ordine di esecuzione dei test.",
    "sezione": "Nome della sequenza test associata. Costituisce il collegamento logico con la sequenza.",
    "ce16": "Riferimento alla calibrazione CE16 (opzionale). Contiene i dati di calibrazione del sensore CE16.",
    "parametro_aggiuntivo": "Parametri aggiuntivi configurabili. Utilizzati per estendere la configurazione del distributore.",
    "curva_comando": "File della curva di comando. Definisce la relazione tra comando e stato attuatore.",
    "limite_inf": "File della curva limite inferiore. Stabilisce il limite minimo accettabile per il test.",
    "limite_sup": "File della curva limite superiore. Stabilisce il limite massimo accettabile per il test.",
    "rampa_xy": "File della rampa XY. Definisce l'andamento temporale della tensione applicata.",
    "file_parametri": "File di configurazione parametri. Contiene i parametri specifici del test.",
    "parametri": "File di configurazione dei parametri del modulo. Utilizzato per il test SCRITTURA PARAMETRI.",
    "punto_x": "Coordinata X del punto. Rappresenta l'asse indipendente della curva.",
    "punto_y": "Coordinata Y del punto. Rappresenta il valore della grandezza misurata.",
}


def get_tooltip_for_parameter(test_id: str, parameter_index: int) -> str | None:
    """Recupera il tooltip per un parametro specifico di un test dalla descrizione dello schema.

    Args:
        test_id: Identificatore del tipo di test (es. 'ACQUISIZIONE_CAN')
        parameter_index: Indice del parametro (1-based)

    Returns:
        La descrizione del parametro come tooltip, oppure None se non disponibile
    """
    schema = get_test_schema_catalog().get(test_id)
    if schema is None:
        return None

    for parameter in schema.parameters:
        if parameter.index == parameter_index:
            # Preferisci description, altrimenti usa label o name
            if parameter.description:
                return parameter.description
            if parameter.label:
                return parameter.label
            if parameter.name:
                return parameter.name

    return None


def get_tooltip_for_descriptor(descriptor: TestParameterDescriptor, test_id: str) -> str | None:
    """Recupera il tooltip per un parametro descritto tramite TestParameterDescriptor.

    Args:
        descriptor: Descrittore del parametro
        test_id: Identificatore del test

    Returns:
        La descrizione del parametro come tooltip
    """
    tooltip = get_tooltip_for_parameter(test_id, descriptor.parameter_index)

    # Se non c'è descrizione specifica, usa il label come fallback
    if tooltip is None and descriptor.label:
        tooltip = descriptor.label

    return tooltip or None


def get_common_tooltip(field_type: str) -> str:
    """Recupera il tooltip predefinito per un campo comune.

    Args:
        field_type: Tipo di campo ('nome_test', 'test_id', 'indice', ecc.)

    Returns:
        Il tooltip predefinito
    """
    return COMMON_FIELD_TOOLTIPS.get(field_type, "")
