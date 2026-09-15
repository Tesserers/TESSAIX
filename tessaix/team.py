"""Datos de equipo para las plantillas de Human Capital y Human Capital+Finance.

Estos son los nombres/roles tal y como aparecen ya escritos en las
plantillas — no se generan con IA, se usan solo para decidir qué bloques
de diapositiva conservar (ver pptx_builder.py).
"""

# Equipo fijo de las propuestas de Human Capital (plantilla_hc.pptx):
# Eduardo Serrano, Manuel Pina y Edward Manrique siempre aparecen — no hay
# nada que elegir aquí.
HC_TEAM = ["Eduardo Serrano", "Manuel Pina", "Edward Manrique"]

# Equipo de las propuestas de Human Capital + Finance (plantilla_hc_finance.pptx):
# Manuel Pina, Mónica Mayoral y Edward Manrique siempre aparecen. Eduardo
# Serrano es opcional — se pregunta en el wizard.
HC_FINANCE_TEAM_BASE = ["Manuel Pina", "Mónica Mayoral", "Edward Manrique"]
HC_FINANCE_TEAM_WITH_EDUARDO = ["Manuel Pina", "Eduardo Serrano", "Mónica Mayoral", "Edward Manrique"]

WARRANTY_MONTHS_OPTIONS = [3, 6, 9]
DEFAULT_WARRANTY_MONTHS = 3

LOCATIONS_MADRID_ONLY = "madrid"
LOCATIONS_ALL = "all"
