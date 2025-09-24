# mappings.py

# Procedure name mapping: SQL procedure name -> list of mapped PG procedure names
# Use this mapping in case any procedure name changes in PostgreSQL regarding the 63-character name restriction.
PROCEDURE_NAME_MAP = {
    #TRAC Customised Procedures List ---------------------------------------------
    #'sql procedure name': ['postgres customised procedure name'],
       
    # ... (add all other mappings here as in your script) ...
}

# Event trigger name mapping: SQL event trigger name -> list of mapped PG event trigger names
EVENT_TRIGGER_NAME_MAP = {
    #"sql event trigger name": ["postgres event trigger name"],
    # ... (add all other mappings here as in your script) ...
}
