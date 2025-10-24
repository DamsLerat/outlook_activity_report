
from datetime import datetime, time
from zoneinfo import ZoneInfo

# === LISTE DE DÉTECTION DES DOSSIERS outlook ===
# (Outlook en français ou anglais)
SENT_FOLDERS = ["Sent Items", "Éléments envoyés", "Envoyés"]


# -----------------------------
# CONFIGURATION Fichiers UTILISATEUR
# -----------------------------
OST_FILE: str = (
    "/media/dlerat/Expansion/SSI-portable/Outlook/Damien.Lerat@supersonicimagine.com - damien.lerat@supersonicimagine.com.ost"  # Chemin vers fichier PST/OST
)
CSV_FILE: str = (
    "/media/dlerat/Expansion/SSI-portable/Outlook/calendrier2.CSV"  # CSV exporté des réunions Outlook
)
JIRA_CSV_FILE: str = (
    "/media/dlerat/Expansion/SSI-portable/Outlook/jira.csv"  # export jira des issues dont je suis le créateur
)

REPO_PATHS: list[str] = [
    "/home/dlerat/git/debian-security-analyzer",
    "/home/dlerat/git/DicomNemaParsingTool",
    "/home/dlerat/git/my-mindmaps",
    "/home/dlerat/git/RubiThreatModel",
    "/home/dlerat/git/RubiThreatModel",
]
OUTPUT_FILE: str = "rapport_activite.xlsx"  # Chemin vers fichier Excel de sortie

# === PÉRIODE D'ANALYSE ===
START_DATE: datetime = datetime(2022, 1, 1, tzinfo=ZoneInfo("Europe/Paris"))
END_DATE: datetime = datetime(2025, 12, 31, tzinfo=ZoneInfo("Europe/Paris"))

#
HEURE_DEBUT_JOURNEE: time = time(9, 0)
HEURE_FIN_JOURNEE: time = time(20, 0)
HEURE__DELTA: int = 15
TPS_PAUSE_MINUTE: int = 60
TPS_PAUSE2_MINUTE: int = 120

DUREE_REDACTION_MAIL_MINUTES: int = 15
DUREE_CREATION_ISSUE_MINUTES: int = 30
DUREE_COMMIT_MINUTES: int = 120

SEUIL_MAIL = time(6, 00)
SEUIL_MEETING = time(7, 45)
SEUIL_ISSUE = time(6, 00)
SEUIL_COMMIT = time(6, 00)
