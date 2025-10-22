#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Creer l'excel de mon activité"""

import pypff
import csv
from datetime import datetime, time, date, timedelta

import pandas as pd
from collections import defaultdict


# === LISTE DE DÉTECTION DES DOSSIERS outlook ===
# (Outlook en français ou anglais)
SENT_FOLDERS = ["Sent Items", "Éléments envoyés", "Envoyés"]


# -----------------------------
# CONFIGURATION UTILISATEUR
# -----------------------------
OST_FILE = "/media/dlerat/Expansion/SSI-portable/Outlook/Damien.Lerat@supersonicimagine.com - damien.lerat@supersonicimagine.com.ost"  # Chemin vers fichier PST/OST
CSV_FILE = "/media/dlerat/Expansion/SSI-portable/Outlook/calendrier.CSV"  # CSV exporté des réunions Outlook
JIRA_CSV_FILE = "/media/dlerat/Expansion/SSI-portable/Outlook/jira.csv"  # export jira des issues dont je suis le créateur

OUTPUT_FILE = "rapport_activite.xlsx"

# === PÉRIODE D'ANALYSE ===
START_DATE = datetime(2022, 1, 1)
END_DATE = datetime(2025, 12, 31)

#
HEURE_DEBUT_JOURNEE = time(9, 0)
HEURE_FIN_JOURNEE = time(18, 30)
DUREE_REDACTION_MAIL_MINUTES = 15
DUREE_CREATION_ISSUE_MINUTES = 30


# =========================================================
def extract_folder_messages(folder, target_names=None):
    """Récupère récursivement les messages selon le nom du dossier"""
    messages = []
    name = (folder.name or "").lower()

    # Si un filtre de dossier est appliqué (ex: Sent Items ou Calendar)
    if target_names and not any(tn.lower() in name for tn in target_names):
        # Explorer quand même les sous-dossiers (certains PST ont "Top of Personal Folders" etc.)
        for i in range(folder.number_of_sub_folders):
            sub = folder.get_sub_folder(i)
            messages.extend(extract_folder_messages(sub, target_names))
        return messages

    # Extraction des messages
    for i in range(folder.number_of_sub_messages):
        try:
            msg = folder.get_sub_message(i)
            sent_time = msg.client_submit_time or msg.delivery_time
            if sent_time and START_DATE <= sent_time <= END_DATE:
                messages.append(msg)
                print(
                    f"sender: {msg.sender_name} sent {sent_time} subject: {msg.subject}"
                )

        except Exception as e:
            print(f"exception message {e}")
            continue

    # Exploration récursive
    for i in range(folder.number_of_sub_folders):
        sub = folder.get_sub_folder(i)
        messages.extend(extract_folder_messages(sub, target_names))

    return messages


# =========================================================
def process_sent_items(pst_file):
    """Extrait les mails envoyés"""
    file = pypff.file()
    file.open(pst_file)

    root = file.get_root_folder()
    print("[*] Recherche des mails envoyés...")
    sent_msgs = extract_folder_messages(root, target_names=SENT_FOLDERS)
    print(f"[+] {len(sent_msgs)} mails envoyés trouvés.")
    file.close()

    data = []
    for msg in sent_msgs:
        try:
            sent_time = msg.client_submit_time or msg.delivery_time
            print(f"sent_time {sent_time}")
            data.append(
                {
                    "type": "mail",
                    "subject": msg.subject or "",
                    "sender": msg.sender_name or "",
                    "date Redaction": sent_time
                    - timedelta(minutes=DUREE_REDACTION_MAIL_MINUTES),
                    "date": sent_time,
                }
            )
        except Exception as e:
            print(f"exception sent items {e}")
            continue
    return data


# =========================================================
def parse_meetings(csv_file):
    """Lit le fichier CSV exporté d Outlook."""
    meetings = []
    with open(csv_file, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                subject = row["Subject"]
                start_date = datetime.strptime(row["Start Date"], "%m/%d/%Y").date()
                end_date = datetime.strptime(row["End Date"], "%m/%d/%Y").date()
                start_time = datetime.strptime(row["Start Time"], "%I:%M:%S %p").time()
                end_time = datetime.strptime(row["End Time"], "%I:%M:%S %p").time()
                start_dt = datetime.combine(start_date, start_time)
                end_dt = datetime.combine(end_date, end_time)
                meetings.append((start_dt, end_dt, subject))
            except Exception as e:
                print(f"Erreur parsing ligne: {e}")
    return meetings


# =========================================================
def build_daily_report(in_emails, in_meetings, in_issues):
    """Fusionne les mails et réunions par jour."""
    daily = defaultdict(lambda: {"emails": [], "meetings": [], "issues": []})

    for tmp_mail in in_emails:
        mail_date = tmp_mail["date"].date()
        if START_DATE <= tmp_mail["date"] <= END_DATE:
            daily[mail_date]["emails"].append((
                tmp_mail["date Redaction"].time(),
                tmp_mail["date"].time(),
                tmp_mail["subject"]),
            )

    for tmp_issue in in_issues:
        issue_dt = datetime.strptime(tmp_issue["Created"], "%d/%b/%y %I:%M %p")
        if START_DATE <= issue_dt <= END_DATE:
            issueDesc = tmp_issue["Issue key"] + " " + tmp_issue["Summary"]
            daily[issue_dt.date()]["issues"].append(
                (issue_dt.time(), issueDesc)
            )

    for meeting_start_dt, meeting_end_dt, meeting_subject in in_meetings:
        meeting_date = meeting_start_dt.date()
        if START_DATE <= meeting_start_dt <= END_DATE:
            daily[meeting_date]["meetings"].append((meeting_start_dt.time(), meeting_end_dt.time(), meeting_subject))

    rows = []
    for day_date, info in sorted(daily.items()):
        all_times = []
        summary_lines = []

        for start, end, subj in info["meetings"]:
            summary_lines.append(
                f"Réunion {start.strftime('%H:%M')}-{end.strftime('%H:%M')}: {subj}"
            )
            all_times += [start, end]

        for start, end, subj in info["emails"]:
            summary_lines.append(f"Mail {end.strftime('%H:%M')}: {subj}")
            all_times.append(end)

        for myTime, summary in info["issues"]:
            summary_lines.append(f"issue {myTime.strftime('%H:%M')}: {summary}")
            all_times.append(myTime)

        if all_times:
            start_day = min(all_times)
            end_day = max(all_times)
        else:
            start_day = end_day = None

        rows.append(
            {
                "Année": day_date.year,
                "Semaine": day_date.isocalendar()[1],
                "Date": day_date.strftime("%Y-%m-%d"),
                "Jour": day_date.strftime("%a"),
                "Résumé": "\n".join(summary_lines),
                "Début": start_day.strftime("%H:%M") if start_day else "",
                "Fin": end_day.strftime("%H:%M") if end_day else "",
            }
        )

    return pd.DataFrame(rows)


def lire_fichier_csv_JIRA(chemin_fichier):
    """
    Lit un fichier CSV et retourne une liste de dictionnaires.

    Args:
        chemin_fichier (str): Le chemin vers le fichier CSV

    Returns:
        list: Liste de dictionnaires contenant les données du CSV
    """
    try:
        with open(chemin_fichier, "r", encoding="utf-8") as fichier:
            lecteur = csv.DictReader(fichier, delimiter=",")

            # Liste pour stocker tous les enregistrements
            donnees = []

            for ligne in lecteur:
                # Créer un dictionnaire avec toutes les colonnes
                enregistrement = {
                    "Issue key": ligne.get("Issue key", ""),
                    "Issue id": ligne.get("Issue id", ""),
                    "Summary": ligne.get("Summary", ""),
                    "Creator": ligne.get("Creator", ""),
                    "Creator Id": ligne.get("Creator Id", ""),
                    "Reporter": ligne.get("Reporter", ""),
                    "Reporter Id": ligne.get("Reporter Id", ""),
                    "Created": ligne.get("Created", ""),
                    "Assignee": ligne.get("Reporter", ""),
                    "Assignee Id": ligne.get("Reporter Id", ""),
                    "Status": ligne.get("Status", ""),
                    "Original estimate": ligne.get("Original estimate", ""),
                    "Priority": ligne.get("Priority", ""),
                    "Custom field (Business Priority)": ligne.get(
                        "Custom field (Business Priority)", ""
                    ),
                    "Custom field (Postpone comment)": ligne.get(
                        "Custom field (Postpone comment)", ""
                    ),
                    "Updated": ligne.get("Updated", ""),
                    "Custom field (last closed date)": ligne.get(
                        "Custom field (last closed date)", ""
                    ),
                    "Time Spent": ligne.get("Time Spent", ""),
                    "Labels": ligne.get(
                        "Labels", ""
                    ),  # Note: il y a plusieurs colonnes Labels
                    "Custom field (Safety Priority)": ligne.get(
                        "Custom field (Safety Priority)", ""
                    ),
                    "Team Id": ligne.get(
                        "Team Id", ""
                    ),  # Note: il y a plusieurs colonnes Labels
                    "Team Name": ligne.get(
                        "Team Name", ""
                    ),  # Note: il y a plusieurs colonnes Labels
                    "Custom field (Product version)": ligne.get(
                        "Custom field (Product version)", ""
                    ),
                    "Custom field (To be fixed in product version)": ligne.get(
                        "Custom field (To be fixed in product version)", ""
                    ),
                }
                donnees.append(enregistrement)
            return donnees
    except FileNotFoundError:
        print(f"Erreur: Le fichier '{chemin_fichier}' n'a pas été trouvé.")
        return []
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier: {e}")
        return []


# -----------------------------
# EXÉCUTION
# -----------------------------
if __name__ == "__main__":
    print("Lecture des mails envoyés...")
    emails = process_sent_items(OST_FILE)
    print(f"{len(emails)} mails trouvés")

    print("Lecture du calendrier CSV...")
    meetings = parse_meetings(CSV_FILE)
    print(f"{len(meetings)} réunions trouvées")

    print("Lecture des issues jira créées...")
    issues = lire_fichier_csv_JIRA(JIRA_CSV_FILE)
    print(f"{len(issues)} issues trouvés")

    print("Génération du rapport...")
    df = build_daily_report(emails, meetings, issues)

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Rapport généré : {OUTPUT_FILE}")
