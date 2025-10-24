#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Créer l'excel de mon activité quotidienne"""

import locale
from collections import defaultdict
import csv
import os
import re
from datetime import datetime, time, date, timedelta
from zoneinfo import ZoneInfo
import random

import pypff
import pandas as pd
import git_stat as git
import constants

# locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")  # si a paris
# paris_tz = pytz.timezone('Europe/Paris')
os.environ["TZ"] = "Europe/Paris"


# =========================================================
def extract_folder_messages(
    folder: pypff.folder, target_names: list[str] = None
) -> list[pypff.message]:
    """Récupère récursivement les messages selon le nom du dossier"""
    messages: list[pypff.message] = []
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
            sent_time = sent_time.replace(tzinfo=ZoneInfo("UTC")).astimezone(
                ZoneInfo("Europe/Paris")
            )
            if sent_time and constants.START_DATE <= sent_time <= constants.END_DATE:
                messages.append(msg)
                print(
                    f"Mail: sender: {msg.sender_name} sent {sent_time} subject: {msg.subject}"
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
def process_sent_items(pst_file: str) -> list[dict]:
    """Extrait les mails envoyés"""
    file: pypff.file = pypff.file()  # pylint: disable=no-member
    file.open(pst_file)

    root: pypff.folder = file.get_root_folder()
    print("[*] Recherche des mails envoyés...")
    sent_msgs = extract_folder_messages(root, target_names=constants.SENT_FOLDERS)
    print(f"[+] {len(sent_msgs)} mails envoyés trouvés.")
    file.close()

    data: list[dict] = []
    for msg in sent_msgs:
        try:
            # les heures des mails sont en utc il faut les convertirs en heure de paris
            sent_time: datetime = msg.client_submit_time or msg.delivery_time
            sent_time = sent_time.replace(tzinfo=ZoneInfo("UTC")).astimezone(
                ZoneInfo("Europe/Paris")
            )
            data.append(
                {
                    "type": "mail",
                    "subject": msg.subject or "",
                    "sender": msg.sender_name or "",
                    "date Redaction": sent_time
                    - timedelta(minutes=constants.DUREE_REDACTION_MAIL_MINUTES),
                    "date Envoi": sent_time,
                }
            )
        except Exception as e:
            print(f"exception sent items {e}")
            continue
    return data


# =========================================================
def parse_meetings(csv_file: str) -> list[dict]:
    """Lit le fichier CSV exporté d Outlook."""
    tmp_meetings = []
    with open(csv_file, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                subject = row["Subject"]
                start_date = datetime.strptime(row["Start Date"], "%m/%d/%Y").date()
                end_date = datetime.strptime(row["End Date"], "%m/%d/%Y").date()
                start_time = datetime.strptime(row["Start Time"], "%I:%M:%S %p").time()
                end_time = datetime.strptime(row["End Time"], "%I:%M:%S %p").time()
                start_dt = datetime.combine(start_date, start_time).replace(
                    tzinfo=ZoneInfo("Europe/Paris")
                )
                end_dt = datetime.combine(end_date, end_time).replace(
                    tzinfo=ZoneInfo("Europe/Paris")
                )
                tmp_meetings.append(
                    {"start_time": start_dt, "end_time": end_dt, "subject": subject}
                )

            except Exception as e:
                print(f"Erreur parsing ligne: {e}")
    return tmp_meetings


# =========================================================
def lire_fichier_csv_jira(chemin_fichier: str) -> list[dict]:
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


# =========================================================
def get_git_stats(repo_path: str) -> list[git.CommitInfo]:
    """Récupère les statistiques Git"""
    git_stats = git.GitCommitAnalyzer(repo_path)
    stats = git_stats.get_statistiques()
    commits = git_stats.get_commits_par_date(constants.START_DATE, constants.END_DATE)
    return commits


# =========================================================
def get_all_commits(repo_paths: list[str]) -> list[dict]:
    """Récupère tous les commits de tous les dépôts"""
    all_commits = []
    for repo_path in repo_paths:
        commits = get_git_stats(repo_path)
        for commit in commits:
            all_commits.append(
                {
                    "date": commit.date_heure,
                    "message": commit.message,
                    "author": commit.auteur,
                    "email": commit.email,
                    "sha1": commit.sha1,
                    "sha1_complet": commit.sha1_complet,
                    "timestamp": commit.timestamp,
                    "repository": os.path.basename(repo_path),
                }
            )

    return all_commits


# =========================================================
def build_daily_report(
    in_emails: list[dict],
    in_meetings: list[dict],
    in_issues: list[dict],
    in_commits: list[dict],
) -> pd.DataFrame:
    """Fusionne les mails et réunions par jour."""
    daily = defaultdict(
        lambda: {"emails": [], "meetings": [], "issues": [], "commits": []}
    )

    for tmp_mail in in_emails:
        mail_date = tmp_mail["date Envoi"].date()
        if constants.START_DATE <= tmp_mail["date Envoi"] <= constants.END_DATE:
            daily[mail_date]["emails"].append(
                (
                    tmp_mail["date Redaction"].time(),
                    tmp_mail["date Envoi"].time(),
                    tmp_mail["subject"],
                ),
            )

    for tmp_issue in in_issues:
        issue_dt = datetime.strptime(tmp_issue["Created"], "%d/%b/%y %I:%M %p").replace(
            tzinfo=ZoneInfo("Europe/Paris")
        )
        if constants.START_DATE <= issue_dt <= constants.END_DATE:
            issue_desc = tmp_issue["Issue key"] + " " + tmp_issue["Summary"]
            daily[issue_dt.date()]["issues"].append(
                (
                    (
                        issue_dt
                        - timedelta(minutes=constants.DUREE_CREATION_ISSUE_MINUTES)
                    ).time(),
                    issue_dt.time(),
                    issue_desc,
                )
            )

    for meeting in in_meetings:
        meeting_date = meeting["start_time"].date()
        if constants.START_DATE <= meeting["start_time"] <= constants.END_DATE:
            daily[meeting_date]["meetings"].append(
                (
                    meeting["start_time"].time(),
                    meeting["end_time"].time(),
                    meeting["subject"],
                )
            )

    for commit in in_commits:
        commit_dt = datetime.strptime(commit["date"], "%Y-%m-%d %H:%M:%S").replace(
            tzinfo=ZoneInfo("Europe/Paris")
        )
        commit_date = commit_dt.date()
        if constants.START_DATE <= commit_dt <= constants.END_DATE:
            daily[commit_date]["commits"].append(
                (
                    (
                        commit_dt - timedelta(minutes=constants.DUREE_COMMIT_MINUTES)
                    ).time(),
                    commit_dt.time(),
                    re.sub(
                        r"[\x00-\x08\x0B-\x0C\x0E-\x1F]",
                        "",
                        str(
                            commit["repository"]
                            + " "
                            + commit["sha1"]
                            + " "
                            + commit["message"]
                        ),
                    ),
                )
            )

    rows = []
    for day_date, info in sorted(daily.items()):
        all_times = []
        summary_lines: list[str] = []

        for start, end, subj in info["meetings"]:
            summary_lines.append(
                f"{start.strftime('%H:%M')}-{end.strftime('%H:%M')}: Réunion {subj}"
            )
            if start != end:  # exclure les meetings qui dure la journée entiere
                all_times += [start, end]

        for start, end, subj in info["commits"]:
            summary_lines.append(f"{end.strftime('%H:%M')}: commits {subj}")
            all_times += [start, end]

        for start, end, subj in info["emails"]:
            summary_lines.append(f"{end.strftime('%H:%M')}: Mail {subj}")
            all_times += [start, end]

        for start, end, summary in info["issues"]:
            summary_lines.append(f"{end.strftime('%H:%M')}: issue {summary}")
            all_times += [start, end]

        if all_times:
            all_times += [
                generer_heure_aleatoire(
                    constants.HEURE_DEBUT_JOURNEE, constants.HEURE__DELTA
                ),
                generer_heure_aleatoire(
                    constants.HEURE_FIN_JOURNEE, constants.HEURE__DELTA
                ),
            ]

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
                "Résumé": "\n".join(sorted(summary_lines)),
                "Début": start_day.strftime("%H:%M") if start_day else "",
                "Fin": end_day.strftime("%H:%M") if end_day else "",
            }
        )

    rows = completer_dates_manquantes(
        constants.START_DATE,
        constants.END_DATE,
        rows
    )

    return pd.DataFrame(rows)


# =========================================================
def completer_dates_manquantes(
    date_debut: datetime, date_fin: datetime, liste_dict: list[dict]
) -> list[dict]:
    """
    Complète la liste de dictionnaires avec toutes les dates manquantes entre date_debut et date_fin.

    Args:
        date_debut: Date de début de la période
        date_fin: Date de fin de la période
        liste_dict: Liste de dictionnaires contenant une clé "Date" au format "%Y-%m-%d"

    Returns:
        Liste complétée avec toutes les dates de la période
    """
    # Convertir les datetime en date pour la comparaison
    debut_date = date_debut.date()
    fin_date = date_fin.date()

    # Créer un set des dates déjà présentes pour une recherche rapide
    dates_existantes = set()
    for item in liste_dict:
        try:
            date_obj = datetime.strptime(item["Date"], "%Y-%m-%d").date()
            dates_existantes.add(date_obj)
        except (KeyError, ValueError):
            continue

    # Créer la liste complète de toutes les dates de la période
    liste_complete = liste_dict.copy()
    date_courante = debut_date

    while date_courante <= fin_date:
        if date_courante not in dates_existantes:
            # Créer un nouveau dictionnaire pour la date manquante
            nouveau_dict = {
                "Année": date_courante.year,
                "Semaine": date_courante.isocalendar()[1],
                "Date": date_courante.strftime("%Y-%m-%d"),
                "Jour": date_courante.strftime("%a"),
                "Résumé": "",
                "Début": "",
                "Fin": "",
            }
            liste_complete.append(nouveau_dict)

        date_courante += timedelta(days=1)

    # Trier la liste par date
    liste_complete_triee = sorted(
        liste_complete, key=lambda x: datetime.strptime(x["Date"], "%Y-%m-%d")
    )

    return liste_complete_triee


# =========================================================
def generer_heure_aleatoire(heure_base: time, delta_minutes: int) -> time:
    """Génère une heure aléatoire dans l'intervalle [heure_base - delta, heure_base + delta]"""
    base_datetime = datetime.combine(date.today(), heure_base)
    delta = random.randint(-delta_minutes, delta_minutes)
    return (base_datetime + timedelta(minutes=delta)).time()


# -----------------------------
# EXÉCUTION
# -----------------------------
if __name__ == "__main__":
    print("Lecture des mails envoyés...")
    emails = process_sent_items(constants.OST_FILE)
    print(f"\t{len(emails)} mails trouvés")

    print("Lecture du calendrier CSV...")
    meetings = parse_meetings(constants.CSV_FILE)
    print(f"\t{len(meetings)} réunions trouvées")

    print("Lecture des issues jira créées...")
    issues = lire_fichier_csv_jira(constants.JIRA_CSV_FILE)
    print(f"\t{len(issues)} issues trouvés")

    print("Lecture des repository git ...")
    commits = get_all_commits(constants.REPO_PATHS)
    print(f"\t{len(commits)} commits trouvés")

    print("Génération du rapport...")
    df = build_daily_report(emails, meetings, issues, commits)

    df.to_excel(constants.OUTPUT_FILE, index=False)
    print(f"✅ Rapport généré : {constants.OUTPUT_FILE}")
