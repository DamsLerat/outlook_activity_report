import subprocess
import os
from typing import List, Dict, Optional, Union
from datetime import datetime, date
from dataclasses import dataclass


@dataclass
class CommitInfo:
    """Classe pour représenter les informations d'un commit"""

    sha1: str
    sha1_complet: str
    auteur: str
    email: str
    date_heure: str
    message: str
    timestamp: Optional[float] = None


class GitCommitAnalyzer:
    """
    Classe pour analyser les commits d'un repository Git
    """

    def __init__(self, repo_path: str):
        """
        Initialise l'analyseur avec le chemin du repository Git

        Args:
            repo_path (str): Chemin vers le repository Git
        """
        self.repo_path = repo_path
        self._validate_repository()

    def _validate_repository(self) -> None:
        """Vérifie que le chemin est un repository Git valide"""
        if not os.path.exists(os.path.join(self.repo_path, ".git")):
            raise ValueError(
                f"Le chemin {self.repo_path} n'est pas un repository Git valide"
            )

    def _executer_commande_git(self, commande: List[str]) -> str:
        """
        Exécute une commande Git et retourne la sortie

        Args:
            commande (List[str]): Commande Git à exécuter

        Returns:
            str: Sortie de la commande
        """
        try:
            result = subprocess.run(
                commande,
                cwd=self.repo_path,
                capture_output=True,
                text=True,
                check=True,
                encoding="utf-8",
            )
            return result.stdout
        except subprocess.CalledProcessError as e:
            if "does not have any commits" in e.stderr:
                return ""  # Aucun commit trouvé
            raise RuntimeError(f"Erreur Git: {e.stderr}")
        except FileNotFoundError:
            raise RuntimeError("Git n'est pas installé ou n'est pas dans le PATH")

    def _parser_ligne_commit(self, ligne: str) -> Optional[CommitInfo]:
        """
        Parse une ligne de sortie Git log en objet CommitInfo

        Args:
            ligne (str): Ligne de sortie du format personnalisé Git

        Returns:
            Optional[CommitInfo]: Objet CommitInfo ou None si parsing échoue
        """
        if not ligne.strip():
            return None

        parts = ligne.split("|", 4)
        if len(parts) != 5:
            return None

        sha1, auteur, email, date_heure, message = parts

        # Nettoyer et formater les données
        message_sur_une_ligne = message.replace("\n", " ").strip()

        try:
            dt = datetime.fromisoformat(date_heure.replace(" ", "T", 1))
            date_formatee = dt.strftime("%Y-%m-%d %H:%M:%S")
            timestamp = dt.timestamp()
        except ValueError:
            date_formatee = date_heure
            timestamp = None

        return CommitInfo(
            sha1=sha1[:8],
            sha1_complet=sha1,
            auteur=auteur,
            email=email,
            date_heure=date_formatee,
            message=message_sur_une_ligne,
            timestamp=timestamp,
        )

    def _convertir_date(self, date_input: Union[str, datetime, date]) -> str:
        """
        Convertit une date en format string pour Git

        Args:
            date_input: Date en string, datetime ou date

        Returns:
            str: Date formatée "YYYY-MM-DD"
        """
        if isinstance(date_input, datetime):
            return date_input.strftime("%Y-%m-%d")
        elif isinstance(date_input, date):
            return date_input.strftime("%Y-%m-%d")
        elif isinstance(date_input, str):
            try:
                datetime.strptime(date_input, "%Y-%m-%d")
                return date_input
            except ValueError:
                raise ValueError(
                    f"Format de date invalide: {date_input}. Utilisez 'YYYY-MM-DD'"
                )
        else:
            raise ValueError("Type de date non supporté")

    def _convertir_datetime(self, datetime_input: Union[str, datetime]) -> str:
        """
        Convertit un datetime en format string pour Git

        Args:
            datetime_input: Datetime en string ou datetime

        Returns:
            str: Datetime formaté "YYYY-MM-DD HH:MM:SS"
        """
        if isinstance(datetime_input, datetime):
            return datetime_input.strftime("%Y-%m-%d %H:%M:%S")
        elif isinstance(datetime_input, str):
            try:
                datetime.strptime(datetime_input, "%Y-%m-%d %H:%M:%S")
                return datetime_input
            except ValueError:
                try:
                    datetime.strptime(datetime_input, "%Y-%m-%d")
                    return datetime_input + " 00:00:00"
                except ValueError:
                    raise ValueError(
                        f"Format datetime invalide: {datetime_input}. Utilisez 'YYYY-MM-DD HH:MM:SS'"
                    )
        else:
            raise ValueError("Type datetime non supporté")

    def get_commits_par_date(
        self,
        date_debut: Union[str, datetime, date],
        date_fin: Union[str, datetime, date],
        auteur: Optional[str] = None,
        branche: str = "HEAD",
        ordre_chronologique: bool = True,
    ) -> List[CommitInfo]:
        """
        Récupère les commits entre deux dates

        Args:
            date_debut: Date de début (format: "YYYY-MM-DD" ou objet date/datetime)
            date_fin: Date de fin (format: "YYYY-MM-DD" ou objet date/datetime)
            auteur: Filtrer par auteur spécifique
            branche: Branche à examiner
            ordre_chronologique: Si True, trie du plus ancien au plus récent

        Returns:
            List[CommitInfo]: Liste des commits dans la période
        """
        date_debut_str = self._convertir_date(date_debut)
        date_fin_str = self._convertir_date(date_fin)

        cmd = [
            "git",
            "log",
            branche,
            "--pretty=format:%H|%an|%ae|%ad|%s",
            "--date=iso-local",
            f"--since={date_debut_str}",
            f"--until={date_fin_str}",
        ]

        if ordre_chronologique:
            cmd.append("--reverse")

        if auteur:
            cmd.append(f"--author={auteur}")

        sortie = self._executer_commande_git(cmd)
        return self._parser_sortie_commits(sortie)

    def get_commits_par_datetime(
        self,
        datetime_debut: Union[str, datetime],
        datetime_fin: Union[str, datetime],
        auteur: Optional[str] = None,
        branche: str = "HEAD",
        ordre_chronologique: bool = True,
    ) -> List[CommitInfo]:
        """
        Récupère les commits entre deux dates et heures précises

        Args:
            datetime_debut: Date et heure de début
            datetime_fin: Date et heure de fin
            auteur: Filtrer par auteur spécifique
            branche: Branche à examiner
            ordre_chronologique: Si True, trie du plus ancien au plus récent

        Returns:
            List[CommitInfo]: Liste des commits dans la période
        """
        datetime_debut_str = self._convertir_datetime(datetime_debut)
        datetime_fin_str = self._convertir_datetime(datetime_fin)

        cmd = [
            "git",
            "log",
            branche,
            "--pretty=format:%H|%an|%ae|%ad|%s",
            "--date=iso-local",
            f'--since="{datetime_debut_str}"',
            f'--until="{datetime_fin_str}"',
        ]

        if ordre_chronologique:
            cmd.append("--reverse")

        if auteur:
            cmd.append(f"--author={auteur}")

        sortie = self._executer_commande_git(cmd)
        return self._parser_sortie_commits(sortie)

    def get_derniers_commits(
        self, nombre: int = 10, branche: str = "HEAD"
    ) -> List[CommitInfo]:
        """
        Récupère les N derniers commits

        Args:
            nombre: Nombre de commits à récupérer
            branche: Branche à examiner

        Returns:
            List[CommitInfo]: Liste des N derniers commits
        """
        cmd = [
            "git",
            "log",
            branche,
            "--pretty=format:%H|%an|%ae|%ad|%s",
            "--date=iso-local",
            f"-n",
            str(nombre),
        ]

        sortie = self._executer_commande_git(cmd)
        return self._parser_sortie_commits(sortie)

    def get_commits_par_auteur(
        self, auteur: str, branche: str = "HEAD"
    ) -> List[CommitInfo]:
        """
        Récupère tous les commits d'un auteur spécifique

        Args:
            auteur: Nom de l'auteur
            branche: Branche à examiner

        Returns:
            List[CommitInfo]: Liste des commits de l'auteur
        """
        cmd = [
            "git",
            "log",
            branche,
            "--pretty=format:%H|%an|%ae|%ad|%s",
            "--date=iso-local",
            f"--author={auteur}",
        ]

        sortie = self._executer_commande_git(cmd)
        return self._parser_sortie_commits(sortie)

    def _parser_sortie_commits(self, sortie: str) -> List[CommitInfo]:
        """
        Parse la sortie de git log en liste d'objets CommitInfo

        Args:
            sortie (str): Sortie de la commande git log

        Returns:
            List[CommitInfo]: Liste des commits parsés
        """
        commits = []
        for ligne in sortie.strip().split("\n"):
            commit = self._parser_ligne_commit(ligne)
            if commit:
                commits.append(commit)
        return commits

    def get_statistiques(self) -> Dict[str, int]:
        """
        Retourne des statistiques basiques sur le repository

        Returns:
            Dict[str, int]: Statistiques du repository
        """
        # Nombre total de commits
        cmd_total = ["git", "rev-list", "--count", "HEAD"]
        total_commits = int(self._executer_commande_git(cmd_total).strip())

        # Nombre d'auteurs
        cmd_auteurs = ["git", "shortlog", "-s", "-n", "HEAD"]
        sortie_auteurs = self._executer_commande_git(cmd_auteurs)
        nb_auteurs = len(sortie_auteurs.strip().split("\n")) if sortie_auteurs else 0

        return {"total_commits": total_commits, "nombre_auteurs": nb_auteurs}


# Exemple d'utilisation
if __name__ == "__main__":
    try:
        # Initialisation
        analyseur = GitCommitAnalyzer("/chemin/vers/ton/repo")

        print("=== STATISTIQUES DU REPOSITORY ===")
        stats = analyseur.get_statistiques()
        print(f"Total commits: {stats['total_commits']}")
        print(f"Nombre d'auteurs: {stats['nombre_auteurs']}")
        print()

        # Exemple 1: Commits entre deux dates
        print("=== COMMITS ENTRE DEUX DATES ===")
        commits = analyseur.get_commits_par_date(
            date_debut="2024-01-01", date_fin="2024-12-31"
        )

        print(f"{len(commits)} commits trouvés:\n")
        for commit in commits:
            print(f"{commit.sha1} - {commit.auteur} - {commit.date_heure}")
            print(f"  {commit.message}")
            print()

        # Exemple 2: Derniers commits
        print("=== 10 DERNIERS COMMITS ===")
        derniers = analyseur.get_derniers_commits(10)
        for commit in derniers:
            print(f"{commit.sha1} - {commit.auteur} - {commit.date_heure}")
            print(f"  {commit.message}")
            print()

        # Exemple 3: Commits d'un auteur spécifique
        print("=== COMMITS PAR AUTEUR ===")
        commits_auteur = analyseur.get_commits_par_auteur("John Doe")
        print(f"{len(commits_auteur)} commits trouvés pour cet auteur")

        # Exemple 4: Avec objets datetime
        print("=== COMMITS AVEC DATETIME ===")
        from datetime import datetime, timedelta

        fin = datetime.now()
        debut = fin - timedelta(days=30)

        commits_recent = analyseur.get_commits_par_datetime(debut, fin)
        print(f"{len(commits_recent)} commits dans les 30 derniers jours")

    except Exception as e:
        print(f"Erreur: {e}")
