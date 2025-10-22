#!/bin/bash

VENV_DIR="OAR-venv/"
REQ_FILE="requirements.txt"

# Vérifie si python3-venv est installé
if ! dpkg -s python3-venv >/dev/null 2>&1; then
    echo "Le paquet python3-venv est manquant. Installation..."
    sudo apt update && sudo apt install -y python3-venv
fi

# Crée l'environnement virtuel s'il n'existe pas
if [ ! -d "$VENV_DIR" ]; then
    echo "Création de l'environnement virtuel dans $VENV_DIR"
    python3 -m venv "$VENV_DIR"
fi

# Active le venv
echo "Activation de l'environnement virtuel..."
source "$VENV_DIR/bin/activate"

# Met à jour pip
echo "Mise à jour de pip..."
pip install --upgrade pip

# Installe les dépendances si requirements.txt existe
if [ -f "$REQ_FILE" ]; then
    echo "Installation des dépendances depuis $REQ_FILE"
    pip install -r "$REQ_FILE"
else
    echo "Aucun fichier $REQ_FILE trouvé. Aucun paquet installé automatiquement."
fi

echo "✅ Environnement prêt. Pour le réactiver plus tard : source $VENV_DIR/bin/activate"
