import os
import re
import shutil

def organize_files_by_satellite(source_dir, target_dir):
    """
    Organise les fichiers dans des dossiers basés sur le premier OBJECT_NAME trouvé et les place dans le dossier cible.
    Args:
        source_dir (str): Répertoire contenant les fichiers à organiser.
        target_dir (str): Répertoire où les fichiers triés seront placés.
    """
    if not os.path.exists(source_dir):
        print(f"Le répertoire {source_dir} n'existe pas.")
        return

    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    # Parcours des fichiers dans le répertoire source
    for filename in os.listdir(source_dir):
        filepath = os.path.join(source_dir, filename)

        # Vérifie si c'est un fichier
        if os.path.isfile(filepath):
            try:
                # Lecture du contenu du fichier
                with open(filepath, 'r', encoding='utf-8') as file:
                    content = file.read()

                # Recherche du premier OBJECT_NAME (expression tolérant des espaces variables)
                match = re.search(r'OBJECT_NAME\s*=\s*([^\s]+)', content)
                if match:
                    satellite_name = match.group(1)

                    # Création du dossier correspondant dans le répertoire cible
                    satellite_dir = os.path.join(target_dir, satellite_name)
                    os.makedirs(satellite_dir, exist_ok=True)

                    # Déplacement du fichier
                    shutil.move(filepath, os.path.join(satellite_dir, filename))
                    print(f"Fichier {filename} déplacé vers {satellite_dir}")
                else:
                    print(f"OBJECT_NAME introuvable dans {filename}.")

            except Exception as e:
                print(f"Erreur lors du traitement du fichier {filename}: {e}")