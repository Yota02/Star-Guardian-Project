from collections import Counter
import os
import re
import matplotlib.pyplot as plt

class GenerateurGraphique:
    
    def __init__(self):
        self.compteur = 0
        
        # Définition des couleurs
        self.colors = [ 
            '#0066CC',  # Bleu foncé
            '#33A1FF',  # Bleu clair
            '#00E6E6',  # Turquoise
            '#00CC99',  # Vert turquoise
            '#66CC33'   # Vert
        ]
        
        # Attributs par défaut, modifiables si nécessaire
        self.dossier_sortie_defaut = "./Star_Guardian/output/graphique"
        if not os.path.exists(self.dossier_sortie_defaut):
            os.makedirs(self.dossier_sortie_defaut)  # Crée le dossier par défaut s'il n'existe pas

    def recuperer_titre_graphique_fichier(self, directory):
        if not os.path.exists(directory):
            print(f"Le répertoire {directory} n'existe pas.")
            return None

        # Parcours des fichiers dans le répertoire
        for filename in os.listdir(directory):
            filepath = os.path.join(directory, filename)

            # Vérifie si c'est un fichier
            if os.path.isfile(filepath):
                try:
                    # Lecture du contenu du fichier
                    with open(filepath, 'r', encoding='utf-8') as file:
                        content = file.read()

                    # Recherche du premier OBJECT_NAME
                    match = re.search(r'OBJECT_NAME\s*=\s*(.+)', content)
                    if match:
                        # Nettoyage du nom pour retirer les caractères inutiles
                        satellite_name = match.group(1).strip()
                        satellite_name = re.sub(r'[^\w\s-]', '', satellite_name)  # Retirer caractères non souhaités
                        return satellite_name
                    else:
                        print(f"OBJECT_NAME introuvable dans {filename}.")
                        return None

                except Exception as e:
                    print(f"Erreur lors du traitement du fichier {filename}: {e}")
                    return None
        return None
    
    def demande_dossier_sortie(self):
        dossier = input("Entrez le chemin du dossier où enregistrer l'image (ou appuyez sur Entrée pour le dossier par défaut) : ")
        
        if not dossier:  # Si aucun chemin n'est fourni, utiliser le dossier par défaut
            dossier = self.dossier_sortie_defaut
            print(f"Utilisation du dossier par défaut : {dossier}")
        
        if not os.path.isdir(dossier):  # Vérifier si le dossier existe
            print(f"Le dossier spécifié '{dossier}' n'existe pas. Création du dossier...")
            os.makedirs(dossier)  # Crée le dossier s'il n'existe pas
        return dossier

    def demande_nom_fichier_sortie(self):
        nom_fichier = input("Entrez le nom du fichier (avec extension, ex. graphique.png (par defaut)) : ")
        
        if not nom_fichier:
            nom_fichier = "graphique.png"  # Nom par défaut
            print(f"Utilisation du nom de fichier par défaut : {nom_fichier}")
        
        return nom_fichier

    def generer_graphique_pays(self, countries, title, demande):
        if(demande):
            dossier = self.demande_dossier_sortie() 
            nom_fichier = self.demande_nom_fichier_sortie()
            chemin_complet = os.path.join(dossier, nom_fichier)  
        else :
            chemin_complet = "./Star_Guardian/output/graphique/graphiqueCountry.png"
        
        # Calculer la distribution des pays
        counter = Counter(countries)
        labels, sizes = zip(*counter.items())  # Extraire les labels et leurs proportions
        
        # Demander les informations d'enregistrement
        dossier = self.demande_dossier_sortie()  # Demander le dossier de sauvegarde à l'utilisateur
        nom_fichier = self.demande_nom_fichier_sortie()  # Demander le nom du fichier à l'utilisateur
        chemin_complet = os.path.join(dossier, nom_fichier)  # Construire le chemin complet

        # Générer le graphique
        plt.figure(figsize=(8, 6))
        plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=self.colors)
        plt.title(title)
        plt.axis('equal')

        # Sauvegarder et fermer la figure
        plt.savefig(chemin_complet)
        plt.close()
        print(f"Graphique sauvegardé dans : {chemin_complet}")

    def generer_graphique_type(self, data, title, demande):
        if(demande):
            dossier = self.demande_dossier_sortie() 
            nom_fichier = self.demande_nom_fichier_sortie()
            chemin_complet = os.path.join(dossier, nom_fichier)  
        else :
            chemin_complet = "./Star_Guardian/output/graphique/graphiqueType.png"
            
        counts = Counter(data)  
        types = list(counts.keys())
        frequencies = list(counts.values())

        # Création d'un camembert
        plt.figure(figsize=(8, 8))
        plt.pie(frequencies, labels=types, autopct='%1.1f%%', startangle=90, colors=plt.cm.tab20.colors)
        plt.title(title)
        plt.axis('equal')  # Assure que le camembert est un cercle
        plt.savefig(chemin_complet)

    def generer_graphique_inclination(self, inclinations, title, demande):
        if(demande):
            dossier = self.demande_dossier_sortie() 
            nom_fichier = self.demande_nom_fichier_sortie()
            chemin_complet = os.path.join(dossier, nom_fichier)  
        else :
            chemin_complet = "./Star_Guardian/output/graphique/graphiqueInclination.png"
        
        if not inclinations:
            print("Aucune inclinaison trouvée dans les fichiers.")
            return
            
        counter = Counter(inclinations)
        
        plt.figure(figsize=(10, 8))
        plt.pie(counter.values(), 
                labels=[f'{k}°' for k in counter.keys()], 
                autopct='%1.1f%%',
                colors=self.colors)
        plt.title(title)
        plt.axis('equal')   
        plt.savefig(chemin_complet)
    
    def generer_graphique_avec_legend(self, data, title, categories_order, demande):
        if(demande):
            dossier = self.demande_dossier_sortie() 
            nom_fichier = self.demande_nom_fichier_sortie()
            chemin_complet = os.path.join(dossier, nom_fichier)  
        else :
            self.compteur = self.compteur + 1
            chemin_complet = "./Star_Guardian/output/graphique/graphique" + str(self.compteur) + ".png"

        # Prépare les données pour le graphique
        labels = []
        sizes = []
        for category in categories_order:
            count = data.count(category)
            if count > 0:  # N'inclut que les catégories non vides
                labels.append(category)
                sizes.append(count)

        # Création du graphique en camembert
        plt.figure(figsize=(12, 8))  
        patches, texts, autotexts = plt.pie(sizes,
                                            labels=labels,
                                            colors=self.colors[:len(labels)],
                                            autopct='%1.1f%%',
                                            startangle=90)

        plt.axis('equal')  # Égalise l'aspect du camembert
        plt.title(title, fontsize=14, pad=20)

        # Ajout de la légende
        plt.legend(patches, labels,
                title="Catégories",
                loc="center left",
                bbox_to_anchor=(1.2, 0.5))

        plt.tight_layout()
        plt.savefig(chemin_complet, bbox_inches='tight', facecolor='white')  # Sauvegarde du graphique
        plt.close()