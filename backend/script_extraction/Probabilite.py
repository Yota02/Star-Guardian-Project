import os
from collections import defaultdict

from backend.script_extraction.ScriptAnalyzerABS import BaseAnalyzer

class CollisionProbabilityAnalyzer(BaseAnalyzer):
    """
    Classe pour analyser les probabilités de collision dans des fichiers CDM.
    """
    
    def __init__(self, input, output, ws, wb):
        """
        Initialise l'analyseur de probabilités de collision.
        
        Args:
            dossier (str, optional): Chemin du dossier contenant les fichiers à analyser.
        """
        super().__init__(input, output, ws, wb)
        self.categories = [
            "≥1E-4",
            "1E-4>X≥1E-5",
            "1E-5>X≥1E-6",
            "1E-6>X≥1E-7",
            "≤1E-8",
            "Entre 1E-7 et 1E-8"
        ]
    
    def process_data(self):
        self.analyze_folder()
    
    def extract_probability(self, file_path):
        """
        Extrait la probabilité de collision d'un fichier CDM.
        
        Args:
            file_path (str): Chemin du fichier à analyser.
            
        Returns:
            float or None: La probabilité de collision ou None si non trouvée.
        """
        try:
            with open(file_path, 'r') as file:
                for line in file:
                    if 'COLLISION_PROBABILITY' in line:
                        probability = float(line.split('=')[1].strip())
                        return probability
            return None
        except Exception as e:
            print(f"Erreur lors de l'extraction de la probabilité: {e}")
            return None
    
    def classify_collision_probability(self, probability):
        """
        Classifie une probabilité de collision selon des seuils définis.
        
        Args:
            probability (float): La probabilité de collision à classifier.
            
        Returns:
            str: La catégorie de la probabilité.
        """
        if probability >= 1e-4:
            return "≥1E-4"
        elif 1e-4 > probability >= 1e-5:
            return "1E-4>X≥1E-5"
        elif 1e-5 > probability >= 1e-6:
            return "1E-5>X≥1E-6"
        elif 1e-6 > probability >= 1e-7:
            return "1E-6>X≥1E-7"
        elif probability <= 1e-8:
            return "≤1E-8"
        else:
            return "Entre 1E-7 et 1E-8"
    
    def nombre_donnee(self):
        """
        Compte le nombre de fichiers par catégorie de probabilité de collision.
        
        Returns:
            tuple: Un tuple contenant le nombre de fichiers pour chaque catégorie de probabilité.
        """

        categories_count = defaultdict(int)
        
        # Parcours des fichiers dans le dossier
        for file_name in os.listdir(self.input):
            if file_name.endswith(".txt"):  # Vérifie si c'est un fichier .txt
                file_path = os.path.join(self.input, file_name)  # Chemin complet du fichier
                
                # Extraction de la probabilité de collision
                probability = self.extract_probability(file_path)
                if probability is not None:  # Si une probabilité est trouvée
                    category = self.classify_collision_probability(probability)
                    categories_count[category] += 1  # Incrémente le compteur pour la catégorie
        
        # On s'assure que les catégories de probabilité sont présentes, sinon on les initialise à 0
        result = (
            categories_count.get("≥1E-4", 0),
            categories_count.get("1E-4>X≥1E-5", 0),
            categories_count.get("1E-5>X≥1E-6", 0),
            categories_count.get("1E-6>X≥1E-7", 0),
            categories_count.get("≤1E-8", 0)
        )
        
        # Retourne le tuple avec les comptages des catégories
        return result
    
    def analyze_folder(self):
        """
        Analyse tous les fichiers .txt dans le dossier et retourne une liste des catégories.
        
        Returns:
            list: Liste des catégories de probabilité trouvées.
        """

        # Liste pour stocker les catégories
        categories_list = []
        
        # Parcourt tous les fichiers .txt du dossier
        for filename in os.listdir(self.input):
            if filename.endswith('.txt'):
                file_path = os.path.join(self.input, filename)
                probability = self.extract_probability(file_path)
                
                if probability is not None:
                    category = self.classify_collision_probability(probability)
                    categories_list.append(category)  # Ajoute la catégorie à la liste
        
        return categories_list
    
    def get_category_counts(self):
        """
        Renvoie un dictionnaire avec le nombre de fichiers par catégorie.
        
        Returns:
            dict: Dictionnaire avec les catégories comme clés et le nombre de fichiers comme valeurs.
        """
        categories_list = self.analyze_folder()
        category_counts = defaultdict(int)
        
        for category in categories_list:
            category_counts[category] += 1
        
        # S'assurer que toutes les catégories sont présentes
        for category in self.categories:
            if category not in category_counts:
                category_counts[category] = 0
        
        return dict(category_counts)