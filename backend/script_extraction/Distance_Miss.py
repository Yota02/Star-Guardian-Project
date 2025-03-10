from backend.script_extraction.ScriptAnalyzerABS import BaseAnalyzer


class MissDistanceAnalyzer(BaseAnalyzer):
    """
    Classe pour analyser les distances de rapprochement (MISS_DISTANCE) dans des fichiers TXT.
    """
    
    def __init__(self, input=None, output=None, ws=None, wb=None):
        """
        Initialise l'analyseur de distances de rapprochement.
        """
        super().__init__(input, output, ws, wb)
        
        self.categories = [
            "≤100m",
            "100m>X≥200m",
            "200m>X≥300m",
            "300m>X≥400m",
            "400m>X≥500m",
            "1000m>X",
            "Non classifié"
        ]
        self._results = None

    def extract_miss_distance(self, file_path):
        """
        Extrait la distance de rapprochement (MISS_DISTANCE) d'un fichier TXT.
        """
        return self.extract_numeric_value(file_path, "MISS_DISTANCE")
    
    def classify_miss_distance(self, distance):
        """
        Classifie une distance (MISS_DISTANCE) selon des catégories définies.
        """
        if distance <= 100:
            return "≤100m"
        elif 100 < distance <= 200:
            return "100m>X≥200m"
        elif 200 < distance <= 300:
            return "200m>X≥300m"
        elif 300 < distance <= 400:
            return "300m>X≥400m"
        elif 400 < distance <= 500:
            return "400m>X≥500m"
        elif distance > 1000:
            return "1000m>X"
        else:
            return "Non classifié"
    
    def process_data(self):
        """
        Analyse les distances de rapprochement dans les fichiers.
        """
        # Extraire et classifier les distances
        self._results = self.analyze_files(
            extraction_func=self.extract_miss_distance,
            classification_func=self.classify_miss_distance
        )
    
    def analyze_folder(self):
        """
        Analyse tous les fichiers .txt dans le dossier et retourne une liste des catégories.
        """
        if self._results is None:
            self.process_data()
        return self._results
    
    def get_category_counts(self):
        """
        Renvoie un dictionnaire avec le nombre de fichiers par catégorie.
        """
        if self._results is None:
            self.process_data()
        return self.count_by_category(self.categories, self._results)
    
    def get_distance_statistics(self):
        """
        Calcule des statistiques sur les distances de rapprochement.
        """
        raw_distances = self.analyze_files(extraction_func=self.extract_miss_distance)
        return self.get_statistics(raw_distances)
    
    