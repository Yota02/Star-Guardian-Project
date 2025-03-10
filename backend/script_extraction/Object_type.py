from backend.script_extraction.ScriptAnalyzerABS import BaseAnalyzer


class ObjectTypeAnalyzer(BaseAnalyzer):
    """
    Classe pour analyser les types d'objets dans des fichiers TXT.
    """
    
    def __init__(self, input=None, output=None, ws=None, wb=None):
        """
        Initialise l'analyseur de types d'objets.
        """
        super().__init__(input, output, ws, wb)
        self._object_types = None
    
    def extract_object_data(self, file_path):
        """
        Extrait le type d'objet et le désignateur international.
        
        Returns:
            tuple: (object_type, designator) ou (None, None) si non trouvé.
        """
        object_type = self.extract_value(file_path, "OBJECT_TYPE", section="OBJECT2")
        designator = self.extract_value(file_path, "INTERNATIONAL_DESIGNATOR", section="OBJECT2")
        
        if object_type and object_type != "NONE":
            return (object_type, designator)
        return (None, None)
    
    def process_data(self):
        """
        Analyse tous les fichiers et extrait les types d'objets.
        """
        # Fonction qui filtre les résultats nuls
        def is_valid(result):
            return result[0] is not None and result[1] is not None
        
        # Extraire les paires (type, designator)
        results = self.analyze_files(
            extraction_func=self.extract_object_data, 
            filter_func=is_valid,
            unique_key=1  # Le designator est l'élément à l'index 1
        )
        
        # Extraire seulement les types d'objets
        self._object_types = [result[0] for result in results]
    
    def analyze_folder(self):
        """
        Analyse tous les fichiers dans le dossier spécifié et retourne la liste des types d'objets.
        """
        if self._object_types is None:
            self.process_data()
        return self._object_types
    
    def nombre_Payload_Debris(self):
        """
        Compte le nombre de PAYLOAD et DEBRIS dans le dossier spécifié.
        """
        if self._object_types is None:
            self.process_data()
        
        # Comptage des occurrences de "PAYLOAD" et "DEBRIS"
        payload_count = self._object_types.count("PAYLOAD")
        debris_count = self._object_types.count("DEBRIS")
        
        return payload_count, debris_count
    
    def plot_object_distribution(self):
        """
        Génère un graphique à barres de la distribution des types d'objets.
        """
        import matplotlib.pyplot as plt
        from collections import Counter
        
        if self._object_types is None:
            self.process_data()
        
        # Compter les occurrences de chaque type d'objet
        counter = Counter(self._object_types)
        
        # Créer le graphique
        labels = list(counter.keys())
        values = list(counter.values())
        
        plt.figure(figsize=(10, 6))
        plt.bar(labels, values, color='skyblue')
        plt.xlabel('Type d\'objet')
        plt.ylabel('Nombre')
        plt.title('Distribution des types d\'objets')
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        return plt