from datetime import datetime
from collections import Counter
from backend.script_extraction.ScriptAnalyzerABS import BaseAnalyzer


class CountryLaunch(BaseAnalyzer):
    """
    Classe pour classifier les objets selon leur année de lancement.
    """
    
    def __init__(self, input=None, output=None, ws=None, wb=None):
        super().__init__(input, output, ws, wb)
        self._results = None
    
    def extract_launch_year(self, file_path):
        """Extrait l'année de lancement depuis l'INTERNATIONAL_DESIGNATOR."""
        designator = self.extract_value(file_path, "INTERNATIONAL_DESIGNATOR")
        if designator:
            try:
                return int(designator.split('-')[0])  # Extraction de l'année
            except ValueError:
                return None
        return None
    
    def count_launches_by_year(self):
        """Compte le nombre de satellites lancés par année."""
        results = self.analyze_files(extraction_func=self.extract_launch_year)
        
        # Filtrer les valeurs nulles et compter le nombre de lancements par année
        launch_counts = dict(Counter(filter(None, results)))
        
        return launch_counts

    def process_data(self):
        """
        Analyse tous les fichiers pour extraire l'année de lancement,
        compte le nombre de satellites par année et écrit les résultats
        dans la colonne 'AI' du fichier Excel.
        """
        launch_counts = self.count_launches_by_year()

        # Écriture des résultats dans la feuille Excel
        if self.ws:
            self.ws["AH1"] = "Année"
            self.ws["AI1"] = "Nombre de satellites"

            for idx, (year, count) in enumerate(sorted(launch_counts.items()), start=2):
                self.ws[f'AH{idx}'] = year
                self.ws[f'AI{idx}'] = count

            # Sauvegarde du fichier Excel
            self.wb.save(self.output)

        return launch_counts
