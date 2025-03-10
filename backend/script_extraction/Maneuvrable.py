from backend.script_extraction.ScriptAnalyzerABS import BaseAnalyzer


class ManeuvrableAnalyzer(BaseAnalyzer):
    """
    Analyseur pour déterminer si les objets sont manœuvrables.
    """
    
    def __init__(self, input=None, output=None, ws=None, wb=None):
        """
        Initialise l'analyseur de manœuvrabilité.
        
        Args:
            input (str, optional): Chemin du dossier contenant les fichiers à analyser.
            output (str, optional): Chemin de sortie pour les résultats.
            ws (worksheet, optional): Feuille de calcul Excel pour l'export.
            wb (workbook, optional): Classeur Excel pour l'export.
        """
        super().__init__(input, output, ws, wb)
        self.maneuvrable_count = 0
        self.non_maneuvrable_count = 0
        self.na_count = 0  # Nouveau compteur pour N/A
    
    def extract_maneuvrable_status(self, file_path):
        """
        Extrait le statut de manœuvrabilité des objets dans un fichier.
        
        Args:
            file_path (str): Chemin du fichier à analyser.
            
        Returns:
            list: Liste contenant les statuts de manœuvrabilité [objet1, objet2]
        """
        status2 = None
        current_object = None
        
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                for line in file:
                    line = line.strip()
                    
                    # Identifier la section d'objet actuelle
                    if line.startswith("OBJECT"):
                        if "OBJECT2" in line:
                            current_object = "OBJECT2"
                    
                    # Chercher la valeur MANEUVERABLE dans la section actuelle
                    elif "MANEUVERABLE" in line and current_object == "OBJECT2":
                        parts = line.split('=', 1)
                        if len(parts) > 1:
                            status2 = parts[1].strip().upper()  # Convertir à la casse supérieure pour uniformiser
                            break  # Sort de la boucle une fois la valeur trouvée
        
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier {file_path}: {e}")
        
        # Normaliser les valeurs
        return [
            self._normalize_maneuvrable_status(status2)
        ]

    def _normalize_maneuvrable_status(self, status):
        """
        Normalise le statut de manœuvrabilité.
        
        Args:
            status (str): Statut brut extrait du fichier.
            
        Returns:
            str: Statut normalisé ('YES', 'N/A', 'NO')
        """            
        
        if status == "YES":
            return "YES"
        elif status in ["N/A", "N A", "N/A "]:  
            return "N/A"
        elif status == "NO":
            return "NO"
        else:
            return "UNKNOWN"  

    def process_data(self):
        """
        Traite les données de manœuvrabilité et compte les différents types.
        
        Returns:
            list: Liste contenant les comptages [manœuvrables, non manœuvrables, N/A]
        """
        self.maneuvrable_count = 0
        self.non_maneuvrable_count = 0
        self.na_count = 0
        
        for file_path in self.get_all_files():
            statuses = self.extract_maneuvrable_status(file_path)
            
            for status in statuses:
                if status == "YES":
                    self.maneuvrable_count += 1
                elif status == "NO":
                    self.non_maneuvrable_count += 1
                elif status == "N/A":
                    self.na_count += 1
        
        return [self.maneuvrable_count, self.na_count, self.non_maneuvrable_count]
    
    def export_to_excel(self):
        """
        Exporte les résultats vers Excel.
        
        Returns:
            bool: True si l'export s'est bien déroulé, False sinon.
        """
        if self.ws is None or self.wb is None:
            print("Feuille Excel ou classeur non défini.")
            return False
        
        try:
            # Mettre à jour les statistiques
            stats = self.process_data()
            
            # Écrire dans la feuille Excel
            self.ws['Z3'] = stats[0]  # Manœuvrables
            self.ws['Z4'] = stats[1]  # Non manœuvrables
            self.ws['Z5'] = stats[2]  # N/A
            
            # Sauvegarder si chemin de sortie défini
            if self.output:
                self.wb.save(self.output)
            
            return True
            
        except Exception as e:
            print(f"Erreur lors de l'export Excel: {e}")
            return False