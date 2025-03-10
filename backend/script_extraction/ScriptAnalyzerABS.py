from abc import ABC, abstractmethod


class BaseAnalyzer(ABC):
    """
    Classe abstraite mère pour les analyseurs de données.
    """
    
    def __init__(self, input=None, output=None, ws=None, wb=None):
        """
        Initialise la classe mère avec le chemin du dossier.
        
        Args:
            input (str, optional): Chemin du dossier contenant les fichiers à analyser.
            output (str, optional): Chemin de sortie pour les résultats.
            ws (worksheet, optional): Feuille de calcul Excel pour l'export.
            wb (workbook, optional): Classeur Excel pour l'export.
        """
        self.input = input
        self.output = output
        self.ws = ws
        self.wb = wb
    
    def extract_value(self, file_path, key, section=None, default=None):
        """
        Fonction générique pour extraire une valeur à partir d'une clé dans un fichier texte.
        
        Args:
            file_path (str): Chemin du fichier à analyser.
            key (str): Clé à rechercher (ex: "MISS_DISTANCE").
            section (str, optional): Section où chercher la clé (ex: "OBJECT2").
            default: Valeur par défaut si la clé n'est pas trouvée.
            
        Returns:
            Valeur extraite ou la valeur par défaut si non trouvée.
        """
        try:
            in_target_section = section is None  # Si section est None, on cherche partout
            
            with open(file_path, 'r', encoding='utf-8') as file:
                for line in file:
                    line = line.strip()
                    
                    # Vérifier si on entre dans la section cible
                    if section and line.startswith("OBJECT") and section in line:
                        in_target_section = True
                    # Vérifier si on sort de la section cible (ligne vide ou nouvelle section)
                    elif section and in_target_section and (line == "" or (line.startswith("OBJECT") and section not in line)):
                        in_target_section = False
                    
                    # Si on est dans la section cible (ou pas de section spécifiée) et que la clé est trouvée
                    if in_target_section and key in line:
                        # Extraire la valeur après le signe '='
                        parts = line.split('=', 1)
                        if len(parts) > 1:
                            value = parts[1].strip()
                            return value
            
            return default
        except Exception as e:
            print(f"Erreur lors de l'extraction de {key}: {e}")
            return default

            
    def extract_numeric_value(self, file_path, key, section=None, default=None):
        """
        Extrait une valeur numérique d'un fichier texte.
        
        Args:
            file_path (str): Chemin du fichier à analyser.
            key (str): Clé à rechercher (ex: "MISS_DISTANCE").
            section (str, optional): Section où chercher la clé (ex: "OBJECT2").
            default: Valeur par défaut si la clé n'est pas trouvée ou la conversion échoue.
            
        Returns:
            float: Valeur numérique extraite ou la valeur par défaut.
        """
        value = self.extract_value(file_path, key, section, default)
        
        if value is not None and value != default:
            # Nettoyer la valeur pour garder uniquement les chiffres et le point décimal
            clean_value = ''.join(c for c in value if c.isdigit() or c == '.')
            
            try:
                return float(clean_value)
            except ValueError:
                print(f"Impossible de convertir en nombre: {value}")
                
        return default
    
    def get_all_files(self, extension=".txt"):
        """
        Renvoie tous les fichiers avec l'extension spécifiée dans le dossier d'entrée.
        
        Args:
            extension (str): Extension des fichiers à filtrer.
            
        Returns:
            list: Liste des chemins complets des fichiers.
        """
        import os
        
        file_paths = []
        
        if self.input and os.path.isdir(self.input):
            for file_name in os.listdir(self.input):
                if file_name.endswith(extension):
                    file_path = os.path.join(self.input, file_name)
                    file_paths.append(file_path)
        
        return file_paths
    
    def analyze_files(self, extraction_func, classification_func=None, filter_func=None, unique_key=None):
        """
        Analyse tous les fichiers avec une fonction d'extraction et éventuellement une fonction de classification.
        
        Args:
            extraction_func: Fonction qui extrait des données d'un fichier. 
                           Doit accepter un chemin de fichier comme paramètre.
            classification_func: Fonction facultative qui classifie le résultat de l'extraction.
            filter_func: Fonction facultative qui filtre les résultats (renvoie True pour inclure, False pour exclure).
            unique_key: Clé facultative pour éviter les doublons (ex: "INTERNATIONAL_DESIGNATOR").
                      Si spécifiée, la fonction d'extraction doit retourner un dict ou un tuple.
            
        Returns:
            list: Liste des résultats extraits ou classifiés.
        """
        results = []
        unique_values = set()
        
        for file_path in self.get_all_files():
            # Extraire les données
            extracted_data = extraction_func(file_path)
            
            # Vérifier si on doit filtrer les doublons
            if unique_key is not None:
                # Si extraction_func retourne un dictionnaire
                if isinstance(extracted_data, dict):
                    unique_value = extracted_data.get(unique_key)
                # Si extraction_func retourne un tuple ou une liste
                elif isinstance(extracted_data, (tuple, list)) and len(extracted_data) > 1:
                    unique_value = extracted_data[1]  # Supposons que la seconde valeur est l'identifiant unique
                else:
                    unique_value = None
                
                # Si la valeur unique est déjà vue ou None, passer au fichier suivant
                if unique_value is None or unique_value in unique_values:
                    continue
                
                unique_values.add(unique_value)
            
            # Appliquer le filtre si spécifié
            if filter_func and not filter_func(extracted_data):
                continue
            
            # Appliquer la classification si spécifiée
            if classification_func and extracted_data is not None:
                classified_data = classification_func(extracted_data)
                results.append(classified_data)
            elif extracted_data is not None:
                results.append(extracted_data)
        
        return results
    
    def count_by_category(self, categories, data):
        """
        Compte le nombre d'occurrences par catégorie.
        
        Args:
            categories (list): Liste des catégories possibles.
            data (list): Données à compter.
            
        Returns:
            dict: Dictionnaire avec le nombre d'occurrences par catégorie.
        """
        from collections import defaultdict
        
        counts = defaultdict(int)
        
        for item in data:
            counts[item] += 1
        
        # S'assurer que toutes les catégories sont présentes
        result = {}
        for category in categories:
            result[category] = counts.get(category, 0)
        
        return result
    
    def get_statistics(self, data):
        """
        Calcule des statistiques sur une liste de valeurs numériques.
        
        Args:
            data (list): Liste de valeurs numériques.
            
        Returns:
            dict: Dictionnaire contenant des statistiques (min, max, moyenne, etc.)
        """
        if not data:
            return {
                "min": None,
                "max": None,
                "moyenne": None,
                "nombre_total": 0
            }
        
        numeric_data = [value for value in data if isinstance(value, (int, float))]
        
        if not numeric_data:
            return {
                "min": None,
                "max": None,
                "moyenne": None,
                "nombre_total": 0
            }
        
        return {
            "min": min(numeric_data),
            "max": max(numeric_data),
            "moyenne": sum(numeric_data) / len(numeric_data),
            "nombre_total": len(numeric_data)
        }
    
    @abstractmethod
    def process_data(self):
        """
        Méthode abstraite pour analyser le dossier.
        Doit être implémentée par les classes enfants.
        """
        pass