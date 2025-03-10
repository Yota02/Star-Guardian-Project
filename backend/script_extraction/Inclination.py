import os
import re
from collections import Counter
import statistics

from backend.script_extraction.Conjonction import ConjunctionAnalyzer
from backend.script_extraction.ScriptAnalyzerABS import BaseAnalyzer

class InclinationAnalyzer(BaseAnalyzer):
    """
    Classe pour analyser les données d'inclinaison des objets spatiaux
    à partir de fichiers texte et exporter les résultats.
    """
    
    def __init__(self, input, output, ws, wb):
        """Initialise l'analyseur d'inclinaison."""
        super().__init__(input, output, ws, wb)
    
    def export_to_excel(self, inclinations):
        """
        Exporte les données d'inclinaison dans une feuille de calcul Excel
        en nettoyant complètement la feuille avant l'insertion.
        
        Args:
            inclinations (list): Liste des inclinaisons groupées
        """
        # Calculer la distribution des inclinaisons
        counter = Counter(inclinations)
        
        try:
            
            # Nettoyer les colonnes T et U
            for row in range(3, self.ws.max_row + 1):
                self.ws[f'T{row}'] = None
                self.ws[f'U{row}'] = None
            
            # Écrire les données à partir de la ligne 3
            for i, (inclinaison, count) in enumerate(sorted(counter.items()), start=3):
                self.ws[f'T{i}'] = inclinaison
                self.ws[f'U{i}'] = count
            
            # Sauvegarder les modifications
            self.wb.save(self.output)
        
        except Exception as e:
            print(f"Erreur lors de l'export Excel: {str(e)}")
    
    def get_inclination_statistics(self, inclinations):
        """
        Calcule des statistiques sur les inclinaisons.
        
        Args:
            inclinations (list): Liste des inclinaisons
            
        Returns:
            dict: Dictionnaire contenant les statistiques
        """
        if not inclinations:
            return {
                "count": 0,
                "min": None,
                "max": None,
                "mean": None,
                "median": None
            }
            
        return {
            "count": len(inclinations),
            "min": min(inclinations),
            "max": max(inclinations),
            "mean": statistics.mean(inclinations),
            "median": statistics.median(inclinations)
        }
    
    def get_inclination_ranges(self, inclinations, bins=10):
        """
        Regroupe les inclinaisons en plages pour faciliter l'analyse.
        
        Args:
            inclinations (list): Liste des inclinaisons
            bins (int): Nombre de plages souhaitées
            
        Returns:
            dict: Dictionnaire {plage: nombre}
        """
        if not inclinations:
            return {}
            
        # Calculer la taille de chaque plage
        min_inc = min(inclinations)
        max_inc = max(inclinations)
        
        if min_inc == max_inc:
            return {f"{min_inc}": len(inclinations)}
            
        bin_size = (max_inc - min_inc) / bins
        
        # Initialiser le dictionnaire des plages
        ranges = {}
        
        # Remplir le dictionnaire
        for inc in inclinations:
            bin_index = min(int((inc - min_inc) / bin_size), bins - 1)
            lower_bound = min_inc + bin_index * bin_size
            upper_bound = min_inc + (bin_index + 1) * bin_size
            
            range_key = f"{lower_bound:.1f}-{upper_bound:.1f}"
            
            if range_key in ranges:
                ranges[range_key] += 1
            else:
                ranges[range_key] = 1
                
        return ranges
    
    def analyze_folder(self, conjunction_files=None):
        """
        Analyse le dossier pour extraire les inclinaisons des objets,
        optionnellement limité aux fichiers de conjonction.
        
        Args:
            conjunction_files (set, optional): Ensemble des fichiers de conjonction à analyser
            
        Returns:
            list: Liste des inclinaisons trouvées (triées par ordre croissant)
        """
        inclinations = []
        
        # Parcourir tous les fichiers .txt du dossier
        for filename in os.listdir(self.input):
            if filename.endswith('.txt'):
                # Si des fichiers de conjonction sont spécifiés, vérifier que le fichier en fait partie
                if conjunction_files is not None and filename not in conjunction_files:
                    continue
                
                file_path = os.path.join(self.input, filename)
                
                try:
                    with open(file_path, 'r') as file:
                        contenu = file.read()
                        # Chercher toutes les sections qui contiennent OBJECT2
                        sections = contenu.split('OBJECT                             =')
                        for section in sections:
                            if section.startswith('OBJECT2'):
                                # Chercher la ligne d'inclinaison dans cette section
                                match = re.search(r'COMMENT Inclination = ([\d.]+)\s+\[deg\]', section)
                                if match:
                                    inclination = float(match.group(1))
                                    inclinations.append(inclination)
                except Exception as e:
                    print(f"Erreur lors de la lecture du fichier {filename}: {str(e)}")
        
        # Trier la liste des inclinaisons par ordre croissant
        inclinations.sort()
        return inclinations

    def group_inclinations(self, inclinations, threshold=2.0):
        """
        Regroupe les inclinaisons avec une stratégie précise sans duplicatas.
        
        Args:
            inclinations (list): Liste des inclinaisons à grouper
            threshold (float): Seuil de regroupement
        
        Returns:
            list: Liste des inclinaisons représentatives après regroupement
        """
        if not inclinations:
            return []
        
        # Convertir les données pour s'assurer qu'elles sont des flottants
        inclinations = [float(inc) for inc in inclinations]
        
        # Compter les occurrences initiales
        initial_counter = Counter(inclinations)
        """ print("Distribution initiale des inclinaisons :")
        for inc, count in sorted(initial_counter.items()):
            print(f"  {inc}: {count} fois") """
        
        # Vérifier le nombre total de valeurs
        total_initial_values = sum(initial_counter.values())
        
        # Trier les inclinaisons uniques par fréquence
        sorted_unique = sorted(set(inclinations), key=lambda x: initial_counter[x], reverse=True)
        
        # Groupes finaux
        final_groups = {}
        processed_inclinationsions = set()
        
        # Première passe : grouper les inclinaisons très proches
        for current in sorted_unique:
            # Ignorer si déjà traité
            if current in processed_inclinationsions:
                continue
            
            # Créer un nouveau groupe
            current_group = {current}
            processed_inclinationsions.add(current)
            
            # Rechercher des inclinaisons similaires
            for other in sorted_unique:
                if other == current or other in processed_inclinationsions:
                    continue
                
                # Critère de proximité
                if abs(other - current) <= threshold:
                    current_group.add(other)
                    processed_inclinationsions.add(other)
            
            # Enregistrer le groupe
            group_representative = max(current_group, key=lambda x: initial_counter[x])
            final_groups[group_representative] = current_group
        
        # Préparer les résultats finaux
        grouped_inclinations = []
        sorted_groups = sorted(
            final_groups.items(), 
            key=lambda x: sum(initial_counter[inc] for inc in x[1]), 
            reverse=True
        )
        
        for representative, group in sorted_groups:
            group_count = sum(initial_counter[inc] for inc in group)
            grouped_inclinations.extend([representative] * group_count)
        
        # Vérifier que le nombre total de valeurs n'a pas changé
        total_final_values = len(grouped_inclinations)
        assert total_initial_values == total_final_values, f"Nombre de valeurs modifié : initial {total_initial_values}, final {total_final_values}"
        
        return grouped_inclinations

    def process_data(self):
        """
        Traite les données du dossier et exporte les résultats.
        """
        # Créer une instance de ConjunctionAnalyzer
        conjunction_analyzer = ConjunctionAnalyzer(self.input, self.output, self.ws, self.wb)
        
        # Extraire les fichiers de conjonction
        conjunction_analyzer.process_data()
        
        # Récupérer le premier fichier de chaque groupe de conjonction
        first_conjunction_files = set()
        for group in conjunction_analyzer.conjunctions.values():
            first_conjunction_files.add(list(group)[0])
        
        # Analyser les inclinaisons uniquement pour ces fichiers
        inclinations = self.analyze_folder(first_conjunction_files)
        
        # Grouper les inclinaisons
        grouped_inclinations = self.group_inclinations(inclinations)
        
        # Obtenir les statistiques
        stats = self.get_inclination_statistics(grouped_inclinations)
        
        # Exporter vers Excel
        self.export_to_excel(grouped_inclinations)
        
        return {
            "inclinations": grouped_inclinations,
            "statistics": stats,
            "ranges": self.get_inclination_ranges(grouped_inclinations)
        }