import os
import re
import pandas as pd
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Set, Optional, Tuple

from backend.script_extraction.ScriptAnalyzerABS import BaseAnalyzer
from backend.script_extraction.Conjonction import ConjunctionAnalyzer

class SatelliteAgeAnalyzer(ConjunctionAnalyzer):
    """
    Classe pour analyser l'âge des satellites impliqués dans les conjonctions.
    """
    
    def __init__(self, input, output, ws, wb):
        super().__init__(input, output, ws, wb)
        # Dictionnaire pour stocker le nombre de satellites par catégorie d'âge
        self.age_counts = defaultdict(int)
        # Pour stocker les satellites déjà comptés (éviter les doublons)
        self.processed_satellites = set()
    
    def extract_satellite_data(self, file_path: str) -> Tuple[str, str, int]:
        """
        Extrait l'identifiant international du satellite OBJECT2 depuis un fichier CDM.
        
        Args:
            file_path (str): Chemin du fichier à analyser.
            
        Returns:
            tuple: (file_name, international_designator, age)
        """
        file_name = os.path.basename(file_path)
        international_designator = None
        age = None
        
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                
                # Rechercher l'identifiant international dans la section OBJECT2
                id_match = re.search(r'OBJECT\s*=\s*OBJECT2.*?INTERNATIONAL_DESIGNATOR\s*=\s*([^\s]+)', 
                                    content, re.DOTALL)
                
                if id_match:
                    international_designator = id_match.group(1).strip()
                    
                    # Extraire l'année depuis l'identifiant international (format YYYY-NNNX)
                    try:
                        year_str = international_designator.split('-')[0]
                        if len(year_str) == 2:  # Format ancien (YY-NNNX)
                            year = int(year_str)
                            # Ajuster pour le siècle (supposer 19xx pour les années > 50, 20xx pour <= 50)
                            if year > 50:
                                year += 1900
                            else:
                                year += 2000
                        else:  # Format moderne (YYYY-NNNX)
                            year = int(year_str)
                        
                        # Calculer l'âge en années
                        current_year = datetime.now().year
                        age = current_year - year
                    except (ValueError, IndexError):
                        print(f"Format d'identifiant international invalide: {international_designator}")
        
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier {file_path}: {str(e)}")
        
        return (file_name, international_designator, age)
    
    def classify_age(self, age):
        """
        Classifie l'âge du satellite selon les catégories définies.
        
        Args:
            age (int): Âge du satellite en années.
            
        Returns:
            str: Catégorie d'âge.
        """
        if age is None:
            return "Inconnu"
        elif age >= 6:
            return "X > 6"
        elif 5 <= age < 6:
            return "5 <= X <= 6"
        elif 4 <= age < 5:
            return "4 <= X <= 5"
        elif 3 <= age < 4:
            return "3 <= X <= 4"
        elif 2 <= age < 3:
            return "2 <= X <= 3"
        elif 1 <= age < 2:
            return "1 <= X <= 2"
        else:
            return "> 1"
    
    def analyze_satellite_ages(self):
        """
        Analyse les âges des satellites impliqués dans les conjonctions.
        
        Returns:
            dict: Nombre de satellites par catégorie d'âge.
        """
        # Réinitialiser les compteurs
        self.age_counts = defaultdict(int)
        self.processed_satellites = set()
        
        # Extraire d'abord les conjonctions si ce n'est pas déjà fait
        if not self.conjunctions:
            self.process_data()
        
        # Pour chaque groupe de conjonction, analyser le premier fichier
        for group_id, files in self.conjunctions.items():
            if not files:
                continue
                
            # Prendre le premier fichier du groupe
            first_file = list(files)[0]
            file_path = os.path.join(self.input, first_file)
            
            # Extraire les données du satellite
            _, international_designator, age = self.extract_satellite_data(file_path)
            
            # Si nous avons un identifiant international valide et que nous ne l'avons pas déjà traité
            if international_designator and international_designator not in self.processed_satellites:
                age_category = self.classify_age(age)
                self.age_counts[age_category] += 1
                self.processed_satellites.add(international_designator)
        
        # Préparer le résultat final avec toutes les catégories possibles (même vides)
        all_categories = ["> 1", "1 <= X <= 2", "2 <= X <= 3", "3 <= X <= 4", 
                          "4 <= X <= 5", "5 <= X <= 6", "X > 6", "Inconnu"]
        
        final_counts = {}
        for category in all_categories:
            final_counts[category] = self.age_counts.get(category, 0)
        
        return final_counts
    
    def export_age_to_excel(self):
        """
        Exporte les résultats d'analyse d'âge dans la colonne AI de la feuille Excel.
        
        Returns:
            bool: True si l'export a réussi, False sinon.
        """
        if not self.ws:
            print("Feuille Excel non définie")
            return False
        
        # Obtenir les résultats d'âge
        age_counts = self.analyze_satellite_ages()
        
        # Écrire les résultats dans la colonne AI (colonne 35)
  
        row = 3 
        
        for category, count in age_counts.items():
            self.ws.cell(row=row, column=35).value = count
            row += 1
        
        return True
    
    def process_data(self):
        """
        Surcharge de la méthode process_data pour inclure l'analyse d'âge.
        """
        # Appeler la méthode de la classe parente pour extraire les conjonctions
        super().process_data()
        
        # Analyser les âges des satellites
        self.analyze_satellite_ages()
        
        # Exporter les résultats d'âge
        self.export_age_to_excel()