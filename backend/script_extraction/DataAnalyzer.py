import os
import re
import pandas as pd
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Dict, List, Set, Optional, Any, Tuple
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import Counter
import statistics

class BaseAnalyzer(ABC):
    """
    Classe abstraite mère pour tous les analyseurs de données.
    Définit la structure de base et les méthodes communes.
    """
    
    def __init__(self, input_path=None, output_path=None, worksheet=None, workbook=None):
        """
        Initialise l'analyseur avec les chemins et objets Excel.
        
        Args:
            input_path (str, optional): Chemin du dossier contenant les fichiers à analyser.
            output_path (str, optional): Chemin du fichier Excel de sortie.
            worksheet: Feuille de calcul Excel pour l'export.
            workbook: Classeur Excel pour l'export.
        """
        self.input_path = input_path
        self.output_path = output_path
        self.worksheet = worksheet
        self.workbook = workbook
    
    def set_input_path(self, input_path):
        """
        Définit le chemin du dossier d'entrée.
        
        Args:
            input_path (str): Chemin du dossier d'entrée.
        """
        self.input_path = input_path
    
    def set_output_path(self, output_path):
        """
        Définit le chemin du fichier Excel de sortie.
        
        Args:
            output_path (str): Chemin du fichier Excel de sortie.
        """
        self.output_path = output_path
    
    def set_worksheet(self, worksheet):
        """
        Définit la feuille de calcul pour l'export.
        
        Args:
            worksheet: Feuille de calcul Excel.
        """
        self.worksheet = worksheet
    
    def set_workbook(self, workbook):
        """
        Définit le classeur pour l'export.
        
        Args:
            workbook: Classeur Excel.
        """
        self.workbook = workbook
    
    @abstractmethod
    def process_data(self):
        """
        Méthode abstraite qui doit être implémentée par les classes dérivées.
        Traite les données et retourne les résultats.
        
        Returns:
            Any: Résultats de l'analyse.
        """
        pass
    
    def extract_data_from_txt(self, file_path):
        """
        Extrait les données d'un fichier texte.
        
        Args:
            file_path (str): Chemin du fichier texte.
            
        Returns:
            dict: Dictionnaire des données extraites.
        """
        data = {}
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()
                for line in lines:
                    if line.strip():
                        line = re.sub(r'\[.*?\]', '', line).strip()
                        parts = line.split('=')
                        if len(parts) == 2:
                            key = parts[0].strip()
                            value = parts[1].strip()
                            data[key] = value
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier {file_path}: {str(e)}")
        return data
    
    def count_files(self):
        """
        Compte le nombre de fichiers dans le dossier d'entrée.
        
        Returns:
            int: Nombre de fichiers.
        """
        if not self.input_path:
            return 0
        
        return len([f for f in os.listdir(self.input_path) 
                   if os.path.isfile(os.path.join(self.input_path, f))])

# -----------------------------------------------------------------------------
# 2. Classes spécialisées d'analyseurs (hiérarchie d'héritage)
# -----------------------------------------------------------------------------
class SimpleCellAnalyzer(BaseAnalyzer):
    """
    Classe abstraite pour les analyseurs qui écrivent dans des cellules individuelles.
    """
    
    @abstractmethod
    def get_cell_values(self) -> Dict[str, Any]:
        """
        Méthode abstraite qui retourne les valeurs à écrire dans les cellules.
        
        Returns:
            Dict[str, Any]: Mapping {référence_cellule: valeur}
        """
        pass
    
    def export_to_excel(self):
        """
        Exporte les valeurs dans les cellules appropriées.
        """
        if not self.worksheet:
            print("Feuille de calcul non définie.")
            return
            
        cell_values = self.get_cell_values()
        for cell_ref, value in cell_values.items():
            try:
                self.worksheet[cell_ref] = value
            except Exception as e:
                print(f"Erreur lors de l'écriture dans la cellule {cell_ref}: {str(e)}")
    
    def process_data(self):
        """
        Traite les données et exporte les résultats.
        
        Returns:
            Dict[str, Any]: Résultats de l'analyse.
        """
        results = self.get_cell_values()
        self.export_to_excel()
        return results

class RangeCellAnalyzer(BaseAnalyzer):
    """
    Classe abstraite pour les analyseurs qui écrivent dans des plages de cellules.
    """
    
    @abstractmethod
    def analyze_data(self) -> Dict[str, Any]:
        """
        Analyse les données et retourne les résultats.
        
        Returns:
            Dict[str, Any]: Résultats de l'analyse.
        """
        pass
    
    @abstractmethod
    def export_to_excel(self, data):
        """
        Exporte les données dans la feuille Excel.
        
        Args:
            data: Données à exporter.
        """
        pass
    
    def process_data(self):
        """
        Traite les données et exporte les résultats.
        
        Returns:
            Dict[str, Any]: Résultats de l'analyse.
        """
        data = self.analyze_data()
        self.export_to_excel(data)
        return data