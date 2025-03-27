import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkthemes import ThemedTk
import os
import json

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import backend.script_execl.Execl as Execl
import backend.script_execl.Execl_Brut as Execl_Brut

from backend.script_extraction import (
    Probabilite, Object_type, Country
)

import backend.classement.script_classement as script_classement
import backend.script_graphique.generateur_graphique as generateur_graphique

class ModernButton(ttk.Button):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        
    def on_enter(self, event):
        self.state(['active'])

    def on_leave(self, event):
        self.state(['!active'])

class SatelliteAnalysisGUI:
    VERSION = None
    
    def __init__(self, root, legal_notice=None, version=None):
        self.root = root
        self.root.geometry("1024x768")
        self.legal_notice = legal_notice
        self.VERSION = version
        self.root.title(f"Star Guardian v{self.VERSION}")  # Titre de la fenêtr
        
        if getattr(sys, 'frozen', False):
            # Si c'est un exécutable PyInstaller, utilisez sys._MEIPASS pour obtenir le chemin d'accès aux fichiers extraits
            icon_path = os.path.join(sys._MEIPASS, 'gui', 'star_guardian.ico')
        else:
            # Si c'est un script Python normal, utilisez un chemin relatif
            icon_path = os.path.join(os.path.dirname(__file__), "star_guardian.ico")
            
        self.root.iconbitmap(icon_path)

        # Configuration du style
        self.setup_styles()
        
        # Variables
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.current_frame = None
        
        self.model_format = tk.StringVar(value="excel")
        
        # Initialisation du générateur de graphiques
        self.graphique = generateur_graphique.GenerateurGraphique()
        
        self.execl = Execl.SatelliteDataProcessor()
        self.ConjonctionAnalyzer = self.execl.getConjonctionAnalyzer()
        
        # Variable 
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.current_frame = None
        
        # Nouvelles variables pour le tableau de bord
        self.files_count = tk.StringVar(value="0")
        self.start_date = tk.StringVar(value="--/--/----")
        self.end_date = tk.StringVar(value="--/--/----")
        self.satellite_name = tk.StringVar(value="Non défini")
        self.conjunction_count = tk.StringVar(value="0")
        
        # Variables pour les paramètres
        self.default_input_dir = tk.StringVar()
        self.default_output_dir = tk.StringVar()
        self.excel_template_path = tk.StringVar()
        self.auto_update_dashboard = tk.BooleanVar(value=True)
        self.theme_selection = tk.StringVar(value="arc")
        
        # Charger les paramètres
        self.config_file = os.path.join(os.path.dirname(__file__), "config.json")
        self.load_settings()

        # Création de l'interface principale
        self.create_main_interface()
        
    def setup_styles(self):
        style = ttk.Style()
        style.configure(
            'Sidebar.TFrame',
            background='#2c3e50'
        )
        style.configure(
            'Content.TFrame',
            background='#f5f6fa'
        )
        style.configure(
            'Title.TLabel',
            font=('Helvetica', 24, 'bold'),
            foreground='#2c3e50',
            background='#f5f6fa'
        )
        style.configure(
            'Subtitle.TLabel',
            font=('Helvetica', 12),
            foreground='#7f8c8d',
            background='#f5f6fa'
        )
        style.configure(
            'Card.TFrame',
            background='white',
            relief='solid',
            borderwidth=1
        )
        style.configure(
            'Settings.TLabelframe',
            background='white',
            font=('Helvetica', 12, 'bold')
        )
        style.configure(
            'Settings.TLabelframe.Label',
            font=('Helvetica', 12, 'bold'),
            foreground='#2c3e50',
            background='white'
        )
        
    def create_main_interface(self):
        # Container principal
        self.main_container = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Sidebar
        self.sidebar = ttk.Frame(self.main_container, style='Sidebar.TFrame')
        self.main_container.add(self.sidebar, weight=1)
        
        # Content
        self.content = ttk.Frame(self.main_container, style='Content.TFrame')
        self.main_container.add(self.content, weight=4)
        
        self.create_sidebar()
        self.show_dashboard()
        
    def create_sidebar(self):
        # Logo ou titre
        title_frame = ttk.Frame(self.sidebar)
        title_frame.pack(fill=tk.X, pady=20, padx=10)

        label = ttk.Label(
            title_frame,
            text="Star Guardian",
            font=('Helvetica', 14, 'bold'),
            foreground='white',
            background='#2c3e50',
            anchor="center"
        )
        label.pack(fill=tk.X, expand=True)  # Ajuste le fond au texte

        # Menu buttons
        menu_items = [
            ("Tableau de bord", self.show_dashboard),
            ("Trier les fichiers", self.show_file_sorting),
            # ("Graphiques", self.show_graphics_page),
            ("Export Excel", self.show_excel_page),
            ("Paramètres", self.show_settings_page),
        ]
        
        for text, command in menu_items:
            btn = ModernButton(
                self.sidebar,
                text=text,
                command=command,
                width=20
            )
            btn.pack(pady=5, padx=10)
            
        about_frame = ttk.Frame(self.sidebar, style='Sidebar.TFrame')
        about_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        
        ModernButton(
            about_frame,
            text="À propos",
            command=self.show_about,
            width=20
        ).pack(side=tk.LEFT)

        # Version label
        version_frame = ttk.Frame(self.sidebar, style='Sidebar.TFrame')
        version_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        
        version_label = ttk.Label(
            version_frame,
            text=f"v{self.VERSION}",
            foreground='white',
            background='#2c3e50',
            font=('Helvetica', 8)
        )
        version_label.pack(side=tk.LEFT)
            
    def clear_content(self):
        for widget in self.content.winfo_children():
            widget.destroy()
            
    def show_about(self):
        """Affiche les mentions légales dans une fenêtre modale"""
        about_window = tk.Toplevel(self.root)
        about_window.title("À propos de Star Guardian")
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        about_window.transient(self.root)
        about_window.grab_set()

        # Centrer la fenêtre
        about_window.geometry("+%d+%d" % (
            self.root.winfo_rootx() + self.root.winfo_width()/2 - 200,
            self.root.winfo_rooty() + self.root.winfo_height()/2 - 150
        ))

        text = tk.Text(about_window, wrap=tk.WORD, padx=20, pady=20)
        text.pack(fill=tk.BOTH, expand=True)
        text.insert("1.0", self.legal_notice)
        text.config(state=tk.DISABLED)

        ModernButton(
            about_window,
            text="Fermer",
            command=about_window.destroy
        ).pack(pady=10)        
    
    def show_dashboard(self):
        self.clear_content()
        
        # Titre
        title_frame = ttk.Frame(self.content, style='Content.TFrame')
        title_frame.pack(fill=tk.X, padx=20, pady=20)
        ttk.Label(
            title_frame,
            text="Tableau de bord",
            style='Title.TLabel'
        ).pack(anchor='w')
        ttk.Label(
            title_frame,
            text="Vue d'ensemble de l'analyse des données",
            style='Subtitle.TLabel'
        ).pack(anchor='w')
        
        # Cards container
        cards_frame = ttk.Frame(self.content, style='Content.TFrame')
        cards_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        # Satellite Info Card
        satellite_card = ttk.Frame(cards_frame, style='Card.TFrame')
        satellite_card.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')
        
        ttk.Label(
            satellite_card,
            text="Satellite",
            font=('Helvetica', 14, 'bold')
        ).pack(pady=(15,5), padx=15)
        
        ttk.Label(
            satellite_card,
            textvariable=self.satellite_name,
            font=('Helvetica', 18)
        ).pack(pady=(0,15), padx=15)
        
        # Statistics cards
        self.create_stat_card(cards_frame, "Fichiers traités", self.files_count, 1, 0)
        self.create_stat_card(cards_frame, "Nombre de conjonctions", self.conjunction_count, 1, 1)
        
        # Dates Card
        dates_card = ttk.Frame(cards_frame, style='Card.TFrame')
        dates_card.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')
        
        ttk.Label(
            dates_card,
            text="Période d'analyse",
            font=('Helvetica', 14, 'bold')
        ).pack(pady=(15,5), padx=15)
        
        dates_frame = ttk.Frame(dates_card)
        dates_frame.pack(pady=(0,15), padx=15)
        
        ttk.Label(
            dates_frame,
            text="Du: ",
            font=('Helvetica', 12)
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            dates_frame,
            textvariable=self.start_date,
            font=('Helvetica', 12)
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            dates_frame,
            text=" au: ",
            font=('Helvetica', 12)
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            dates_frame,
            textvariable=self.end_date,
            font=('Helvetica', 12)
        ).pack(side=tk.LEFT)
        
        # Configure grid weights
        cards_frame.grid_columnconfigure(0, weight=1)
        cards_frame.grid_columnconfigure(1, weight=1)
        for i in range(3):
            cards_frame.grid_rowconfigure(i, weight=1)
        
        # Si le paramètre d'auto-update est activé, mettre à jour le tableau de bord
        """ if self.auto_update_dashboard.get():
            self.update_dashboard_info() """
        
    def update_dashboard_info(self):
        try:
            satellite = self.execl.nom_satellite()
            self.satellite_name.set(satellite if satellite else "Non défini")
            
            files = self.execl.compter_fichiers()
            self.files_count.set(str(files) if files is not None else "0")
            
            # Vérifiez si date_analyzer existe avant d'appeler ses méthodes
            if hasattr(self.execl, 'date_analyzer') and self.execl.date_analyzer is not None:
                all_dates = self.execl.date_analyzer._collect_dates()  # Collecter toutes les dates

                # Trouver les dates minimales et maximales
                start_date = self.execl.date_analyzer.find_min_date(all_dates)  # Passer la liste de dates à la méthode
                end_date = self.execl.date_analyzer.find_max_date(all_dates)  # Passer la liste de dates à la méthode
                
                self.start_date.set(start_date if start_date else "--/--/----")
                self.end_date.set(end_date if end_date else "--/--/----")
            else:
                self.start_date.set("--/--/----")
                self.end_date.set("--/--/----")
            
            self.ConjonctionAnalyzer  = self.execl.getConjonctionAnalyzer()
            
            # Vérifiez si ConjonctionAnalyzer existe avant d'appeler ses méthodes
            if self.ConjonctionAnalyzer  is not None:
                self.ConjonctionAnalyzer.process_data()
                self.conjunction_count.set(str(self.ConjonctionAnalyzer.get_conjunction_count()))
            else:
                self.conjunction_count.set("0")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue lors de la mise à jour du tableau de bord: {str(e)}")
    
    def create_stat_card(self, parent, title, value_var, row, col):
        card = ttk.Frame(parent, style='Card.TFrame')
        card.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')
        
        ttk.Label(
            card,
            text=title,
            font=('Helvetica', 12)
        ).pack(pady=(15,5), padx=15)
        
        ttk.Label(
            card,
            textvariable=value_var,
            font=('Helvetica', 24, 'bold')
        ).pack(pady=(0,15), padx=15)
        
        parent.grid_columnconfigure(col, weight=1)
        parent.grid_rowconfigure(row, weight=1)
        
    def show_graphics_page(self):
        self.clear_content()
        
        title_frame = ttk.Frame(self.content, style='Content.TFrame')
        title_frame.pack(fill=tk.X, padx=20, pady=20)
        ttk.Label(
            title_frame,
            text="Génération de graphiques",
            style='Title.TLabel'
        ).pack(anchor='w')
        
        options_frame = ttk.Frame(self.content, style='Card.TFrame')
        options_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        graph_types = [
            ("Tous les graphiques", self.generate_all_graphics),
            ("Distribution par pays", lambda: self.generate_specific_graph("country")),
            ("Types d'objets", lambda: self.generate_specific_graph("type")),
            ("Probabilités de collision", lambda: self.generate_specific_graph("probability")),
            ("Distances minimales", lambda: self.generate_specific_graph("distance")),
            ("Manœuvrabilité", lambda: self.generate_specific_graph("maneuverable")),
            ("Inclinaison", lambda: self.generate_specific_graph("inclination"))
        ]
        
        for text, command in graph_types:
            ModernButton(
                options_frame,
                text=text,
                command=command
            ).pack(pady=10, padx=20, fill=tk.X)
            
    def show_file_sorting(self):
        self.clear_content()
        
        title_frame = ttk.Frame(self.content, style='Content.TFrame')
        title_frame.pack(fill=tk.X, padx=20, pady=20)
        ttk.Label(
            title_frame,
            text="Tri des fichiers",
            style='Title.TLabel'
        ).pack(anchor='w')
        
        form_frame = ttk.Frame(self.content, style='Card.TFrame')
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Source directory
        source_frame = ttk.Frame(form_frame)
        source_frame.pack(fill=tk.X, padx=20, pady=10)
        ttk.Label(source_frame, text="Répertoire source:").pack(anchor='w')
        source_entry = ttk.Entry(source_frame, textvariable=self.input_dir)
        source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)
        
        # Utiliser le répertoire par défaut s'il est défini
        if self.default_input_dir.get():
            self.input_dir.set(self.default_input_dir.get())
            
        ModernButton(
            source_frame,
            text="Parcourir",
            command=lambda: self.input_dir.set(filedialog.askdirectory())
        ).pack(side=tk.RIGHT, padx=5)
        
        # Target directory
        target_frame = ttk.Frame(form_frame)
        target_frame.pack(fill=tk.X, padx=20, pady=10)
        ttk.Label(target_frame, text="Répertoire cible:").pack(anchor='w')
        target_entry = ttk.Entry(target_frame, textvariable=self.output_dir)
        target_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)
        
        # Utiliser le répertoire par défaut s'il est défini
        if self.default_output_dir.get():
            self.output_dir.set(self.default_output_dir.get())
            
        ModernButton(
            target_frame,
            text="Parcourir",
            command=lambda: self.output_dir.set(filedialog.askdirectory())
        ).pack(side=tk.RIGHT, padx=5)
        
        # Sort button
        ModernButton(
            form_frame,
            text="Trier les fichiers",
            command=self.sort_files
        ).pack(pady=20)
        
    def sort_files(self):
        if not self.input_dir.get() or not self.output_dir.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner les répertoires source et cible")
            return
            
        try:
            files_count = len([f for f in os.listdir(self.input_dir.get()) if os.path.isfile(os.path.join(self.input_dir.get(), f))])
            
            script_classement.organize_files_by_satellite(
                self.input_dir.get(),
                self.output_dir.get()
            )

            messagebox.showinfo("Succès", "Les fichiers ont été triés avec succès")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue: {str(e)}")
        
    def show_excel_page(self):
        self.clear_content()
        
        title_frame = ttk.Frame(self.content, style='Content.TFrame')
        title_frame.pack(fill=tk.X, padx=20, pady=20)
        ttk.Label(
            title_frame,
            text="Export Excel",
            style='Title.TLabel'
        ).pack(anchor='w')
        
        form_frame = ttk.Frame(self.content, style='Card.TFrame')
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Source directory selection
        source_frame = ttk.Frame(form_frame)
        source_frame.pack(fill=tk.X, padx=20, pady=10)
        ttk.Label(source_frame, text="Répertoire source:").pack(anchor='w')
        source_entry = ttk.Entry(source_frame, textvariable=self.input_dir)
        source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)
        
        # Utiliser le répertoire par défaut s'il est défini
        if self.default_input_dir.get():
            self.input_dir.set(self.default_input_dir.get())
            
        ModernButton(
            source_frame,
            text="Parcourir",
            command=lambda: self.input_dir.set(filedialog.askdirectory())
        ).pack(side=tk.RIGHT, padx=5)
        
        ModernButton(
            form_frame,
            text="Générer le fichier Excel",
            command=self.generate_excel
        ).pack(pady=20)
              
    def generate_excel(self):
        input_directory = self.input_dir.get()
        if not input_directory:
            input_directory = filedialog.askdirectory(title="Sélectionnez le répertoire source")
            if not input_directory:
                return
                
        try:
            # Utiliser le modèle défini dans les paramètres s'il existe
            if self.excel_template_path.get():
                model_calc = self.excel_template_path.get()
            else:
                current_dir = os.path.dirname(os.path.realpath(__file__))
                # Construire le chemin relatif vers le modèle approprié selon le format
                if self.model_format.get() == "excel":
                    model_file = 'config_excel.xlsx'
                elif self.model_format.get() == "calc":
                    model_file = 'config_excel.xlsx'  # Toujours utiliser Excel comme modèle intermédiaire
                else:
                    model_file = 'config_excel.xlsx'
                    
                model_calc = os.path.join(current_dir, '..', 'config', model_file)
                # Convertir ce chemin en chemin absolu
                model_calc = os.path.abspath(model_calc)
            
            # Vérifier que le fichier modèle existe
            if not os.path.exists(model_calc):
                messagebox.showerror("Erreur", f"Le fichier modèle n'existe pas: {model_calc}")
                return
            
            # Configurer le processeur
            self.execl.set_dossier(input_directory)
            self.execl.setCheminModel(model_calc)
            
            # Obtenir le nom du satellite pour le fichier de sortie
            nom_satellite = self.execl.nom_satellite() or "Satellite_Anonyme"
            nom_fichier = re.sub(r'[^\w\s-]', '', nom_satellite)
            
            # Créer un chemin de sortie temporaire pour Excel (toujours travailler avec Excel en interne)
            output_dir = self.default_output_dir.get() or os.path.dirname(os.path.realpath(__file__))  # Changer ici
            temp_excel_output = os.path.join(output_dir, f"{nom_fichier}_Analyse.xlsx")
            
            # Définir le chemin de sortie temporaire pour le traitement interne
            self.execl.set_chemin_sortie(temp_excel_output)
            
            # Définir le format à utiliser pour la conversion finale
            format_final = self.model_format.get()
            self.execl.set_format(format_final)
            
            # Exécuter l'analyse (ceci produira un fichier Excel)
            success = self.execl.copier_modele_excel(model_calc)
            if not success:
                messagebox.showerror("Erreur", "Impossible de copier le modèle Excel.")
                return
                
            success = self.execl.executer_analyse()
            
            if not success:
                messagebox.showerror("Erreur", "L'analyse a échoué.")
                return
            
            # Si le format n'est pas Excel, convertir le fichier
            final_output = temp_excel_output
            if format_final != "excel":
                # Construire le chemin pour le fichier de sortie final
                if format_final == "calc":
                    final_output = os.path.join(output_dir, f"{nom_fichier}_Analyse.ods")
                
                # Convertir le fichier Excel en format final
                conversion_success = self.execl.convert_to_format(temp_excel_output, final_output, format_final)
                if conversion_success:
                    # Supprimer le fichier Excel temporaire si la conversion a réussi
                    try:
                        os.remove(temp_excel_output)
                        # Mettre à jour le chemin de sortie
                        self.execl.chemin_sortie = final_output
                    except Exception as e:
                        print(f"Avertissement: Impossible de supprimer le fichier temporaire: {e}")
                else:
                    messagebox.showwarning("Attention", 
                        f"La conversion vers {format_final} a échoué. Le fichier Excel a été conservé.")
                    final_output = temp_excel_output

            if self.auto_update_dashboard.get():
                self.update_dashboard_info()
            
            messagebox.showinfo("Succès", f"Le fichier a été généré")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("Erreur", f"Une erreur est survenue: {str(e)}\n\nDétails: {error_details}")


    def generate_all_graphics(self):
        directory = self.input_dir.get()
        if not directory:
            directory = filedialog.askdirectory()
            if not directory:
                return

        try:
            self.graphique.generer_graphique_type(
                Object_type.analyze_folder(directory),
                f"{self.graphique.recuperer_titre_graphique_fichier(directory)} - Object type",
                False
            )
            self.graphique.generer_graphique_avec_legend(
                Probabilite.analyze_folder(directory),
                f"{self.graphique.recuperer_titre_graphique_fichier(directory)} - Collision_Prob",
                ["≥1E-4", "1E-4>X≥1E-5", "1E-5>X≥1E-6", "1E-6>X≥1E-7", "Entre 1E-7 et 1E-8", "≤1E-8"],
                False
            )

            messagebox.showinfo("Succès", "Tous les graphiques ont été générés")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue: {str(e)}")
            
    def generate_specific_graph(self, graph_type):
        directory = self.input_dir.get()
        if not directory:
            directory = filedialog.askdirectory()
            if not directory:
                return
            
        try:
            if graph_type == "country":
                self.graphique.generer_graphique_avec_legend(
                    Country.analyze_folder(directory),
                    f"{self.graphique.recuperer_titre_graphique_fichier(directory)} - Country",
                    Country.get_unique_countries(Country.analyze_folder(directory)),
                    True
                )
            elif graph_type == "type":
                self.graphique.generer_graphique_type(
                    Object_type.analyze_folder(directory),
                    f"{self.graphique.recuperer_titre_graphique_fichier(directory)} - Object type",
                    True
                )
            # TODO: Ajouter les autres types de graphiques (probabilité, distance, etc.)
            
            messagebox.showinfo("Succès", "Le graphique a été généré")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur est survenue: {str(e)}")
            
    def show_settings_page(self):
        """Affiche la page des paramètres"""
        self.clear_content()
        
        # Titre
        title_frame = ttk.Frame(self.content, style='Content.TFrame')
        title_frame.pack(fill=tk.X, padx=20, pady=20)
        ttk.Label(
            title_frame,
            text="Paramètres",
            style='Title.TLabel'
        ).pack(anchor='w')
        ttk.Label(
            title_frame,
            text="Configuration de l'application",
            style='Subtitle.TLabel'
        ).pack(anchor='w')
        
        # Création du cadre pour les paramètres
        settings_frame = ttk.Frame(self.content, style='Card.TFrame')
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Section des chemins par défaut
        paths_frame = ttk.LabelFrame(settings_frame, text="Chemins par défaut", style='Settings.TLabelframe')
        paths_frame.pack(fill=tk.X, expand=False, padx=20, pady=10)
        
        # Répertoire d'entrée par défaut
        input_dir_frame = ttk.Frame(paths_frame)
        input_dir_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        ttk.Label(input_dir_frame, text="Répertoire source par défaut:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(input_dir_frame, textvariable=self.default_input_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ModernButton(
            input_dir_frame,
            text="Parcourir",
            command=lambda: self.default_input_dir.set(filedialog.askdirectory())
        ).pack(side=tk.RIGHT, padx=5)
        
        # Répertoire de sortie par défaut
        output_dir_frame = ttk.Frame(paths_frame)
        output_dir_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        ttk.Label(output_dir_frame, text="Répertoire cible par défaut:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(output_dir_frame, textvariable=self.default_output_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ModernButton(
            output_dir_frame,
            text="Parcourir",
            command=lambda: self.default_output_dir.set(filedialog.askdirectory())
        ).pack(side=tk.RIGHT, padx=5)
        
        # Modèle Excel
        excel_template_frame = ttk.Frame(paths_frame)
        excel_template_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        ttk.Label(excel_template_frame, text="Modèle Excel:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(excel_template_frame, textvariable=self.excel_template_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ModernButton(
            excel_template_frame,
            text="Parcourir",
            command=lambda: self.excel_template_path.set(filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx")]))
        ).pack(side=tk.RIGHT, padx=5)
        
        # Format du modèle
        model_format_frame = ttk.Frame(paths_frame)
        model_format_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        ttk.Label(model_format_frame, text="Format du modèle:").pack(side=tk.LEFT, padx=5)
        
        # Options de format
        format_options = [
            ("Microsoft Excel (*.xlsx)", "excel"),
            # ("LibreOffice Calc (*.ods)", "calc"),
        ]
        
        config_files_frame = ttk.LabelFrame(settings_frame, text="Fichiers de configuration", style='Settings.TLabelframe')
        config_files_frame.pack(fill=tk.X, expand=False, padx=20, pady=10)
        
        # Configuration des pays
        country_config_frame = ttk.Frame(config_files_frame)
        country_config_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        ttk.Label(country_config_frame, text="Fichier de configuration des pays:").pack(side=tk.LEFT, padx=5)
        
        def show_country_config():
            try:
                # Chemin vers le fichier de configuration des pays
                config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config', 'Country_2025-01-24.xlsx'))
                
                if os.path.exists(config_path):
                    # Ouvre l'explorateur Windows en sélectionnant le fichier
                    os.system(f'explorer /select,"{config_path}"')
                else:
                    messagebox.showwarning("Attention", "Le fichier de configuration des pays (Country_2025-01-24.xlsx) n'existe pas.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'accéder au fichier: {str(e)}")
        
        ModernButton(
            country_config_frame,
            text="Localiser le fichier",
            command=show_country_config
        ).pack(side=tk.LEFT, padx=5)
        
        format_subframe = ttk.Frame(model_format_frame)
        format_subframe.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        for text, value in format_options:
            ttk.Radiobutton(
                format_subframe,
                text=text,
                variable=self.model_format,
                value=value
            ).pack(anchor=tk.W, padx=5, pady=2)
        
        
        # Section des préférences interface
        ui_frame = ttk.LabelFrame(settings_frame, text="Interface utilisateur", style='Settings.TLabelframe')
        ui_frame.pack(fill=tk.X, expand=False, padx=20, pady=10)
        
        # Mise à jour automatique du tableau de bord
        auto_update_frame = ttk.Frame(ui_frame)
        auto_update_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        ttk.Checkbutton(
            auto_update_frame, 
            text="Mettre à jour automatiquement le tableau de bord après l'analyse", 
            variable=self.auto_update_dashboard
        ).pack(side=tk.LEFT, padx=5)
        
        # Sélection du thème
        theme_frame = ttk.Frame(ui_frame)
        theme_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        ttk.Label(theme_frame, text="Thème:").pack(side=tk.LEFT, padx=5)
        
        # Liste des thèmes disponibles (à adapter selon les thèmes de ttkthemes)
        available_themes = ["arc", "aquativo", "black", "blue", "clearlooks", "elegance", "kroc", "radiance", "vista", "yaru"]
        theme_combo = ttk.Combobox(
            theme_frame, 
            textvariable=self.theme_selection,
            values=available_themes,
            state="readonly"
        )
        theme_combo.pack(side=tk.LEFT, padx=5)
        
        # Bouton pour appliquer le thème
        ModernButton(
            theme_frame,
            text="Appliquer",
            command=self.apply_theme
        ).pack(side=tk.LEFT, padx=5)
        
        # Boutons de sauvegarde/annulation
        buttons_frame = ttk.Frame(settings_frame)
        buttons_frame.pack(fill=tk.X, expand=False, padx=20, pady=20)
        
        ModernButton(
            buttons_frame,
            text="Enregistrer les paramètres",
            command=self.save_settings
        ).pack(side=tk.RIGHT, padx=5)
        
        ModernButton(
            buttons_frame,
            text="Restaurer les valeurs par défaut",
            command=self.reset_settings
        ).pack(side=tk.LEFT, padx=5)
    
    def generate_excel(self):
        input_directory = self.input_dir.get()
        if not input_directory:
            input_directory = filedialog.askdirectory(title="Sélectionnez le répertoire source")
            if not input_directory:
                return
                
        try:
            # Utiliser le modèle défini dans les paramètres s'il existe
            if self.excel_template_path.get():
                model_calc = self.excel_template_path.get()
            else:
                current_dir = os.path.dirname(os.path.realpath(__file__))
                # Construire le chemin relatif vers le modèle approprié selon le format
                if self.model_format.get() == "excel":
                    model_file = 'config_excel.xlsx'
                elif self.model_format.get() == "calc":
                    model_file = 'config_excel.xlsx'  # Toujours utiliser Excel comme modèle intermédiaire
                else:
                    model_file = 'config_excel.xlsx'
                    
                model_calc = os.path.join(current_dir, '..', 'config', model_file)
                # Convertir ce chemin en chemin absolu
                model_calc = os.path.abspath(model_calc)
            
            # Vérifier que le fichier modèle existe
            if not os.path.exists(model_calc):
                messagebox.showerror("Erreur", f"Le fichier modèle n'existe pas: {model_calc}")
                return
            
            # Configurer le processeur
            self.execl.set_dossier(input_directory)
            self.execl.setCheminModel(model_calc)
            
            # Obtenir le nom du satellite pour le fichier de sortie
            nom_satellite = self.execl.nom_satellite() or "Satellite_Anonyme"
            nom_fichier = re.sub(r'[^\w\s-]', '', nom_satellite)
            
            # Générer un nom versionné pour le fichier
            output_dir = self.default_output_dir.get() or os.path.dirname(model_calc)
            nom_fichier_versionne = self.execl.generer_nom_fichier_versionne(output_dir, nom_fichier)
            
            # Déterminer l'extension selon le format choisi
            format_final = self.model_format.get()
            extension = ".xlsx" if format_final == "excel" else ".ods"
            
            # Demander à l'utilisateur où enregistrer le fichier avec le nom versionné comme suggestion
            output_file = filedialog.asksaveasfilename(
                title="Enregistrer le fichier sous",
                initialdir=output_dir,
                initialfile=f"{nom_fichier_versionne}{extension}",
                filetypes=[("Excel files", "*.xlsx"), ("LibreOffice Calc files", "*.ods")]
            )
            
            if not output_file:
                return  # L'utilisateur a annulé
                
            # Extraire le répertoire et le nom du fichier
            output_dir = os.path.dirname(output_file)
            
            # Créer un chemin de sortie temporaire pour Excel (toujours travailler avec Excel en interne)
            temp_excel_output = os.path.join(output_dir, f"{os.path.basename(output_file).split('.')[0]}_temp.xlsx")
            
            # Définir le chemin de sortie temporaire pour le traitement interne
            self.execl.set_chemin_sortie(temp_excel_output)
            
            # Définir le format à utiliser pour la conversion finale
            self.execl.set_format(format_final)   
                
            success = self.execl.executer_analyse()
            
            if not success:
                messagebox.showerror("Erreur", "L'analyse a échoué.")
                return
            
            # Si le format n'est pas Excel, convertir le fichier
            if format_final != "excel":
                # Convertir le fichier Excel en format final
                conversion_success = self.execl.convert_to_format(temp_excel_output, output_file, format_final)
                if conversion_success:
                    # Supprimer le fichier Excel temporaire si la conversion a réussi
                    try:
                        os.remove(temp_excel_output)
                        # Mettre à jour le chemin de sortie
                        self.execl.chemin_sortie = output_file
                    except Exception as e:
                        print(f"Avertissement: Impossible de supprimer le fichier temporaire: {e}")
                else:
                    messagebox.showwarning("Attention", 
                        f"La conversion vers {format_final} a échoué. Le fichier Excel a été conservé.")
                    # Renommer le fichier temporaire vers le nom choisi par l'utilisateur
                    try:
                        os.rename(temp_excel_output, output_file.replace(extension, ".xlsx"))
                        self.execl.chemin_sortie = output_file.replace(extension, ".xlsx")
                    except Exception as e:
                        print(f"Erreur lors du renommage du fichier: {e}")
                        self.execl.chemin_sortie = temp_excel_output
            else:
                # Pour Excel, renommer simplement le fichier temporaire vers le nom choisi par l'utilisateur
                try:
                    os.rename(temp_excel_output, output_file)
                    self.execl.chemin_sortie = output_file
                except Exception as e:
                    print(f"Erreur lors du renommage du fichier: {e}")
                    self.execl.chemin_sortie = temp_excel_output

            if self.auto_update_dashboard.get():
                self.update_dashboard_info()
            
            messagebox.showinfo("Succès", f"Le fichier a été généré: {self.execl.chemin_sortie}")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("Erreur", f"Une erreur est survenue: {str(e)}\n\nDétails: {error_details}")
    
    def apply_theme(self):
        """Applique le thème sélectionné"""
        try:
            self.root.set_theme(self.theme_selection.get())
            # Forcer le rafraîchissement de l'interface
            self.clear_content()
            self.show_settings_page()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'appliquer le thème: {str(e)}")
    
    def save_settings(self):
        """Sauvegarde les paramètres dans un fichier JSON"""
        settings = {
            "default_input_dir": self.default_input_dir.get(),
            "default_output_dir": self.default_output_dir.get(),
            "excel_template_path": self.excel_template_path.get(),
            "auto_update_dashboard": self.auto_update_dashboard.get(),
            "theme": self.theme_selection.get(),
            "model_format": self.model_format.get()  # Ajout du format du modèle
        }
        
        try:
            with open(self.config_file, 'w') as f:
                json.dump(settings, f, indent=4)
            messagebox.showinfo("Succès", "Les paramètres ont été enregistrés avec succès")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de sauvegarder les paramètres: {str(e)}")

    def load_settings(self):
        """Charge les paramètres depuis un fichier JSON"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    settings = json.load(f)
                
                self.default_input_dir.set(settings.get("default_input_dir", ""))
                self.default_output_dir.set(settings.get("default_output_dir", ""))
                self.excel_template_path.set(settings.get("excel_template_path", ""))
                self.auto_update_dashboard.set(settings.get("auto_update_dashboard", True))
                self.model_format.set(settings.get("model_format", "excel"))  # Chargement du format
                
                theme = settings.get("theme", "arc")
                self.theme_selection.set(theme)
                # Appliquer le thème chargé
                try:
                    self.root.set_theme(theme)
                except Exception:
                    # Si le thème n'existe pas, utiliser le thème par défaut
                    self.root.set_theme("arc")
                    self.theme_selection.set("arc")
        except Exception as e:
            # En cas d'erreur, utiliser les valeurs par défaut
            self.default_input_dir.set("")
            self.default_output_dir.set("")
            self.excel_template_path.set("")
            self.auto_update_dashboard.set(True)
            self.model_format.set("excel")  # Format par défaut
            self.theme_selection.set("arc")
            print(f"Erreur lors du chargement des paramètres: {str(e)}")
    
    def reset_settings(self):
        """Réinitialise les paramètres aux valeurs par défaut"""
        self.default_input_dir.set("")
        self.default_output_dir.set("")
        self.excel_template_path.set("")
        self.auto_update_dashboard.set(True)
        self.model_format.set("excel")  # Réinitialisation du format par défaut
        self.theme_selection.set("arc")
        
        # Appliquer le thème par défaut
        self.root.set_theme("arc")
        
        # Mettre à jour l'interface
        self.clear_content()
        self.show_settings_page()
        
        messagebox.showinfo("Paramètres réinitialisés", "Les paramètres ont été réinitialisés aux valeurs par défaut")