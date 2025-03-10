from ttkthemes import ThemedTk
from gui.interface import SatelliteAnalysisGUI


def main():
    root = ThemedTk(theme="arc")  # Utilisation d'un thème moderne
    app = SatelliteAnalysisGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()