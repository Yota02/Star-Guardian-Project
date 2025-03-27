"""
Star Guardian Project
Copyright (c) 2024 
Tous droits réservés.

Ce logiciel est protégé par les lois sur le droit d'auteur et les traités internationaux.
Toute reproduction ou distribution non autorisée de ce programme, ou d'une partie de celui-ci,
peut entraîner des sanctions civiles et pénales sévères et fera l'objet de poursuites dans 
toute la mesure permise par la loi.
"""

__version__ = "1.0.1"
__author__ = "Alexis Michaux-Kinet"
__copyright__ = "Copyright (c) 2025"
__license__ = "Propriétaire"

from ttkthemes import ThemedTk
from gui.interface import SatelliteAnalysisGUI

LEGAL_NOTICE = """
Star Guardian Project v{version}
{copyright}

Ce logiciel est la propriété exclusive de son auteur.
Développé par : {author}
""".format(version=__version__, copyright=__copyright__, author=__author__)

def main():
    root = ThemedTk(theme="arc")
    app = SatelliteAnalysisGUI(root, legal_notice=LEGAL_NOTICE, version=__version__)
    root.mainloop()

if __name__ == "__main__":
    main()