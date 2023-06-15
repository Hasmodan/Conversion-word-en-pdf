import os
import datetime
import win32com.client
from tqdm import tqdm
import pdfplumber
from menu import monemnu
import time
monemnu()
time.sleep(3)

# Obtenir le chemin absolu du répertoire actuel
repertoire = os.getcwd()

# Créer le dossier "fichiers convertis" s'il n'existe pas déjà
dossier_convertis = os.path.join(repertoire, "fichiers convertis")
if not os.path.exists(dossier_convertis):
    os.mkdir(dossier_convertis)

# Récupérer la date du jour au format souhaité
date_du_jour = datetime.datetime.now().strftime("%Y-%m-%d")

# Récupérer la liste des fichiers Word dans le répertoire actuel
fichiers_word = [fichier for fichier in os.listdir(repertoire) if fichier.endswith(".docx")]

# Compteur de fichiers convertis
nb_fichiers_convertis = 0

# Conversion des fichiers Word en PDF avec barre de progression
with tqdm(total=len(fichiers_word), desc="Conversion en cours", unit="fichier") as pbar:
    # Initialiser l'application Word
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False

    for fichier in fichiers_word:
        # Vérifier si le fichier PDF existe déjà dans le dossier "fichiers convertis"
        fichier_pdf = os.path.splitext(fichier)[0] + "_" + date_du_jour + ".pdf"
        fichier_pdf_converti = os.path.join(dossier_convertis, fichier_pdf)
        if os.path.exists(fichier_pdf_converti):
            pbar.update(1)
            continue

        # Chemin du fichier Word d'origine
        chemin_word = os.path.join(repertoire, fichier)

        # Chemin du fichier PDF de sortie
        chemin_pdf = os.path.join(dossier_convertis, fichier_pdf)

        try:
            # Ouvrir le document Word
            doc = word_app.Documents.Open(chemin_word)

            # Enregistrer le document au format PDF
            doc.SaveAs(chemin_pdf, FileFormat=17)  # 17 représente le format PDF

            # Fermer le document
            doc.Close()

            # Vérifier si le fichier PDF converti est lisible
            with pdfplumber.open(chemin_pdf) as pdf:
                if len(pdf.pages) > 0:
                    # Incrémenter le compteur de fichiers convertis uniquement si le PDF est lisible
                    nb_fichiers_convertis += 1

            pbar.update(1)
        except Exception as e:
            print(f"Erreur lors de la conversion du fichier {fichier} : {e}")

    # Fermer l'application Word
    word_app.Quit()

# Obtenir le nombre total de fichiers Word et de fichiers PDF convertis
nb_fichiers_word = len(fichiers_word)
nb_fichiers_pdf = nb_fichiers_convertis

# Afficher le rapport de conversion
rapport_conversion = f"Nombre total de fichiers Word : {nb_fichiers_word}\n"
rapport_conversion += f"Nombre total de fichiers PDF convertis : {nb_fichiers_pdf}\n"

# Enregistrer le rapport de conversion dans un fichier texte
chemin_rapport = os.path.join(repertoire, "rapport_de_conversion.txt")
with open(chemin_rapport, "w") as fichier_rapport:
    fichier_rapport.write(rapport_conversion)

print("Conversion terminée.")
print("Rapport de conversion enregistré : rapport_de_conversion.txt")
print("le programme quittera automatiquement dans quelques instants")
time.sleep(3)