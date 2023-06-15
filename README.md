# Conversion-word-en-pdf
Convertir des fichiers word en pdf au sein du dossier source avec fichier log 

Le script va convertir les fichiers word présent dans le dossier en pdf, en créant un autre dossier dans le répertoire source afin d'éviter la pollution visuel.
Un petit fichier log sera crée aussi pour check si la conversion est bien faitre entre le nombre de word et pdf. 

Comme module j'utilise « PdfLumber » pour ouvrir chaque fichier pdf converti et vérifier si le nombre de pages est supérieur à zéro. Si c'est le cas, le fichier PDF est considéré comme lisible.
OS, DateTime, Win32com, et TQDM

Crée par Mickael CHATELIN
