# 🍪 Trafik Saveurs — Application Bureau Hors Ligne

نظام إدارة المبيعات والمخزون والتوزيع — تطبيق مكتبي يعمل بدون إنترنت 100%

## ✨ Fonctionnalités récentes
- **A icône dédiée** : `icon.ico` / `icon.png` (générés par `make_icon.py`) — affichée dans la barre de titre, la taskbar Windows et le raccourci bureau.
- **100% hors ligne** : base SQLite locale (`data/trafik_saveurs.db`), aucune connexion Internet requise.
- **Import Excel** 📥 :
  - Produits / Achats & Stock : bouton « استيراد إكسل » avec détection automatique des colonnes (arabe/français), aperçu avant import, modèle téléchargeable.
  - Clients : bouton « استيراد زبائن (إكسل) » + نموذج جاهز.
- **Export Excel** 📗 : Produits, Ventes (+ lignes et résumé), Achats (+ lignes), Clients, et Rapport périodique multi-feuilles avec styles, RTL automatique en arabe, filtres et volets figés.
- Export CSV et impression HTML conservés.

## 🚀 Lancement rapide (Windows)
```bat
pip install -r requirements.txt
python app.py
```

## 📦 Créer un exécutable .exe (fichier unique, icône incluse)
Depuis le dossier du projet, avec Python installé :
```bat
pip install pyinstaller openpyxl pillow
pyinstaller --noconfirm --onefile --windowed ^
  --name "TrafikSaveurs" ^
  --icon "icon.ico" ^
  app.py
```
Le fichier `dist\TrafikSaveurs.exe` est autonome (hors ligne). Créez un raccourci vers le Bureau :
clic droit → Envoyer vers → Bureau (créer un raccourci). L'icône Trafik Saveurs apparaîtra dans la taskbar grâce à l'AppUserModelID défini dans le code.

> Alternative sans compilation : utilisez `TrafikSaveurs.vbs` pour lancer `pythonw app.py` sans console, ou créez un raccourci `.lnk` pointant vers `pythonw.exe app.py` avec l'icône `icon.ico`.

## 📁 Structure
```
app.py              # Application Tkinter principale
index.html          # Interface web de secours (locale)
icon.ico / icon.png # Icônes de l'application
make_icon.py        # Régénère les icônes
data/               # Base SQLite (créée automatiquement)
sauvegardes/        # Sauvegardes .db
```

## 🔐 Connexion par défaut
Utilisateur : `admin` — Mot de passe : `admin123`
