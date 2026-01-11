# launcher.py
import os
import sys
import threading
import time
import webbrowser

# Nom du fichier Streamlit à lancer (modifie si nécessaire)
APP_FILE = "martha.py"
PORT = 8501
# Répertoire courant = dossier où se trouve le launcher
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def run_streamlit():
    """
    Lance Streamlit via son CLI interne (API) en s'assurant que l'argument
    argv contient ce qu'il faut.
    """
    # Change le répertoire pour que les chemins relatifs d'app.py fonctionnent
    os.chdir(BASE_DIR)

    # Prépare sys.argv comme si on avait appelé "streamlit run app.py"
    sys.argv = ["streamlit", "run", APP_FILE, "--server.port", str(PORT), "--server.headless", "true"]

    # Importer et appeler le CLI interne de streamlit
    try:
        # streamlit >=1.0, on utilise le module cli interne
        from streamlit.web import cli as stcli
    except Exception:
        # fallback pour anciennes versions
        import subprocess as stcli

    # Exécute
    sys.exit(stcli.main())

def open_browser_when_ready(url, timeout=10):
    """Ouvre le navigateur lorsque le serveur répondra (on attend simplement un peu)."""
    time.sleep(2)  # petit délai initial
    webbrowser.open(url, new=2)

if __name__ == "__main__":
    url = f"http://localhost:{PORT}"
    t = threading.Thread(target=run_streamlit, daemon=True)
    t.start()

    # Ouvre le navigateur (ce thread peut être lancé dès que Streamlit est up)
    open_browser_when_ready(url)

    # Maintiens le launcher actif tant que le thread existe
    try:
        while t.is_alive():
            t.join(timeout=1)
    except KeyboardInterrupt:
        print("Arrêt demandé. Fermeture...")
        sys.exit(0)
