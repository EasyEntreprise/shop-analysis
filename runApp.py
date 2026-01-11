###########################
## Importation Librairies
############
from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
import os
import sys
import subprocess
import webbrowser
from PIL import Image, ImageTk


# Fonction pour trouver le chemin des ressources
def resource_path(relative_path):
    """
    Permet de trouver le bon chemin vers les ressources (images, etc.),
    que le programme soit lancé en .py ou en .exe (PyInstaller).
    """
    try:
        # Si le programme est compilé avec PyInstaller
        base_path = sys._MEIPASS
    except Exception:
        # Sinon, chemin du dossier courant
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Fonction pour charger et redimensionner une image si elle est trop grande
def load_image(relative_path, max_width=800, max_height=600):
    """
    Charge une image en tenant compte du chemin PyInstaller et la redimensionne
    si sa taille dépasse max_width ou max_height.
    """
    image_path = resource_path(relative_path)
    img = Image.open(image_path)

    # Si l'image est plus grande que les dimensions max, on la redimensionne
    if img.width > max_width or img.height > max_height:
        img.thumbnail((max_width, max_height))

    return ImageTk.PhotoImage(img)


#### La Fonction de Fermeture fénétre ###
"""
Cette fonction permet de Se déconnecter du programme pour revenir à la fénétre d'identification
"""
def fermer_fenetre():
    root.destroy()

#########################
#  Paramettrage Fenetre
#####

root = Tk()
root.wm_attributes("-alpha", 1)
ttk.Style().configure("TP.Frame", background= "#fd0309")
root.title("Shop Business analysis")
root.geometry('500x400+500+180')
root.iconbitmap(resource_path("images/dashboard_icon_182989.ico"))
import threading



# Chargement des images 
#--------------------------------------------------------------------------
screenx_path = resource_path("images/Visualisations.png")
if os.path.exists(screenx_path):
    screenx = PhotoImage(file=screenx_path)
else:
    print(f"⚠️ Fichier image introuvable : {screenx_path}")
#--------------------------------------------------------------------------
profil = load_image("images/rodrigue-N.png", max_width=60, max_height=60) 
#--------------------------------------------------------------------------
logo = load_image("images/easy-logo.jpg", max_width=60, max_height=60) 
#--------------------------------------------------------------------------
close = load_image("images/close.png", max_width=50, max_height=50) 
#--------------------------------------------------------------------------


# Configuration 

style = ttk.Style(root)
style.configure("YB.TButton", background = "#B0C4DE", font=('Arial Narrow', 15))
style.configure("ZB.TButton", background = "#B0C4DE", font=('Aptos (Corps)', 12))
style.configure("YL.TLabel", background = "#2A2A2B",  foreground = "white", font=('Arial Narrow', 10))
style.configure("ZL.TLabel", background = "#B0C4DE",  foreground = "black", font=('AaronBecker-Heavy', 18))
style.configure("KL.TLabel", background = "#B0C4DE",  foreground = "black", font=('Aptos (Corps)', 12))
style.configure("SL.TLabel", background = "#2A2A2B",  foreground = "#ECF0F4", font=('AaronBecker-Heavy', 18))
style.configure("Y.TFrame", background = "#B0C4DE")
style.configure("YF.TFrame", background = "#2A2A2B")
style.configure("ZF.TFrame", background = "#01010E")

############################
## Lancer server streamlit
##############
"""
Dans cette partie nous allons lancer le server streamlit en appuyant sur un bouton
"""

def run_streamlit():
    # Lancer server Streamlit
    subprocess.Popen(["streamlit", "run", "martha.py", "--server.headless=true"]) # Ouvre un navigateur automatiquement
    webbrowser.open("http://localhost:8501")

def open_dashboard():
    # Lancer dans un thread séparé pour ne pas bloquer tkinter
    threading.Thread(target= run_streamlit).start()

def ouvrir_lien():
    webbrowser.open("https://easyentreprise-shop-analysis-martha-daujgn.streamlit.app/")


## Frame tkinter

body = ttk.Frame(root, width= 500, height= 380, style = "Y.TFrame")
head = ttk.Frame(root, width= 500, height= 90, style = "YF.TFrame")
foot = ttk.Frame(root, width= 500, height= 20, style = "ZF.TFrame")

# Profils

ttk.Label(head, text="""   SHOPS BUSINESS ANALYSIS 
            DASHBOARD""", width= 80, style= "SL.TLabel").place(x= 110, y= 10)
ttk.Label(head, text= "Rodrigue NSINSULU ", style= "YL.TLabel", width= 20, image= profil, compound= TOP).place(x= 5, y= 2)
ttk.Label(body, style= "YL.TLabel", width= 20, image= logo, compound= TOP).place(x= 437, y= 335)
ttk.Label(body, text="Welcome to our Data Analysis application", style= "ZL.TLabel", width= 50).place(x= 15, y= 150)
ttk.Label(body, text="Please select a button to launch the analysis server", style= "KL.TLabel", width= 50).place(x= 50, y= 190)

# Buttons 
ttk.Button(body, text = "Local Server", image= screenx, compound= LEFT, style = "YB.TButton", command= lambda : [open_dashboard()]).place(x= 100, y= 240)
ttk.Button(body, text = "Server Online", image= screenx, compound= LEFT, style = "YB.TButton", command= lambda : [ouvrir_lien()]).place(x= 245, y= 240)
#ttk.Button(body, text="Close App", image= close, compound= LEFT, style = "ZB.TButton", command= lambda : [fermer_fenetre()]).place(x= 150, y= 300)

foot.place(x=0, y=380)
head.place(x=0, y=0)
body.place(x=0, y=0)
root.mainloop()
