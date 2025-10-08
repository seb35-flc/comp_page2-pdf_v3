import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from PyPDF2 import PdfReader, PdfWriter
import threading

def afficher_aide():
    """Affiche une fenêtre d’aide avec le manuel utilisateur."""
    aide_text = """
=== Manuel d’utilisation – Outil de fusion des secondes pages de PDF ===

Cet outil extrait automatiquement la deuxième page d’une liste de PDF et les fusionne dans un seul fichier.

------------------------------------------------------------
1. LANCEMENT
- Double-cliquez sur l’exécutable.
- Choisissez “Avec filtre(s)” ou “Sans filtre”.

------------------------------------------------------------
2. FILTRES
- Si vous choisissez “Avec filtre(s)”, ajoutez un ou plusieurs mots-clés.
  Exemple : Master, L2H2
- Seuls les PDF dont le NOM contient tous les mots-clés saisis seront pris en compte.

------------------------------------------------------------
3. SÉLECTION DES PDF
- Le programme prend uniquement les fichiers PDF dont le nom commence par 'P_'.
- La recherche s’effectue dans le dossier sélectionné et ses sous-dossiers.

------------------------------------------------------------
4. TRAITEMENT
- Une barre de progression s’affiche.
- Le programme extrait uniquement la page 2 de chaque PDF.

------------------------------------------------------------
5. RÉSULTAT
- Un fichier Fusion_Secondes_Pages.pdf est créé dans le dossier choisi.
- Une fenêtre propose de l’ouvrir directement.

------------------------------------------------------------
REMARQUES
- Les fichiers PDF protégés ou ne contenant qu’une page sont ignorés.
- Le temps de traitement dépend du nombre et de la taille des PDF.
    """

    aide_window = tk.Toplevel()
    aide_window.title("Aide - Fusion des secondes pages PDF")
    aide_window.geometry("600x500")

    text_area = tk.Text(aide_window, wrap="word", font=("Arial", 11))
    text_area.insert(tk.END, aide_text)
    text_area.config(state="disabled")  # lecture seule
    text_area.pack(side="left", fill=tk.BOTH, expand=True, padx=10, pady=10)

    scrollbar = tk.Scrollbar(aide_window, command=text_area.yview)
    scrollbar.pack(side="right", fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)

def choisir_dossier():
    dossier = filedialog.askdirectory(title="Sélectionnez le dossier contenant les PDF")
    return dossier

def demander_filtres():
    """Ouvre une fenêtre pour ajouter plusieurs filtres cumulés."""
    filtres = []

    def ajouter_filtre():
        texte = entry_filtre.get().strip()
        if texte:
            filtres.append(texte)
            listbox.insert(tk.END, texte)
            entry_filtre.delete(0, tk.END)

    def valider():
        root_f.destroy()

    root_f = tk.Toplevel()
    root_f.title("Filtres de recherche")
    root_f.geometry("400x300")

    label = tk.Label(root_f, text="Ajoutez les filtres souhaités (ils s'appliqueront TOUS)")
    label.pack(pady=5)

    entry_filtre = tk.Entry(root_f, width=30)
    entry_filtre.pack(pady=5)

    btn_ajouter = tk.Button(root_f, text="Ajouter le filtre", command=ajouter_filtre)
    btn_ajouter.pack(pady=5)

    listbox = tk.Listbox(root_f, height=8)
    listbox.pack(pady=5, fill=tk.BOTH, expand=True)

    btn_valider = tk.Button(root_f, text="Valider", command=valider)
    btn_valider.pack(pady=10)

    root_f.grab_set()  # bloque l’accès à la fenêtre principale
    root_f.wait_window()

    return filtres

def fusionner_pdfs(dossier, filtres, progress_var, label_status, root):
    writer = PdfWriter()

    fichiers = []
    for racine, _, fichiers_dans_dossier in os.walk(dossier):
        for fichier in fichiers_dans_dossier:
            if fichier.startswith("P_") and fichier.lower().endswith(".pdf"):
                chemin_complet = os.path.join(racine, fichier)
                if all(f.lower() in fichier.lower() for f in filtres):
                    fichiers.append(chemin_complet)

    if not fichiers:
        messagebox.showwarning("Aucun fichier trouvé", "Aucun PDF correspondant aux critères.")
        root.destroy()
        return

    fichiers.sort()
    total = len(fichiers)

    for i, pdf in enumerate(fichiers, start=1):
        try:
            reader = PdfReader(pdf)
            if len(reader.pages) >= 2:
                page = reader.pages[1]  # 2e page
                writer.add_page(page)
        except Exception as e:
            print(f"Erreur avec {pdf}: {e}")

        progress_var.set(int((i / total) * 100))
        label_status.config(text=f"Fusion en cours... ({i}/{total})")
        root.update_idletasks()

    sortie = os.path.join(dossier, "Fusion_Secondes_Pages.pdf")
    with open(sortie, "wb") as f:
        writer.write(f)

    root.destroy()
    messagebox.showinfo("Succès", f"Fusion terminée !\nFichier créé : {sortie}")

    if messagebox.askyesno("Ouvrir le fichier ?", "Souhaitez-vous ouvrir le fichier fusionné ?"):
        try:
            if sys.platform.startswith('win'):
                os.startfile(sortie)
            elif sys.platform.startswith('darwin'):
                subprocess.run(["open", sortie])
            else:
                subprocess.run(["xdg-open", sortie])
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier : {e}")

def lancer_script():
    dossier = choisir_dossier()
    if not dossier:
        return

    utiliser_filtres = messagebox.askyesno("Filtres", "Souhaitez-vous appliquer des filtres de recherche ?")
    filtres = []
    if utiliser_filtres:
        filtres = demander_filtres()

    root_progress = tk.Toplevel()
    root_progress.title("Fusion des PDF")
    root_progress.geometry("400x150")

    label_status = tk.Label(root_progress, text="Fusion en attente...")
    label_status.pack(pady=10)

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(root_progress, variable=progress_var, maximum=100)
    progress_bar.pack(pady=10, fill=tk.X, padx=20)

    thread = threading.Thread(target=fusionner_pdfs, args=(dossier, filtres, progress_var, label_status, root_progress))
    thread.start()

def main():
    root = tk.Tk()
    root.title("Fusionneur de la 2e page des PDF")
    root.geometry("400x250")

    label = tk.Label(root, text="Fusionneur de la 2e page des PDF", font=("Arial", 12))
    label.pack(pady=20)

    btn_lancer = tk.Button(root, text="Sélectionner le dossier et lancer", command=lancer_script, width=30)
    btn_lancer.pack(pady=10)

    btn_aide = tk.Button(root, text="Aide", command=afficher_aide, width=30)
    btn_aide.pack(pady=10)

    btn_quitter = tk.Button(root, text="Quitter", command=root.quit, width=30)
    btn_quitter.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
