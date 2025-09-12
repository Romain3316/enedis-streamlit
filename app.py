import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def choisir_periode_et_mode(annees):
    choix = {"valeur": None, "debut": None, "fin": None, "mode": None}

    def exporter():
        selection = combo.get()
        mode = combo_mode.get()
        if selection == "Période personnalisée":
            debut = entry_debut.get()
            fin = entry_fin.get()
            if not debut or not fin:
                messagebox.showerror("Erreur", "Veuillez entrer une date de début et une date de fin.")
                return
            choix["valeur"] = "personnalisee"
            choix["debut"] = debut
            choix["fin"] = fin
        else:
            choix["valeur"] = selection
        choix["mode"] = mode
        fenetre.destroy()

    fenetre = tk.Tk()
    fenetre.title("Sélection de la période et du mode horaire")

    tk.Label(fenetre, text="Choisissez la période à exporter :").pack(pady=10)

    options = ["Toutes les données"] + sorted(annees) + ["Période personnalisée"]
    combo = ttk.Combobox(fenetre, values=options, state="readonly")
    combo.set("Toutes les données")
    combo.pack(pady=5)

    # Zone pour période personnalisée
    frame_perso = tk.Frame(fenetre)
    tk.Label(frame_perso, text="Début (JJ/MM/AAAA) :").grid(row=0, column=0, padx=5, pady=2)
    entry_debut = tk.Entry(frame_perso)
    entry_debut.grid(row=0, column=1, padx=5, pady=2)
    tk.Label(frame_perso, text="Fin (JJ/MM/AAAA) :").grid(row=1, column=0, padx=5, pady=2)
    entry_fin = tk.Entry(frame_perso)
    entry_fin.grid(row=1, column=1, padx=5, pady=2)
    frame_perso.pack(pady=5)

    # Choix mode de gestion des heures
    tk.Label(fenetre, text="Gestion des jours à 23h / 25h :").pack(pady=10)
    options_mode = ["Heures réelles (23h / 25h)", "Forcer 24h/jour"]
    combo_mode = ttk.Combobox(fenetre, values=options_mode, state="readonly")
    combo_mode.set("Heures réelles (23h / 25h)")
    combo_mode.pack(pady=5)

    bouton = tk.Button(fenetre, text="Valider", command=exporter)
    bouton.pack(pady=10)

    fenetre.mainloop()
    return choix

def choisir_format_export():
    choix = {"format": None}

    def exporter_csv():
        choix["format"] = "csv"
        fenetre.destroy()

    def exporter_excel():
        choix["format"] = "excel"
        fenetre.destroy()

    fenetre = tk.Tk()
    fenetre.title("Choix du format d'export")

    tk.Label(fenetre, text="Sélectionnez le format de sortie :").pack(pady=10)

    tk.Button(fenetre, text="CSV (.csv)", command=exporter_csv).pack(pady=5)
    tk.Button(fenetre, text="Excel (.xlsx)", command=exporter_excel).pack(pady=5)

    fenetre.mainloop()
    return choix["format"]

def traiter_fichier(fichier):
    # 1. Lecture
    df = pd.read_excel(fichier)

    # 2. Garder seulement colonnes utiles
    df = df[["Unité", "Horodate", "Valeur"]]

    # 3. Supprimer "VAR"
    df = df[df["Unité"].str.upper().isin(["W", "KW"])]

    # 4. Conversion date
    df["Horodate"] = pd.to_datetime(df["Horodate"], dayfirst=True, errors="coerce")

    # 5. Agrégation horaire → moyenne
    df = df.set_index("Horodate")
    df = df.resample("1H").mean(numeric_only=True)
    df = df.dropna().reset_index()

    # 6. Extraire années disponibles
    annees_dispo = df["Horodate"].dt.year.unique().astype(str).tolist()

    # 7. Sélecteur utilisateur
    choix = choisir_periode_et_mode(annees_dispo)

    # 8. Filtrage selon choix
    if choix["valeur"] not in ["Toutes les données", "personnalisee"]:
        annee = int(choix["valeur"])
        df = df[df["Horodate"].dt.year == annee]

    elif choix["valeur"] == "personnalisee":
        debut = pd.to_datetime(choix["debut"], dayfirst=True, errors="coerce")
        fin = pd.to_datetime(choix["fin"], dayfirst=True, errors="coerce")
        if pd.isna(debut) or pd.isna(fin):
            messagebox.showerror("Erreur", "Format de date invalide. Utilisez JJ/MM/AAAA.")
            return
        df = df[(df["Horodate"] >= debut) & (df["Horodate"] <= fin)]

    # 9. Gestion des jours 23h/25h
    if choix["mode"] == "Forcer 24h/jour":
        full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
        df = df.set_index("Horodate").reindex(full_range)
        df.index.name = "Horodate"
        df["Valeur"] = df["Valeur"].interpolate(method="linear")  # interpolation si trous
        df = df.reset_index()

    # 10. Vérification des trous
    trous = []
    full_range = pd.date_range(df["Horodate"].min(), df["Horodate"].max(), freq="1H")
    missing = full_range.difference(df["Horodate"])
    if not missing.empty:
        for d in missing:
            trous.append(d.strftime("%d/%m/%Y %H:%M:%S"))

    # 11. Format final
    df["Date"] = df["Horodate"].dt.date
    df["Heure"] = df["Horodate"].dt.time
    df = df.rename(columns={"Valeur": "Moyenne_Conso"})
    df_final = df[["Unité", "Date", "Heure", "Moyenne_Conso"]]

    # 12. Choix du format d'export
    format_export = choisir_format_export()

    # 13. Export
    if format_export == "csv":
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
        )
        if save_path:
            df_final.to_csv(save_path, index=False, sep=";")
            msg = f"Export terminé en CSV : {save_path}\n"

    elif format_export == "excel":
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        if save_path:
            df_final.to_excel(save_path, index=False)
            msg = f"Export terminé en Excel : {save_path}\n"

    else:
        return

    # 14. Message de synthèse
    if trous:
        msg += f"\n⚠️ Données manquantes aux horodatages suivants :\n" + "\n".join(trous[:10])
        if len(trous) > 10:
            msg += f"\n... et {len(trous)-10} autres."
    else:
        msg += "\n✅ Pas de données manquantes."
    messagebox.showinfo("Résultat", msg)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    fichier = filedialog.askopenfilename(title="Choisissez un fichier Enedis", filetypes=[("Excel files", "*.xlsx *.xls")])
    if fichier:
        traiter_fichier(fichier)

