from itertools import permutations, product
import pandas as pd
import numpy as np
import heapq
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt

# Pfad zu Ihrer Excel-Datei
file_path = 'C:/Users/alexa/OneDrive/Dokumente/Masterarbeit/Sheet.xlsx'

# Laden des Excel-Sheets in einen DataFrame
df = pd.read_excel(file_path, index_col=0)

# Gruppen von Projektionen erstellen
gruppen = {}
for projektion in df.columns:
    gruppe = projektion[0]
    if gruppe not in gruppen:
        gruppen[gruppe] = []
    gruppen[gruppe].append(projektion)

# Erstellen eines Lookup-Dictionary für die Indizes
projektion_index = {proj: i for i, proj in enumerate(df.columns)}

# Umwandeln des DataFrame in ein numpy Array
matrix = df.to_numpy()

# Erstellen eines Heaps für jede Projektion
top_kombinationen_pro_projektion = {proj: [] for proj in df.columns}

# Funktion zur Berechnung der Summe für eine Kombination
def berechne_summe_kombi(kombi):
    summe = 0
    for i in range(len(kombi)):
        for j in range(i + 1, len(kombi)):
            summe += matrix[projektion_index[kombi[j]], projektion_index[kombi[i]]]
    return summe

# Zählvariable für die Anzahl der verarbeiteten Kombinationen
zaehler = 0

# Durchgehen aller Permutationen der Gruppen und Berechnung der Summen
for kombi in product(*gruppen.values()):
    if len(set(kombi)) == len(kombi):
        summe = berechne_summe_kombi(kombi)
        durchschnitt = summe / len(kombi)
        for proj in kombi:
            if len(top_kombinationen_pro_projektion[proj]) < 20:
                heapq.heappush(top_kombinationen_pro_projektion[proj], (durchschnitt, summe, kombi))
            else:
                heapq.heappushpop(top_kombinationen_pro_projektion[proj], (durchschnitt, summe, kombi))
        zaehler += 1
        if zaehler % 1000000 == 0:
            print(f"\nVerarbeitete Kombinationen: {zaehler}")
            for proj, heap in top_kombinationen_pro_projektion.items():
                top_kombinationen = sorted(heap, reverse=True)[:20]
                print(f"Aktuelle Top 20 Kombinationen für {proj}:")
                for rang, (durchschnitt, summe, kombi) in enumerate(top_kombinationen, 1):
                    print(f"  {rang}. {kombi}, Summe: {summe}, Durchschnitt: {durchschnitt:.2f}")

# Vor dem Scree-Diagramm: Eingabe der Anzahl der Top-Kombinationen
while True:
    anzahl_top_kombinationen = int(input("Geben Sie die Anzahl der Top-Kombinationen pro Projektion ein (zwischen 1 und 20): "))

    # Anpassen der Top-Kombinationen basierend auf der gewählten Anzahl
    angepasste_top_kombinationen_pro_projektion = {proj: sorted(heap, reverse=True)[:anzahl_top_kombinationen] for proj, heap in top_kombinationen_pro_projektion.items()}

    # Erstellen der Distanzmatrix
    alle_kombinationen = [kombi for top_liste in angepasste_top_kombinationen_pro_projektion.values() for _, _, kombi in top_liste]
    distanzmatrix = [[sum(1 for p1, p2 in zip(kombi1, kombi2) if p1 != p2) for kombi2 in alle_kombinationen] for kombi1 in alle_kombinationen]

    # Erstellen des Scree-Diagramms
    intra_cluster_distanz = []
    for k in range(1, 11):
        kmeans = KMeans(n_clusters=k)
        kmeans.fit(distanzmatrix)
        intra_cluster_distanz.append(kmeans.inertia_)

    plt.figure(figsize=(10, 6))
    plt.plot(range(1, 11), intra_cluster_distanz, marker='o')
    plt.title('Scree-Diagramm')
    plt.xlabel('Anzahl der Cluster')
    plt.ylabel('Intra-Cluster-Distanz')
    plt.show()

    # Eingabeaufforderung für Speichern oder neue Eingabe
    weiter = input("Geben Sie 'speichern' ein, um fortzufahren, oder 'neu', um eine andere Anzahl von Top-Kombinationen zu wählen: ")
    if weiter.lower() == 'speichern':
        break

# Eingabeaufforderung für die Anzahl der Cluster
optimale_cluster_anzahl = int(input("Bitte geben Sie die optimale Anzahl an Clustern ein: "))

# Führe KMeans mit der optimalen Anzahl an Clustern aus
kmeans_optimal = KMeans(n_clusters=optimale_cluster_anzahl)
kmeans_optimal.fit(distanzmatrix)
cluster_labels = kmeans_optimal.labels_

# Berechnung des Vorkommens jeder Variable in jedem Cluster
variable_vorkommen_pro_cluster = {i: {} for i in range(optimale_cluster_anzahl)}
for cluster_index in range(optimale_cluster_anzahl):
    cluster_indices = np.where(cluster_labels == cluster_index)[0]
    cluster_kombinationen = [alle_kombinationen[index] for index in cluster_indices]
    gesamtanzahl_kombinationen = len(cluster_kombinationen)

    # Zähle die Vorkommen jeder Variable
    for kombination in cluster_kombinationen:
        for variable in kombination:
            if variable not in variable_vorkommen_pro_cluster[cluster_index]:
                variable_vorkommen_pro_cluster[cluster_index][variable] = 0
            variable_vorkommen_pro_cluster[cluster_index][variable] += 1

    # Umwandlung in Prozentwerte
    for variable, vorkommen in variable_vorkommen_pro_cluster[cluster_index].items():
        variable_vorkommen_pro_cluster[cluster_index][variable] = (vorkommen / gesamtanzahl_kombinationen) * 100

# Ausgabe des Vorkommens der Variablen pro Cluster in Prozent
for cluster_index in range(optimale_cluster_anzahl):
    print(f"\nVariable Vorkommen in Cluster {cluster_index + 1}:")
    for variable, prozent in variable_vorkommen_pro_cluster[cluster_index].items():
        print(f"{variable}: {prozent:.2f}%")

# Export der Ergebnisse in eine Excel-Datei
with pd.ExcelWriter('C:/Users/alexa/OneDrive/Desktop/Clusterergebnisse.xlsx') as writer:
    # Erstes Sheet: Top 20 Kombinationen pro Projektion
    for proj, heap in top_kombinationen_pro_projektion.items():
        top_kombinationen = sorted(heap, reverse=True)[:20]
        kombi_df = pd.DataFrame(top_kombinationen, columns=['Durchschnitt', 'Summe', 'Kombination'])
        kombi_df.to_excel(writer, sheet_name=f'Top 20 {proj}', index=False)

    # Zweites Sheet: Distanzmatrix
    df_distanzmatrix = pd.DataFrame(distanzmatrix, index=[f'{proj}{i+1}' for proj in top_kombinationen_pro_projektion for i in range(anzahl_top_kombinationen)], columns=[f'{proj}{i+1}' for proj in top_kombinationen_pro_projektion for i in range(anzahl_top_kombinationen)])
    df_distanzmatrix.to_excel(writer, sheet_name='Distanzmatrix')

    # Drittes Sheet: Clusterergebnisse
    cluster_df = pd.DataFrame(variable_vorkommen_pro_cluster).T.fillna(0)
    cluster_df_transposed = cluster_df.transpose()  # Spalten und Zeilen vertauschen
    cluster_df_transposed.to_excel(writer, sheet_name='Clusterergebnisse transponiert')

