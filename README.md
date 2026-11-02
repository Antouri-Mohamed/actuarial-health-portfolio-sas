# 📊 Analyse de Sinistralité & Provisionnement d'un Portefeuille Santé (SAS)

Projet quantitatif et actuariel réalisé dans le cadre du Master 1 MAS (parcours IMSA) à Aix-Marseille Université.

---

## 🎯 Problématique & Périmètre
* **Périmètre :** Étude sur 24 mois de soins (2018–2019) d'un portefeuille de complémentaire santé.
* **Volumétrie :** 46 857 mois-garantie (unité d'exposition), 3 064 assurés, 2 000 contrats et 112 garanties.
* **Objectif :** Analyser la dégradation du ratio Prestations / Cotisations (S/P), identifier les facteurs de risque et évaluer les provisions pour sinistres à payer (PSAP).

---

## 🔍 Principaux Résultats
1. **Évolution du S/P :** Le ratio S/P global est de 80,0 %, mais il subit une hausse de +14 points, passant de 72,5 % en 2018 à 86,5 % en 2019.
2. **Identification de l'anti-sélection :** 
   * Les assurés présents dès 2018 voient leur S/P s'améliorer en 2019 (68,6 %).
   * Les 1 345 nouveaux entrants de 2019 affichent un S/P de **168,6 %**.
   * Consommation concentrée au démarrage : S/P de **430 %** dès le mois d'adhésion.
3. **Poste de soins critique :** Explosion du coût de l'hospitalisation (+61 % par mois-garantie), portée à 58 % par les nouveaux assurés.
4. **Effet de composition :** Forte concentration des nouveaux sur la garantie d'entrée de gamme **ATP1** (S/P de 149 %).

---

## 📐 Méthode Actuarielle : Provisionnement Chain Ladder
* Construction d'un **triangle de liquidation trimestriel** (trimestre de soins × délai de règlement) pour estimer la charge finale au 31/12/2019.
* Sans provision, le S/P 2019 paraissait à 71,6 % (la dégradation était masquée par le décalage de facturation).
* **Provision estimée :** 179 012 € pour 2019, amenant le S/P estimé à 85,7 % (écart de seulement 0,9 % avec la réalité observée jusqu'en 2024).

---

## 🛠️ Outils & Compétences mobilisées
* **SAS Base / SAS SQL :** Étape DATA, jointures relationnelles, macros automatisées (%profil, %top, %sp), calculs matriciels.
* **Excel avancé :** Tableaux croisés dynamiques de synthèse et contrôle strict à écart nul.
* **Techniques actuarielles :** Mesure de l'exposition (mois-garantie), Chain Ladder, indice de Gini (0,71), tests de robustesse (écrêtement des sinistres majeurs).

---

## 📁 Contenu du dépôt
* `Code_SAS_projet.sas` : Programme SAS complet (sections 0 à 21).
* `Presentation_projet.pdf` : Support de présentation avec l'ensemble des visualisations.
* `Synthese_projet.pdf` : Note de synthèse détaillée de l'étude.
