# Experienciah — The View Bouznika · Team Building

Dossiers commerciaux **Team Building** pour l'hôtel **The View Bouznika** (groupe EXPERIENCIAH, marque The View), à destination de la Direction Marketing.

Le projet part d'un brief reçu par email et décline deux gammes d'offres team building destinées aux entreprises, agences évènementielles et agences de voyage.

---

## 🌐 Sites en ligne (Netlify)

| Gamme | Description | Fichier source | Lien |
|---|---|---|---|
| **Team Building** | 5 offres « fun » et originales | `team-building-the-view-bouznika.html` | https://the-view-bouznika-teambuilding.netlify.app |
| **Collection Signature** | 4 offres premium enrichies | `signature-the-view-bouznika.html` | https://the-view-bouznika-signature.netlify.app |

---

## 📁 Contenu

```
.
├── team-building-the-view-bouznika.html   # Site des 5 offres fun
├── signature-the-view-bouznika.html       # Site des 4 offres premium
├── Team-Building-The-View-Bouznika.docx   # Document Word (5 offres) pour relecture
├── build-docx.js                          # Script de génération du .docx (docx-js)
├── package.json                           # Dépendance : docx
└── TR_ Offres Team Building TVB.eml        # Brief d'origine
```

---

## 🎯 Les 9 offres

**Gamme Team Building (fun)**
1. 🏝️ Koh-Bouznika — aventure type Koh-Lanta (3J/2N)
2. 🎒 La Colo des Grands — nostalgie / colonie de vacances (3J/2N)
3. 🎶 Bouznika Sound System — festival musical (3J/2N)
4. 🇲🇦 Dar Dyalna — immersion culturelle marocaine (3J/2N)
5. ☀️ Good Vibes Only — feel-good, format court (2J/1N)

**Collection Signature (premium)**
1. Reset — Détox & Reconnexion (3J/2N)
2. Drive — Performance & Énergie (3J/2N)
3. Savor — Gastronomie & Cohésion (3J/2N)
4. Horizon — Séminaire stratégique & Bien-être (2J/1N)

Chaque offre comprend : programme heure par heure, fiche logistique (espaces / matériel / prestataires / ratios), souvenir tangible, et — pour la Collection Signature — cadre de mesure du ROI, argumentaire de vente et accroche email.

---

## 🏨 Les atouts de l'hôtel (socle des offres)

- Wellness médicalisé : InBody, LPG, hammams, douches à jet, bains hydromassants, piscine intérieure chauffée (centre de 3 200 m²)
- Front de mer, grande piscine extérieure, golf adjacent, terrains de tennis, vélos
- Chapiteau + immenses salles de réunion
- 3 restaurants : méditerranéen, italien, chiringuito
- Position : à 40 min de Casablanca comme de Rabat (45 min de l'aéroport CMN)

---

## 🛠️ Régénérer le document Word

```bash
npm install
node build-docx.js
```

---

*Document de travail — version pour relecture.*
