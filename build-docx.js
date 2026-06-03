const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageBreak, Header, Footer, PageNumber, TabStopType, TabStopPosition,
} = require("docx");

// ---------- palette ----------
const OCEAN = "0D3B4F", TEAL = "1C7293", GOLD = "B8860B", MUTED = "5B6B70", LIGHT = "F1ECE2";
const KOH = "1F8A70", COLO = "D9701F", SOUND = "7B4BBF", DAR = "B14A2E", VIBES = "B8901F";
const CW = 9026; // A4 content width with 1" margins

// ---------- helpers ----------
const t = (text, opts = {}) => new TextRun({ text, ...opts });

const para = (children, opts = {}) =>
  new Paragraph({ children: Array.isArray(children) ? children : [t(children)], ...opts });

const h1 = (text, color) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 120 },
  children: [t(text, { bold: true, color: color || OCEAN, size: 36 })],
});

const h2 = (text, color) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 80 },
  children: [t(text, { bold: true, color: color || TEAL, size: 26 })],
});

const kv = (label, value) => new Paragraph({
  spacing: { after: 60 },
  children: [t(label + " : ", { bold: true, color: OCEAN }), t(value)],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: "bul", level: 0 },
  spacing: { after: 40 },
  children: Array.isArray(text) ? text : [t(text)],
});

const cell = (children, { w, fill, bold, color, align } = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  margins: { top: 70, bottom: 70, left: 120, right: 120 },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 1, color: "D9D2C4" },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: "D9D2C4" },
    left: { style: BorderStyle.SINGLE, size: 1, color: "D9D2C4" },
    right: { style: BorderStyle.SINGLE, size: 1, color: "D9D2C4" },
  },
  children: (Array.isArray(children) ? children : [children]).map((c) =>
    typeof c === "string"
      ? new Paragraph({ children: [t(c, { bold: bold || false, color: color || "000000" })], alignment: align })
      : c
  ),
});

// Timeline table for a day
const dayTable = (rows, accent) => new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [1300, 7726],
  rows: rows.map((r, i) => new TableRow({
    children: [
      cell(r[0], { w: 1300, fill: i % 2 === 0 ? LIGHT : "FFFFFF", bold: true, color: OCEAN }),
      cell([new Paragraph({ children: parseBold(r[1]) })], { w: 7726, fill: i % 2 === 0 ? LIGHT : "FFFFFF" }),
    ],
  })),
});

// parse **bold** markers into runs
function parseBold(s) {
  const parts = s.split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map((p) =>
    p.startsWith("**") && p.endsWith("**")
      ? t(p.slice(2, -2), { bold: true, color: OCEAN })
      : t(p)
  );
}

const dayHeading = (label, title, accent) => new Paragraph({
  spacing: { before: 200, after: 80 },
  children: [
    t(label + "  ", { bold: true, color: "FFFFFF", size: 20, highlight: undefined }),
    t(title, { bold: true, color: accent, size: 24 }),
  ],
});

// 3-column logistics table
const logiTable = (espaces, materiel, presta, accent) => {
  const w = Math.floor(CW / 3);
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      cell("Espaces", { w, fill: accent, bold: true, color: "FFFFFF" }),
      cell("Matériel", { w, fill: accent, bold: true, color: "FFFFFF" }),
      cell("Prestataires", { w: CW - 2 * w, fill: accent, bold: true, color: "FFFFFF" }),
    ],
  });
  const maxLen = Math.max(espaces.length, materiel.length, presta.length);
  const listCell = (arr, width) => cell(
    arr.map((x) => new Paragraph({ spacing: { after: 30 }, children: [t("– ", { color: accent, bold: true }), t(x)] })),
    { w: width }
  );
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [w, w, CW - 2 * w],
    rows: [
      headerRow,
      new TableRow({ children: [listCell(espaces, w), listCell(materiel, w), listCell(presta, CW - 2 * w)] }),
    ],
  });
};

const twistBox = (lines, accent) => new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [CW],
  rows: [new TableRow({
    children: [cell(
      [
        new Paragraph({ spacing: { after: 60 }, children: [t("LE TWIST QUI CHANGE TOUT", { bold: true, color: "FFFFFF", size: 18 })] }),
        ...lines.map((l) => new Paragraph({ spacing: { after: 40 }, children: parseBoldWhite(l) })),
      ],
      { w: CW, fill: accent }
    )],
  })],
});

function parseBoldWhite(s) {
  const parts = s.split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map((p) =>
    p.startsWith("**") && p.endsWith("**")
      ? t(p.slice(2, -2), { bold: true, color: "FFFFFF" })
      : t(p, { color: "FFFFFF" })
  );
}

const statLine = (pairs, accent) => new Paragraph({
  spacing: { before: 80, after: 160 },
  children: pairs.flatMap(([v, l], i) => [
    t(v, { bold: true, color: accent }),
    t(" " + l + (i < pairs.length - 1 ? "      " : ""), { color: MUTED }),
  ]),
});

// ============ CONTENT ============
const children = [];

// ----- COVER -----
children.push(
  new Paragraph({ spacing: { before: 1800, after: 0 }, alignment: AlignmentType.CENTER,
    children: [t("THE VIEW BOUZNIKA", { bold: true, color: OCEAN, size: 56 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
    children: [t("— EXPERIENCIAH —", { color: GOLD, size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 40 },
    children: [t("Dossier Team Building", { bold: true, color: TEAL, size: 40 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
    children: [t("Cinq expériences clés en main — face à l'océan", { italics: true, color: MUTED, size: 26 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [t("6 offres   ·   2 à 3 jours   ·   15 à 120+ participants   ·   40 min de Casa & Rabat", { color: MUTED, size: 20 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1400 },
    children: [t("Document de travail — version pour relecture", { italics: true, color: MUTED, size: 18 })] }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ----- INTRO -----
children.push(
  h1("Le positionnement", OCEAN),
  para("The View Bouznika réunit ce qu'aucun concurrent ne combine sur la côte : un wellness médicalisé (InBody, LPG, hammams, bains hydromassants), un front de mer avec golf adjacent et chapiteau, trois restaurants thématisés, le tout sur un site compact entre Casablanca et Rabat. Ces cinq offres transforment ces atouts en expériences mémorables.", { spacing: { after: 120 } }),
  h2("Benchmark concurrentiel", TEAL),
  bullet([t("Mazagan Beach Resort (El Jadida) : ", { bold: true }), t("250 ha, golf 18 trous, MICE intégré. Faille — site trop dispersé, peu intime, wellness non central dans le discours MICE.")]),
  bullet([t("Conrad Rabat Arzana : ", { bold: true }), t("resort 5★ front de mer (Plage des Nations), gastronomie premium, spa. Faille — team building peu structuré, pas d'offres packagées thématisées, pas de golf.")]),
  bullet([t("Sofitel Marrakech Lounge & Spa : ", { bold: true }), t("5★ urbain, spa Shiseido, piscine, forte réputation MICE. Faille — pas de front de mer, pas de golf sur site, éloigné des grands bassins d'entreprises côtiers.")]),
  bullet([t("Fairmont Royal Palm Marrakech : ", { bold: true }), t("resort 5★ avec golf 18 trous, cadre exceptionnel. Faille — 3h de route depuis Casablanca, positionnement luxe inaccessible pour groupes corporate standards.")]),
  bullet([t("Club Med La Palmeraie Marrakech : ", { bold: true }), t("formule tout-compris, sports variés, grande capacité. Faille — image « vacances en famille », manque de personnalisation premium et de dimension wellness médicalisée.")]),
  bullet([t("Tikida Golf Palace Agadir : ", { bold: true }), t("5★ avec golf, proche d'Agadir. Faille — trop éloigné de Casablanca-Rabat (5h de route), offre team building non structurée.")]),
  bullet([t("Agences spécialisées (Red Rock, Morocco Retreats, AMA Voyages) : ", { bold: true }), t("formats packagés à thématique RH claire. Faille — aucun lieu propre, elles dépendent des hôtels et ne peuvent proposer le wellness médicalisé intégré.")]),
  h2("Le boulevard stratégique", TEAL),
  bullet([t("Wellness médicalisé et mesurable (InBody, LPG, bains hydromassants, hammam) intégré au programme team building ", {}), t("— personne ne le fait sur la côte atlantique marocaine. Les participants repartent avec des données concrètes sur leur état physique.")]),
  bullet([t("Site 5 étoiles compact et complet ", {}), t("(front de mer + golf 18 trous + chapiteau modulable + 3 restaurants thématisés + spa 3 200 m²) sur un seul domaine, à l'inverse de Mazagan où tout est dispersé sur 250 ha.")]),
  bullet([t("Proximité stratégique imbattable : ", {}), t("40 min de Casablanca, 40 min de Rabat. Aucun autre resort 5 étoiles avec ces équipements n'est aussi bien positionné entre les deux capitales économiques.")]),
  bullet([t("Capacité d'accueil jusqu'à 120+ participants ", {}), t("sur les offres grand format — chapiteau de grande capacité, plage privatisable, espaces modulables — là où les hôtels urbains saturent vite.")]),
  bullet([t("Le Chay Lounge : une signature unique. ", {}), t("Espace thé marocain authentique avec thés de crus, ghriyba signature maison et tea times — un actif de différenciation qu'aucun concurrent ne possède, et qui crée une empreinte mémorielle forte.")]),
  bullet([t("Souvenir tangible systématique ", {}), t("à chaque offre : carnet de colo, collier d'immunité, zellige peint, clip vidéo — le client ne repart jamais sans un objet qui prolonge l'expérience au bureau.")]),
  new Paragraph({ children: [new PageBreak()] }),
);

// ----- OFFERS -----
function offer({ num, name, accent, kicker, format, vibe, cible, fil, days, twist, espaces, materiel, presta, stats }) {
  children.push(
    new Paragraph({ spacing: { after: 20 }, children: [t(`OFFRE ${num} · ${kicker}`, { bold: true, color: accent, size: 18 })] }),
    h1(name, accent),
    new Paragraph({ spacing: { after: 120 }, children: [t(format, { color: MUTED, size: 20 })] }),
    new Paragraph({ spacing: { after: 140 }, children: [t("« " + vibe + " »", { italics: true, color: "000000", size: 26 })] }),
    kv("Cible", cible),
    kv("Fil rouge", fil),
    h2("Programme heure par heure", accent),
  );
  days.forEach((d) => {
    children.push(dayHeading(d.label, d.title, accent), dayTable(d.rows, accent),
      new Paragraph({ spacing: { after: 80 }, children: [] }));
  });
  if (twist) children.push(twistBox(twist, accent), new Paragraph({ spacing: { after: 100 }, children: [] }));
  children.push(h2("Fiche logistique", accent), logiTable(espaces, materiel, presta, accent), statLine(stats, accent));
  children.push(new Paragraph({ children: [new PageBreak()] }));
}

offer({
  num: "A", name: "Koh-Bouznika", accent: KOH, kicker: "Aventure",
  format: "3 Jours / 2 Nuits  ·  30 à 120 participants",
  vibe: "Un Koh-Lanta version bord de mer — plus glamour, moins boueux. Tribus, totems, immunités, alliances secrètes… et le confort 5 étoiles le soir.",
  cible: "Équipes commerciales, start-ups, agences en quête de fun et de compétition saine.",
  fil: "Le groupe est divisé en tribus (Mer, Vent, Sable, Soleil) — classement et rebondissements sur 3 jours.",
  days: [
    { label: "JOUR 1", title: "Naufragés en terre inconnue", rows: [
      ["10h30", "Arrivée, **accueil thé signature & Ghriyba maison au Chay Lounge** — thé de crus, pâtisserie signature The View, remise des room-keys"],
      ["11h00", "**Cérémonie d'ouverture** : règles du jeu, tirage des tribus au sort (Mer, Vent, Sable, Soleil), distribution des bandanas de couleur"],
      ["11h30", "**Atelier Totem** (1h) : par tribu, sculpture d'un totem en bois flotté ramassé sur la plage, confection du drapeau de tribu (peinture tissu), composition du cri de guerre et de la danse d'intimidation — présentés aux autres tribus à 12h30"],
      ["13h00", "Déjeuner « naufragés » au chiringuito, pieds dans le sable"],
      ["14h30", "**Épreuve 1 — Orientation** : parcours de 10 balises GPS sur le domaine (plage, jardin, golf adjacent) — chaque balise révèle une énigme sur l'histoire de Bouznika, de la marque ou de l'équipe. Les tribus partent à 2 min d'intervalle, smartphone interdit, seule une carte papier est fournie. La tribu qui collecte le plus de balises en 90 min remporte les points."],
      ["16h30", "Pause + 1re rotation wellness (tribu en tête = hammam prioritaire)"],
      ["19h30", "Apéro coucher de soleil"],
      ["20h00", "**Dîner « Survivor BBQ »** autour du feu"],
      ["21h30", "**1er Conseil de tribu** : débrief, stratégie, annonce du classement"],
    ]},
    { label: "JOUR 2", title: "Les grandes épreuves", rows: [
      ["7h30", "Réveil musculaire face mer (option)"],
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Épreuve 2 — Parcours d'aventure** (4 stations en simultané, 15 min par station, rotation des tribus) : ① Traversée d'équilibre sur slacklines au-dessus du sable — tomber = revenir au départ ; ② Boot Camp plage : pompes, gainage, squat-sauts en formation tribu — le temps collectif est mesuré ; ③ Puzzle géant en bois (100 pièces) assemblé par la tribu pieds dans l'eau ; ④ Tir à l'arc sur cibles flottantes en bord de mer — 5 flèches par membre, score cumulé. Matériel nécessaire : slacklines × 4, sets tir à l'arc × 4, puzzles × 4, chronomètres"],
      ["12h00", "**Épreuve de confort** : la tribu gagnante remporte l'accès VIP wellness"],
      ["13h00", "Déjeuner méditerranéen"],
      ["14h30", "**Rotation parallèle** : wellness (hammam / bain hydromassant / douche à jet) et beach games + golf"],
      ["16h30", "**Épreuve 3 — Dégustation à l'aveugle** (épices, miels, huiles)"],
      ["20h00", "Dîner italien sous le chapiteau"],
      ["21h30", "**Grand Conseil dramatique** : torches, « élimination » symbolique"],
    ]},
    { label: "JOUR 3", title: "La finale & les poteaux", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Épreuve mythique des poteaux** : 4 poteaux plantés dans le sable à 40 cm de hauteur, un représentant par tribu monte dessus (pieds joints) et doit tenir le plus longtemps possible sans toucher le sol — le dernier debout remporte l'immunité pour sa tribu. Les autres membres soutiennent moralement depuis la plage. Ambiance musicale, maître de jeu en live."],
      ["11h00", "**Course à l'immunité finale** : grand jeu de piste chrono sur l'ensemble du domaine — 5 énigmes à résoudre en équipe, la dernière menant à l'emplacement du trophée caché. La tribu qui le trouve en premier le brandit au signal."],
      ["12h00", "**Grand vote final** + couronnement, remise du trophée"],
      ["13h00", "**Brunch des vainqueurs** au bord de la piscine"],
      ["14h30", "Photo de groupe avec totems, départ"],
    ]},
  ],
  twist: ["Le collier d'immunité ne sauve pas du jeu : il offre un soin wellness premium. Le spa devient une récompense désirée, pas une option oubliée."],
  espaces: ["Plage (épreuves)", "Golf adjacent", "Chiringuito + feu de camp", "Hammam / bains", "Chapiteau (J2 + conseils)"],
  materiel: ["Kit tir à l'arc + cibles", "Slackline / poteaux d'équilibre", "Cordages + balises", "Bandanas 4 couleurs", "Torches de conseil (LED)"],
  presta: ["1 maître de jeu (clé)", "1 coach sécurité", "Agents spa (rotations)", "Surveillant baignade"],
  stats: [["1 / 15", "animateur / pax"], ["5-6", "encadrants (60 pax)"], ["30-80", "jauge"]],
});

offer({
  num: "B", name: "La Colo des Grands", accent: COLO, kicker: "Nostalgie",
  format: "3 Jours / 2 Nuits  ·  30 à 50 participants",
  vibe: "La colonie de vacances qu'on aurait tous voulu avoir, version adulte. Nostalgie, rires, lâcher-prise total — l'offre qui crée le plus de complicité.",
  cible: "Équipes qui se connaissent peu, fusion d'équipes, intégration de nouveaux arrivants.",
  fil: "Badges, carnet de colo, chefs de chambrée, boum du dernier soir. Tout le monde redevient gamin.",
  days: [
    { label: "JOUR 1", title: "Jour de rentrée", rows: [
      ["11h00", "Arrivée, **distribution des kits de colo** (casquette, badge, carnet, gourde)"],
      ["11h30", "Constitution des chambrées + élection des chefs de chambrée"],
      ["12h00", "**Grand jeu de piste de bienvenue** (45 min) : 8 étapes à travers tout le domaine (chiringuito, plage, golf, spa, jardins, restaurants) — chaque étape révèle un indice sur le règlement de la colo et un secret de chambrée. Format rallye : les chambrées partent à 3 min d'intervalle, une fiche de route papier par groupe, pas de smartphone. Bonus : la chambrée qui trouve le mot de passe final en premier décroche la meilleure table au déjeuner."],
      ["13h00", "Déjeuner convivial grandes tablées"],
      ["14h30", "**Goûter à l'ancienne** : crêpes & chocolat chaud au chiringuito"],
      ["15h00", "Ateliers tournants (3 stations × 30 min) : **Boot Camp plage** (pompes, gainage, relais en équipe sur le sable — encadré, accessible à tous, matériel : cordes, cônes, chrono) / **Beach games** (molkky, pétanque géante, badminton) / **Blason de chambrée** (création collective sur grand carton : couleurs, devise, dessin — exposé au mur toute la colo)"],
      ["19h30", "Dîner"],
      ["21h00", "**Veillée talents** : sketch / chanson par chambrée"],
      ["22h00", "Chamallows grillés au feu de camp"],
    ]},
    { label: "JOUR 2", title: "La grande journée jeux", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Olympiades rétro** : course en sac, tir à la corde, relais à 3 jambes, course à l'œuf"],
      ["12h00", "**Grande bataille d'eau** encadrée à la piscine"],
      ["13h00", "Déjeuner"],
      ["14h30", "Ateliers détente (3 stations × 30 min) : **Séance soin hammam & gommage au sel** (restant dans l'univers spa The View) / **Atelier photo polaroid + légendes** (chacun colle ses photos dans le carnet de colo avec la légende du moment) / **Initiation yoga nidra** (relaxation guidée de 20 min, ambiance coussins & musique douce sous le chapiteau)"],
      ["18h00", "Préparation des costumes de la boum"],
      ["20h00", "Dîner"],
      ["21h30", "**LA GRANDE BOUM** déguisée (thème années 90), DJ"],
    ]},
    { label: "JOUR 3", title: "Dernier jour avant les vacances", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Réveil yoga doux** sur la plage"],
      ["10h30", "**Réveille-méninges** : grand quiz sur l'équipe et l'entreprise"],
      ["11h30", "**Remise des diplômes de fin de colo** + derniers mots dans les carnets"],
      ["13h00", "Brunch de départ"],
      ["14h30", "Départ — chacun repart avec son carnet de colo rempli"],
    ]},
  ],
  twist: ["Le carnet de colo rempli — mots des collègues, polaroids, best-of des bêtises. Un souvenir émotionnel surpuissant, à coût quasi nul."],
  espaces: ["Plage (olympiades)", "Piscine (bataille d'eau)", "Chapiteau (boum + veillée)", "Chiringuito (goûter)"],
  materiel: ["Kits colo (casquettes, badges, carnets)", "Matériel olympiades rétro", "Crêpière + chamallows", "Polaroids + pellicules", "Déco « boum 90s »"],
  presta: ["1 DJ (boum)", "1 prof de yoga", "Moniteurs costumés"],
  stats: [["1 / 12", "moniteur / pax"], ["3-4", "moniteurs (40 pax)"], ["30-50", "jauge"]],
});

offer({
  num: "C", name: "Bouznika Sound System", accent: SOUND, kicker: "Festival",
  format: "3 Jours / 2 Nuits  ·  40 à 80 participants",
  vibe: "Coachella au bord de l'Atlantique. Musique, création, énergie collective — et un contenu ultra-instagrammable pour les marques qui veulent rayonner.",
  cible: "Équipes créatives, marketing, agences évènementielles, entreprises tech / com / média.",
  fil: "Chaque équipe est un groupe de musique : nom, logo, hymne — et la scène le dernier soir.",
  days: [
    { label: "JOUR 1", title: "Le line-up", rows: [
      ["11h00", "Arrivée, **bracelet festival** + welcome drink"],
      ["11h30", "Formation des groupes, brief de la mission finale"],
      ["12h00", "**Atelier identité** : nom, logo peint sur bannière, premiers riffs"],
      ["13h00", "Déjeuner « food-truck festival »"],
      ["14h30", "**Initiation percussions** & atelier rythme collectif"],
      ["16h30", "Rotation wellness (« loge des artistes » : hammam / douche à jet)"],
      ["20h00", "Dîner"],
      ["21h30", "**Silent disco** sur la plage — zéro nuisance pour le golf et le voisinage"],
    ]},
    { label: "JOUR 2", title: "En studio & sur scène", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Atelier songwriting** avec coach musical : un tube célèbre aux couleurs de l'entreprise"],
      ["11h30", "Atelier chorégraphie & mise en scène"],
      ["13h00", "Déjeuner"],
      ["14h30", "**Répétitions** par groupes + rotation wellness en parallèle"],
      ["16h30", "Atelier déco de scène / accessoires"],
      ["20h00", "**Dîner italien « backstage »**"],
      ["21h30", "**Blind-test géant** + battle de karaoké"],
    ]},
    { label: "JOUR 3", title: "Le grand concert", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Balance** / dernières répétitions sur la scène"],
      ["12h00", "Déjeuner léger"],
      ["14h00", "**LE GRAND CONCERT** : chaque groupe joue son hymne devant le jury, captation pro"],
      ["15h30", "**Remise des « disques d'or »** + photo de groupe"],
      ["16h00", "Pot de clôture, départ (clip vidéo envoyé après)"],
    ]},
  ],
  twist: ["Captation pro de chaque performance, montée en clip vidéo remis à l'entreprise. Souvenir viral pour le client — et contenu en or pour vendre l'offre suivante."],
  espaces: ["Plage (silent disco)", "Chapiteau / plage (scène)", "Salles de réunion (ateliers)", "Hammam (« loge »)"],
  materiel: ["Silent disco (casques 3 canaux) — location", "Scène + sono + éclairage — location", "Captation vidéo pro — location", "Lot de percussions", "Bannières + peinture"],
  presta: ["1 coach musical / DJ producteur (clé)", "1 technicien son & lumière", "1 vidéaste pro"],
  stats: [["1 / 15", "animateur / pax"], ["16 min", "jauge plancher (rentabilité)"], ["40-80", "jauge optimale"]],
});

offer({
  num: "D", name: "Dar Dyalna", accent: DAR, kicker: "Culture",
  format: "3 Jours / 2 Nuits  ·  30 à 50 participants",
  vibe: "La chaleur de la grande famille marocaine, l'art de recevoir, le partage. Élégante et authentique — parfaite pour les clients étrangers et les incentives.",
  cible: "Agences de voyage, incentives internationaux, comités mixtes, clients « vrai Maroc » haut de gamme.",
  fil: "Le groupe devient une famille qui prépare ensemble une grande Dyafa (festin de fête) pour le dernier jour.",
  days: [
    { label: "JOUR 1", title: "L'accueil à la marocaine", rows: [
      ["11h00", "Arrivée, **accueil au Chay Lounge** : thé signature The View (sélection de crus), Ghriyba signature maison, tea time à la marocaine — la première empreinte de l'hospitalité The View"],
      ["11h30", "Mot de bienvenue, constitution de la « grande famille »"],
      ["12h00", "**Atelier « Souk »** (1h) : chaque binôme reçoit une liste d'ingrédients et d'épices (ras el hanout, safran, argan, cumin, eau de fleur d'oranger…) cachés en 12 points du domaine avec une énigme pour chaque localisation. Budget fictif de 100 DH par équipe, choix à faire — l'équipe qui revient avec le meilleur assortiment (jugé par le chef) démarre l'atelier cuisine avec un avantage."],
      ["13h00", "Déjeuner méditerranéen du terroir"],
      ["14h30", "Ateliers tournants (3 stations × 35 min) : **Pâtisserie Ghriyba** avec notre pâtissier — préparation de la recette signature The View (noix de coco, sésame, eau de fleur d'oranger) que chacun repart avec en sachet personnalisé / **Calligraphie arabe** encadrée par un artiste — chacun calligraphie le nom de l'entreprise ou un mot en arabe sur une carte à emporter / **Hammam au savon noir & gommage kessa** en rotation au spa"],
      ["20h00", "**Dîner du terroir** + conteur (halqa) au coin du feu"],
    ]},
    { label: "JOUR 2", title: "Les mains à la pâte", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Grand atelier cuisine en brigades** : pastilla, tajine, pâtisserie (rotation wellness « détox »)"],
      ["12h30", "**Déjeuner = leurs créations**, jury croisé"],
      ["14h30", "Atelier au choix : **zellige / poterie** ou **initiation Gnawa**"],
      ["16h30", "Temps libre piscine / hammam"],
      ["20h00", "**Soirée lounge orientale** (musique live, derbouka)"],
    ]},
    { label: "JOUR 3", title: "La grande Dyafa", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Préparation collective de la Dyafa** + dressage des grandes tablées"],
      ["11h30", "**Atelier carreau de zellige peint** : chacun y inscrit un mot / un engagement"],
      ["12h30", "**LE GRAND DÉJEUNER DE FÊTE** partagé, musique live"],
      ["14h00", "**Photo de famille** + remise des livrets de recettes"],
      ["14h30", "Départ — avec son zellige + le livre de recettes de l'équipe"],
    ]},
  ],
  twist: ["Chacun repart avec un carreau de zellige peint à la main et le livre de recettes de l'équipe. Ancrage culturel rare — aucun concurrent corporate ne joue cette carte avec autant de soin."],
  espaces: ["Cuisines des restaurants", "Salles / chapiteau (ateliers)", "Hammam traditionnel", "Extérieurs (Dyafa)"],
  materiel: ["Carreaux de zellige bruts + peintures", "Tour de poterie / argile", "Savon noir & gants kessa", "Déco orientale (tapis, lanternes)"],
  presta: ["Maâlem Gnawa + musiciens", "1 conteur (halqa)", "Artiste calligraphe", "Pâtissier (Ghriyba)", "Maître artisan zellige"],
  stats: [["1 / 12", "animateur / pax"], ["1 / 8-10", "chef / cuisiniers amateurs"], ["30-50", "jauge"]],
});

offer({
  num: "E", name: "Good Vibes Only", accent: VIBES, kicker: "Feel-good · Format court",
  format: "2 Jours / 1 Nuit  ·  ≤ 25 participants",
  vibe: "Une parenthèse joyeuse et légère — rire, soleil, bien-être sans prise de tête. L'anti-séminaire stressant, et la porte d'entrée idéale pour faire goûter The View.",
  cible: "Petites équipes (≤ 25), récompense de fin de trimestre, comités qui veulent souffler, week-ends.",
  fil: "Une seule règle : Good Vibes Only. Que du plaisir, zéro PowerPoint.",
  days: [
    { label: "JOUR 1", title: "Lâcher-prise", rows: [
      ["11h30", "Arrivée, **cocktail de bienvenue (mocktails)** au chiringuito face mer"],
      ["12h00", "Mot d'accueil + règle unique : Good Vibes Only"],
      ["12h30", "Déjeuner détendu en terrasse"],
      ["14h00", "**Séance de yoga du rire** (désarme tout le monde en 10 min)"],
      ["15h00", "**Beach games décomplexés** (mölkky, pétanque, badminton, paddle)"],
      ["16h30", "**Rotation wellness « bonheur »** + bain de soleil piscine"],
      ["20h00", "**Dîner italien convivial**"],
      ["21h30", "**Blind-test / karaoké** sur la terrasse"],
    ]},
    { label: "JOUR 2", title: "Recharge & sourire", rows: [
      ["8h00", "Réveil au choix : marche sur la plage, aquagym, ou grasse mat' assumée"],
      ["9h00", "Petit-déjeuner"],
      ["10h00", "**Atelier « cercle des compliments »** : chacun repart avec une carte de mots de ses collègues"],
      ["11h00", "Temps libre piscine / dernier hammam"],
      ["12h00", "**Brunch healthy & gourmand** au bord de la piscine"],
      ["13h30", "Photo de groupe, départ recentrés"],
    ]},
  ],
  twist: ["Court, abordable, émotionnellement marquant — la marge opérationnelle est élevée et les clients reviennent ensuite sur un format 3 jours. Le produit d'appel parfait."],
  espaces: ["Chiringuito + terrasse", "Plage", "Piscine", "Hammam / bains"],
  materiel: ["Beach games (mölkky, pétanque, paddle)", "Cartes « cercle des compliments »", "Système karaoké + sono — location"],
  presta: ["1 coach « yoga du rire » (signature)", "1 animateur blind-test (interne possible)"],
  stats: [["1 / 25", "animateur / pax"], ["2-3", "encadrants"], ["≤ 25", "jauge"]],
});

const GRAND = "1A5276";

offer({
  num: "F", name: "Ocean Summit", accent: GRAND, kicker: "Grand Format · 100-200 participants",
  format: "3 Jours / 2 Nuits  ·  100 à 200 participants",
  vibe: "La seule formule team building 5 étoiles front de mer conçue pour les grands groupes — sans sacrifier ni la qualité, ni le wellness, ni la personnalisation.",
  cible: "Grands groupes corporate (100-200 pax), conventions annuelles, kick-offs régionaux, séminaires de direction avec toutes les équipes réunies.",
  fil: "Le groupe est divisé en 8 à 10 tribus de 15-20 personnes. Chaque tribu vit son propre parcours tout en contribuant à une grande fresque collective révélée le dernier soir.",
  days: [
    { label: "JOUR 1", title: "L'installation — accueil & immersion", rows: [
      ["10h00-12h00", "Arrivée en vagues par tribus (toutes les 20 min), **accueil thé & Ghriyba au Chay Lounge** par groupe de 20 max — flux maîtrisé, pas de goulot"],
      ["11h00", "**Cérémonie d'ouverture plénière** sous le chapiteau : discours du dirigeant, présentation des tribus, règles du jeu Grand Format"],
      ["12h30", "Déjeuner en grandes tablées (buffet méditerranéen sur la terrasse et le chiringuito)"],
      ["14h30", "**Épreuve de cohésion simultanée** : 8-10 ateliers parallèles sur l'ensemble du domaine (plage + golf + jardins), chaque tribu fait son atelier en 45 min puis rotation — maître de jeu central avec talkie-walkie"],
      ["17h00", "**Rotation wellness** : 4 créneaux de 30 min au spa, 2 tribus par créneau — les autres en free time piscine / plage"],
      ["19h30", "Apéritif au coucher de soleil sur la grande terrasse"],
      ["20h30", "**Grand dîner festif** sous le chapiteau + plage — tables de tribu, ambiance musicale live"],
    ]},
    { label: "JOUR 2", title: "Les grandes épreuves collectives", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Olympiades Ocean Summit** : 6 épreuves en simultané sur la plage (beach volley, relais aquatique, tir à la corde, puzzle géant collectif, tir à l'arc, course de radeaux en bois flotté) — toutes les tribus en même temps, 1 encadrant par épreuve"],
      ["13h00", "Déjeuner"],
      ["14h30", "**Atelier fresque collective** : chaque tribu peint un panneau (1m × 1m) représentant sa vision de l'entreprise — les 10 panneaux s'assemblent en une grande murale exposée lors de la cérémonie finale"],
      ["17h00", "Rotation wellness (2e vague)"],
      ["20h00", "**Dîner thématique** — chaque table de tribu présente son panneau en 2 min"],
      ["21h30", "**Grand Conseil** : annonce du classement, suspense maintenu jusqu'au J3"],
    ]},
    { label: "JOUR 3", title: "La grande finale & la fresque révélée", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Épreuve finale simultanée** : chaque tribu envoie 2 représentants pour la grande épreuve d'équilibre (poteaux) — le public des autres tribus encourage depuis la plage"],
      ["11h00", "**Assemblage de la fresque collective** : les 10 panneaux sont réunis et exposés devant tout le groupe — révélation du nom de l'œuvre collective choisie par vote"],
      ["12h00", "**Cérémonie de clôture plénière** : remise du trophée, discours, photo de groupe officielle"],
      ["13h00", "Brunch de victoire collectif"],
      ["14h30", "Départ — chacun repart avec son bandana de tribu + une photo de la fresque collective"],
    ]},
  ],
  twist: [
    "La fresque collective de 10 m² assemblée le dernier jour — chaque tribu y a contribué sans voir le tout. La révélation est un moment de forte émotion. L'œuvre reste exposée à The View Bouznika comme trace de votre groupe.",
    "Le Grand Format ne sacrifie pas la personnalisation : rotation wellness maintenue, accueil en petits groupes au Chay Lounge, ateliers simultanés avec 1 encadrant par poste. La qualité 5 étoiles tient à 200 pax.",
  ],
  espaces: ["Chapiteau plénière (200 pax)", "Plage privatisée (épreuves)", "Terrasses + chiringuito (repas)", "Spa 3 200 m² (rotations)", "Golf adjacent (balises)", "Jardins (ateliers créatifs)"],
  materiel: ["10 panneaux toile + peintures (fresque)", "Kits tribus × 10 (bandanas, drapeaux)", "Matériel olympiades × 10 postes", "Sono + éclairage chapiteau", "10 talkies-walkies (coordination)"],
  presta: ["1 directeur de jeu (chef d'orchestre)", "1 encadrant par poste épreuve (min 10)", "Équipe spa renforcée (rotations 200 pax)", "1 DJ / animateur soirée", "1 vidéaste (capsule J+3)"],
  stats: [["1 / 15-20", "encadrant / pax"], ["10-12", "encadrants minimum"], ["100-200", "jauge"]],
});

// ----- PRICING -----
children.push(h1("Grille tarifaire", OCEAN));
children.push(para([t("Prix indicatifs par personne, TTC, base chambre double. Dégressifs selon la taille du groupe. ", {}), t("Chiffres à valider avec le contrôle de gestion (tarif chambre, food cost, masse salariale).", { italics: true, color: MUTED })], { spacing: { after: 140 } }));

const cw1 = 2426, cw2 = 1320;
const priceHeader = new TableRow({ tableHeader: true, children: [
  cell("Offre", { w: cw1, fill: OCEAN, bold: true, color: "FFFFFF" }),
  cell("10-15 pax", { w: cw2, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("16-30 pax", { w: cw2, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("31-50 pax", { w: cw2, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("51-80 pax", { w: cw2, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("81-120 pax", { w: cw2, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("100-200 pax", { w: cw2, fill: GRAND, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
]});
const priceRow = (name, a, b, c, d, e, f, accent) => new TableRow({ children: [
  cell(name, { w: cw1, bold: true, color: accent }),
  cell(a, { w: cw2, align: AlignmentType.CENTER }),
  cell(b, { w: cw2, align: AlignmentType.CENTER }),
  cell(c, { w: cw2, align: AlignmentType.CENTER }),
  cell(d, { w: cw2, align: AlignmentType.CENTER }),
  cell(e, { w: cw2, align: AlignmentType.CENTER }),
  cell(f, { w: cw2, align: AlignmentType.CENTER }),
]});
children.push(new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [cw1, cw2, cw2, cw2, cw2, cw2, cw2],
  rows: [
    priceHeader,
    priceRow("Koh-Bouznika (3J/2N)", "8 200", "7 200", "6 500", "5 900", "5 400", "—", KOH),
    priceRow("Colo des Grands (3J/2N)", "7 800", "6 900", "6 200", "—", "—", "—", COLO),
    priceRow("Sound System (3J/2N)", "—", "9 500", "8 200", "7 300", "6 800", "—", SOUND),
    priceRow("Dar Dyalna (3J/2N)", "8 600", "7 500", "6 800", "—", "—", "—", DAR),
    priceRow("Good Vibes (2J/1N)", "5 200", "4 200", "—", "—", "—", "—", VIBES),
    priceRow("Ocean Summit — Grand Format (3J/2N)", "—", "—", "—", "—", "5 200", "4 800", GRAND),
  ],
}));
children.push(new Paragraph({ spacing: { before: 60, after: 200 }, children: [t("Montants en MAD, TTC, base chambre double. Prix de référence (jauge standard) : Good Vibes 4 200 · Colo 6 200 · Koh-Bouznika 6 500 · Dar Dyalna 6 800 · Sound System 8 200 · Ocean Summit 4 800. À valider avec le contrôle de gestion.", { italics: true, color: MUTED, size: 18 })] }));

children.push(h2("Suppléments & options", TEAL));
const optRow = (label, val) => new TableRow({ children: [
  cell(label, { w: 6826 }), cell(val, { w: 2200, bold: true, color: OCEAN, align: AlignmentType.RIGHT }),
]});
children.push(new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [6826, 2200],
  rows: [
    optRow("Supplément chambre single", "+ 700 / nuit"),
    optRow("Golf 18 trous (green fee + voiturette)", "+ 900"),
    optRow("Transfert A/R Casablanca / Rabat (autocar)", "+ 250"),
    optRow("Transfert aéroport CMN", "+ 400"),
    optRow("Vidéaste pro + clip (inclus dans Sound System)", "+ 180"),
    optRow("Soin wellness premium additionnel (LPG / InBody+)", "+ 350"),
    optRow("Journée supplémentaire (passage en 4J/3N)", "+ 1 900 à 2 400"),
  ],
}));

children.push(h1("Conditions commerciales", OCEAN));
[
  ["Acompte.", " 40 % à la confirmation, solde à J-7."],
  ["Gratuité.", " 1 participant offert par tranche de 20 (place organisateur)."],
  ["Marge revendeur agences.", " intégrer 10-15 % dans une grille « net agence » dédiée."],
  ["Jauge plancher.", " facturation minimum garantie sur 15 pax (16 pour Sound System)."],
  ["Saisonnalité.", " grille haute saison +10-15 %, basse saison -10 % pour remplir les creux."],
  ["Matériel amorti.", " le kit branding + beach games (30-40 KMAD) s'amortit sur ~10 événements."],
].forEach(([b, txt]) => children.push(bullet([t(b, { bold: true, color: OCEAN }), t(txt)])));

// ----- DOC -----
const doc = new Document({
  creator: "The View Bouznika",
  title: "Dossier Team Building — The View Bouznika",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22, color: "1B2327" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Georgia", color: OCEAN }, paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Georgia", color: TEAL }, paragraph: { spacing: { before: 220, after: 80 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 260 } } } }] },
  ]},
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: 9026 }],
      children: [t("The View Bouznika — Dossier Team Building", { color: MUTED, size: 16 }), t("\tPage ", { color: MUTED, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], color: MUTED, size: 16 })],
    })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("Team-Building-The-View-Bouznika.docx", buf);
  console.log("OK — Team-Building-The-View-Bouznika.docx");
});
