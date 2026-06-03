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
    children: [t("5 offres   ·   3 jours / 2 nuits   ·   15 à 80 participants   ·   40 min de Casa & Rabat", { color: MUTED, size: 20 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1400 },
    children: [t("Document de travail — version pour relecture", { italics: true, color: MUTED, size: 18 })] }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ----- INTRO -----
children.push(
  h1("Le positionnement", OCEAN),
  para("The View Bouznika réunit ce qu'aucun concurrent ne combine sur la côte : un wellness médicalisé (InBody, LPG, hammams, bains hydromassants), un front de mer avec golf adjacent et chapiteau, trois restaurants thématisés, le tout sur un site compact entre Casablanca et Rabat. Ces cinq offres transforment ces atouts en expériences mémorables.", { spacing: { after: 120 } }),
  h2("Benchmark concurrentiel", TEAL),
  bullet([t("Mazagan Beach Resort (El Jadida) : ", { bold: true }), t("250 ha, golf 18 trous, MICE intégré. Faille — trop vaste, peu intime, wellness non central.")]),
  bullet([t("Conrad Rabat Arzana : ", { bold: true }), t("premium feutré, gastronomie et banquets. Faille — team building peu structuré en offres packagées.")]),
  bullet([t("Agences spécialisées (Red Rock, Morocco Retreats) : ", { bold: true }), t("formats packagés à thématique RH. Faille — aucun lieu propre, elles dépendent des hôtels.")]),
  h2("Le boulevard stratégique", TEAL),
  bullet("Wellness mesurable (InBody, LPG, detox) intégré au team building — personne ne l'a sur la côte."),
  bullet("Site compact et complet (front de mer + golf + chapiteau + 3 restaurants) à l'inverse de Mazagan dispersé."),
  bullet("Proximité imbattable : 40 min de Casablanca comme de Rabat."),
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
  format: "3 Jours / 2 Nuits  ·  30 à 80 participants",
  vibe: "Un Koh-Lanta version bord de mer — plus glamour, moins boueux. Tribus, totems, immunités, alliances secrètes… et le confort 4 étoiles le soir.",
  cible: "Équipes commerciales, start-ups, agences en quête de fun et de compétition saine.",
  fil: "Le groupe est divisé en tribus (Mer, Vent, Sable, Soleil) — classement et rebondissements sur 3 jours.",
  days: [
    { label: "JOUR 1", title: "Naufragés en terre inconnue", rows: [
      ["10h30", "Arrivée, welcome juice au chiringuito, remise des room-keys"],
      ["11h00", "**Cérémonie d'ouverture** : règles du jeu, tirage des tribus, distribution des bandanas"],
      ["11h30", "**Atelier Totem** : totem en bois flotté, drapeau et cri de guerre sur la plage"],
      ["13h00", "Déjeuner « naufragés » au chiringuito, pieds dans le sable"],
      ["14h30", "**Épreuve 1 — Orientation** : parcours sur le domaine + golf, balises à énigmes"],
      ["16h30", "Pause + 1re rotation wellness (tribu en tête = hammam prioritaire)"],
      ["19h30", "Apéro coucher de soleil"],
      ["20h00", "**Dîner « Survivor BBQ »** autour du feu"],
      ["21h30", "**1er Conseil de tribu** : débrief, stratégie, annonce du classement"],
    ]},
    { label: "JOUR 2", title: "Les grandes épreuves", rows: [
      ["7h30", "Réveil musculaire face mer (option)"],
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Épreuve 2 — Parcours d'aventure** : équilibre dans l'eau, parcours du combattant, puzzle géant, tir à l'arc"],
      ["12h00", "**Épreuve de confort** : la tribu gagnante remporte l'accès VIP wellness"],
      ["13h00", "Déjeuner méditerranéen"],
      ["14h30", "**Rotation parallèle** : wellness (hammam / bain hydromassant / douche à jet) et beach games + golf"],
      ["16h30", "**Épreuve 3 — Dégustation à l'aveugle** (épices, miels, huiles)"],
      ["20h00", "Dîner italien sous le chapiteau"],
      ["21h30", "**Grand Conseil dramatique** : torches, « élimination » symbolique"],
    ]},
    { label: "JOUR 3", title: "La finale & les poteaux", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Épreuve mythique des poteaux** : équilibre sur la plage"],
      ["11h00", "**Orientation finale** pour le collier d'immunité"],
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
      ["12h00", "**Grand jeu de piste de bienvenue** à la découverte du camp"],
      ["13h00", "Déjeuner convivial grandes tablées"],
      ["14h30", "**Goûter à l'ancienne** : crêpes & chocolat chaud au chiringuito"],
      ["15h00", "Ateliers tournants : accrobranche soft / beach games / blason de chambrée"],
      ["19h30", "Dîner"],
      ["21h00", "**Veillée talents** : sketch / chanson par chambrée"],
      ["22h00", "Chamallows grillés au feu de camp"],
    ]},
    { label: "JOUR 2", title: "La grande journée jeux", rows: [
      ["8h00", "Petit-déjeuner"],
      ["9h30", "**Olympiades rétro** : course en sac, tir à la corde, relais à 3 jambes, course à l'œuf"],
      ["12h00", "**Grande bataille d'eau** encadrée à la piscine"],
      ["13h00", "Déjeuner"],
      ["14h30", "Ateliers détente : bracelets brésiliens, photo polaroid, rotation wellness"],
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
      ["11h00", "Arrivée, **thé à la menthe & cornes de gazelle**, eau de fleur d'oranger"],
      ["11h30", "Mot de bienvenue, constitution de la « grande famille »"],
      ["12h00", "**Atelier « Souk »** : chasse aux ingrédients & épices dans le domaine"],
      ["13h00", "Déjeuner méditerranéen du terroir"],
      ["14h30", "Ateliers tournants : **henné & calligraphie** / wellness (hammam au savon noir)"],
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
  presta: ["Maâlem Gnawa + musiciens", "1 conteur (halqa)", "Artistes henné + calligraphe", "Maître artisan zellige"],
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

// ----- PRICING -----
children.push(h1("Grille tarifaire", OCEAN));
children.push(para([t("Prix indicatifs par personne, TTC, base chambre double. Dégressifs selon la taille du groupe. ", {}), t("Chiffres à valider avec le contrôle de gestion (tarif chambre, food cost, masse salariale).", { italics: true, color: MUTED })], { spacing: { after: 140 } }));

const priceHeader = new TableRow({ tableHeader: true, children: [
  cell("Offre", { w: 2826, fill: OCEAN, bold: true, color: "FFFFFF" }),
  cell("10-15 pax", { w: 1550, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("16-30 pax", { w: 1550, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("31-50 pax", { w: 1550, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
  cell("51-80 pax", { w: 1550, fill: OCEAN, bold: true, color: "FFFFFF", align: AlignmentType.CENTER }),
]});
const priceRow = (name, a, b, c, d, accent) => new TableRow({ children: [
  cell(name, { w: 2826, bold: true, color: accent }),
  cell(a, { w: 1550, align: AlignmentType.CENTER }),
  cell(b, { w: 1550, align: AlignmentType.CENTER }),
  cell(c, { w: 1550, align: AlignmentType.CENTER }),
  cell(d, { w: 1550, align: AlignmentType.CENTER }),
]});
children.push(new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [2826, 1550, 1550, 1550, 1550],
  rows: [
    priceHeader,
    priceRow("Koh-Bouznika (3J/2N)", "8 200", "7 200", "6 500", "5 900", KOH),
    priceRow("Colo des Grands (3J/2N)", "7 800", "6 900", "6 200", "—", COLO),
    priceRow("Sound System (3J/2N)", "—", "9 500", "8 200", "7 300", SOUND),
    priceRow("Dar Dyalna (3J/2N)", "8 600", "7 500", "6 800", "—", DAR),
    priceRow("Good Vibes (2J/1N)", "5 200", "4 200", "—", "—", VIBES),
  ],
}));
children.push(new Paragraph({ spacing: { before: 60, after: 200 }, children: [t("Montants en MAD. Prix de référence (jauge standard) : Good Vibes 4 200 · Colo 6 200 · Koh-Bouznika 6 500 · Dar Dyalna 6 800 · Sound System 8 200.", { italics: true, color: MUTED, size: 18 })] }));

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
