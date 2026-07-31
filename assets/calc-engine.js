// calc-engine.js — gedeeld rekenmotor voor alle pagina's van dit domein.
// Bron methode Hobin/IGAK (tabel + parameters): Officiële rekentool kinderalimentatie
// (methode Hobin), versie maart 2026, gepubliceerd door de Belgische hoven en rechtbanken.
// Officiële rekentool: https://www.rechtbanken-tribunaux.be/nl/node/3520
// Handleiding: https://www.ordevanvlaamsebalies.be/ovb_filesystem/ovb/Documenten/burgerlijk-recht/Handleiding-rekentool-Hobin.pdf
// Tabel geëxtraheerd op 2026-07-31 uit het officiële Excel-bestand (rekentool-hobin-maart-2026.xlsx,
// tabblad "IGAK 2025") — waarden hier zijn een directe kopie, niet herberekend of geschat.
// Bij twijfel over een resultaat: vergelijk met de officiële rekentool (link hierboven).
// Bron-JSON (leesbaar, gedocumenteerd): zie /data/igak-table-2025.json en /data/hobin-params.json

// --- IGAK 2025 tabel: aandeel van het gezinsinkomen per kind, per leeftijd (0-24) en "kolom" ---
// IGAK_AGE_TABLE[leeftijd][kolom-1] = aandeel (0-1) van het gezinsinkomen voor 1 kind van die leeftijd.
// leeftijd-index 0..23 = 0..23 jaar, index 24 = 24 jaar of meer.
const IGAK_AGE_TABLE = [[0.151448, 0.14922, 0.146993, 0.108775, 0.107171, 0.105568, 0.093987, 0.09265, 0.091225, 0.087483, 0.086147, 0.0849, 0.082851, 0.081604, 0.080445, 0.079287, 0.078129, 0.076882, 0.076437, 0.075278, 0.074209, 0.07412, 0.073051, 0.071893], [0.151448, 0.14922, 0.146993, 0.108775, 0.107171, 0.105568, 0.093987, 0.09265, 0.091225, 0.087483, 0.086147, 0.0849, 0.082851, 0.081604, 0.080445, 0.079287, 0.078129, 0.076882, 0.076437, 0.075278, 0.074209, 0.07412, 0.073051, 0.071893], [0.151448, 0.14922, 0.146993, 0.108775, 0.107171, 0.105568, 0.093987, 0.09265, 0.091225, 0.087483, 0.086147, 0.0849, 0.082851, 0.081604, 0.080445, 0.079287, 0.078129, 0.076882, 0.076437, 0.075278, 0.074209, 0.07412, 0.073051, 0.071893], [0.151448, 0.14922, 0.146993, 0.108775, 0.107171, 0.105568, 0.093987, 0.09265, 0.091225, 0.087483, 0.086147, 0.0849, 0.082851, 0.081604, 0.080445, 0.079287, 0.078129, 0.076882, 0.076437, 0.075278, 0.074209, 0.07412, 0.073051, 0.071893], [0.151448, 0.14922, 0.146993, 0.108775, 0.107171, 0.105568, 0.093987, 0.09265, 0.091225, 0.087483, 0.086147, 0.0849, 0.082851, 0.081604, 0.080445, 0.079287, 0.078129, 0.076882, 0.076437, 0.075278, 0.074209, 0.07412, 0.073051, 0.071893], [0.151448, 0.14922, 0.146993, 0.108775, 0.107171, 0.105568, 0.093987, 0.09265, 0.091225, 0.087483, 0.086147, 0.0849, 0.082851, 0.081604, 0.080445, 0.079287, 0.078129, 0.076882, 0.076437, 0.075278, 0.074209, 0.07412, 0.073051, 0.071893], [0.153846, 0.151584, 0.149321, 0.110498, 0.108869, 0.10724, 0.095475, 0.094118, 0.09267, 0.088869, 0.087511, 0.086244, 0.084163, 0.082896, 0.081719, 0.080543, 0.079367, 0.0781, 0.077647, 0.076471, 0.075385, 0.075294, 0.074208, 0.073032], [0.156322, 0.154023, 0.151724, 0.112276, 0.110621, 0.108966, 0.097011, 0.095632, 0.094161, 0.090299, 0.08892, 0.087632, 0.085517, 0.08423, 0.083034, 0.081839, 0.080644, 0.079356, 0.078897, 0.077701, 0.076598, 0.076506, 0.075402, 0.074207], [0.158879, 0.156542, 0.154206, 0.114112, 0.11243, 0.110748, 0.098598, 0.097196, 0.095701, 0.091776, 0.090374, 0.089065, 0.086916, 0.085607, 0.084393, 0.083178, 0.081963, 0.080654, 0.080187, 0.078972, 0.07785, 0.077757, 0.076636, 0.075421], [0.16152, 0.159145, 0.15677, 0.11601, 0.114299, 0.112589, 0.100238, 0.098812, 0.097292, 0.093302, 0.091876, 0.090546, 0.088361, 0.087031, 0.085796, 0.084561, 0.083325, 0.081995, 0.08152, 0.080285, 0.079145, 0.07905, 0.07791, 0.076675], [0.164251, 0.161836, 0.15942, 0.117971, 0.116232, 0.114493, 0.101932, 0.100483, 0.098937, 0.094879, 0.09343, 0.092077, 0.089855, 0.088502, 0.087246, 0.08599, 0.084734, 0.083382, 0.082899, 0.081643, 0.080483, 0.080386, 0.079227, 0.077971], [0.167076, 0.164619, 0.162162, 0.12, 0.118231, 0.116462, 0.103686, 0.102211, 0.100639, 0.096511, 0.095037, 0.093661, 0.0914, 0.090025, 0.088747, 0.087469, 0.086192, 0.084816, 0.084324, 0.083047, 0.081867, 0.081769, 0.08059, 0.079312], [0.17, 0.1675, 0.165, 0.1221, 0.1203, 0.1185, 0.1055, 0.104, 0.1024, 0.0982, 0.0967, 0.0953, 0.093, 0.0916, 0.0903, 0.089, 0.0877, 0.0863, 0.0858, 0.0845, 0.0833, 0.0832, 0.082, 0.0807], [0.17276, 0.170219, 0.167679, 0.124082, 0.122253, 0.120424, 0.107213, 0.105688, 0.104062, 0.099794, 0.09827, 0.096847, 0.09451, 0.093087, 0.091766, 0.090445, 0.089124, 0.087701, 0.087193, 0.085872, 0.084652, 0.084551, 0.083331, 0.08201], [0.17574, 0.173156, 0.170571, 0.126223, 0.124362, 0.122501, 0.109062, 0.107512, 0.105858, 0.101516, 0.099965, 0.098518, 0.09614, 0.094693, 0.093349, 0.092005, 0.090661, 0.089214, 0.088697, 0.087353, 0.086113, 0.086009, 0.084769, 0.083425], [0.1785, 0.175875, 0.17325, 0.128205, 0.126315, 0.124425, 0.110775, 0.1092, 0.10752, 0.10311, 0.101535, 0.100065, 0.09765, 0.09618, 0.094815, 0.09345, 0.092085, 0.090615, 0.09009, 0.088725, 0.087465, 0.08736, 0.0861, 0.084735], [0.18137, 0.178703, 0.176036, 0.130266, 0.128346, 0.126426, 0.112556, 0.110956, 0.109249, 0.104768, 0.103168, 0.101674, 0.09922, 0.097726, 0.09634, 0.094953, 0.093566, 0.092072, 0.091539, 0.090152, 0.088871, 0.088765, 0.087484, 0.086097], [0.18424, 0.181531, 0.178821, 0.132328, 0.130377, 0.128426, 0.114337, 0.112712, 0.110978, 0.106426, 0.1048, 0.103283, 0.10079, 0.099273, 0.097864, 0.096455, 0.095046, 0.093529, 0.092987, 0.091578, 0.090278, 0.090169, 0.088869, 0.08746], [0.187, 0.18425, 0.1815, 0.13431, 0.13233, 0.13035, 0.11605, 0.1144, 0.11264, 0.10802, 0.10637, 0.10483, 0.1023, 0.10076, 0.09933, 0.0979, 0.09647, 0.09493, 0.09438, 0.09295, 0.09163, 0.09152, 0.0902, 0.08877], [0.18976, 0.186969, 0.184179, 0.136292, 0.134283, 0.132274, 0.117763, 0.116088, 0.114302, 0.109614, 0.10794, 0.106377, 0.10381, 0.102247, 0.100796, 0.099345, 0.097894, 0.096331, 0.095773, 0.094322, 0.092982, 0.092871, 0.091531, 0.09008], [0.19274, 0.189906, 0.187071, 0.138433, 0.136392, 0.134351, 0.119612, 0.117912, 0.116098, 0.111336, 0.109635, 0.108048, 0.10544, 0.103853, 0.102379, 0.100905, 0.099431, 0.097844, 0.097277, 0.095803, 0.094443, 0.094329, 0.092969, 0.091495], [0.1955, 0.192625, 0.18975, 0.140415, 0.138345, 0.136275, 0.121325, 0.1196, 0.11776, 0.11293, 0.111205, 0.109595, 0.10695, 0.10534, 0.103845, 0.10235, 0.100855, 0.099245, 0.09867, 0.097175, 0.095795, 0.09568, 0.0943, 0.092805], [0.19826, 0.195344, 0.192429, 0.142397, 0.140298, 0.138199, 0.123038, 0.121288, 0.119422, 0.114524, 0.112775, 0.111142, 0.10846, 0.106827, 0.105311, 0.103795, 0.102279, 0.100646, 0.100063, 0.098547, 0.097147, 0.097031, 0.095631, 0.094115], [0.20124, 0.198281, 0.195321, 0.144538, 0.142407, 0.140276, 0.124887, 0.123112, 0.121218, 0.116246, 0.11447, 0.112813, 0.11009, 0.108433, 0.106894, 0.105355, 0.103816, 0.102159, 0.101567, 0.100028, 0.098608, 0.098489, 0.097069, 0.09553], [0.204, 0.201, 0.198, 0.14652, 0.14436, 0.1422, 0.1266, 0.1248, 0.12288, 0.11784, 0.11604, 0.11436, 0.1116, 0.10992, 0.10836, 0.1068, 0.10524, 0.10356, 0.10296, 0.1014, 0.09996, 0.09984, 0.0984, 0.09684]];

// KOLOM_MAPPING[aantal_kinderen][categorie-1] = kolomnummer (1-24) in IGAK_AGE_TABLE.
// categorie 1 = gezinsinkomen tot 2000€/maand, 2 = 2001-4000€, 3 = 4001-6000€, 4 = 6001-8000€, 5 = meer dan 8000€.
const IGAK_KOLOM_MAPPING = {"1": [3, 2, 1, 2, 3], "2": [6, 5, 4, 5, 6], "3": [9, 8, 7, 8, 9], "4": [12, 11, 10, 11, 12], "5": [15, 14, 13, 14, 15], "6": [18, 17, 16, 17, 18], "7": [21, 20, 19, 20, 21], "8": [24, 23, 22, 23, 24]};

const IGAK_INCOME_CAP = 10000;

function igakCategorie(gezinsinkomen) {
  if (gezinsinkomen > 8000) return 5;
  if (gezinsinkomen >= 6001) return 4;
  if (gezinsinkomen >= 4001) return 3;
  if (gezinsinkomen >= 2001) return 2;
  return 1;
}

// Kost van 1 kind van gegeven leeftijd, in EUR/maand (aandeel × geplafonneerd gezinsinkomen).
function igakKindKost(leeftijd, aantalKinderen, gezinsinkomen) {
  const inkomenGeplafonneerd = Math.min(gezinsinkomen, IGAK_INCOME_CAP);
  const categorie = igakCategorie(inkomenGeplafonneerd);
  const n = Math.min(Math.max(aantalKinderen, 1), 8);
  const kolom = IGAK_KOLOM_MAPPING[String(n)][categorie - 1];
  const leeftijdIndex = Math.min(Math.max(Math.round(leeftijd), 0), 24);
  const aandeel = IGAK_AGE_TABLE[leeftijdIndex][kolom - 1];
  return aandeel * inkomenGeplafonneerd;
}

// --- Hobin-parameters (leefloon default, verblijfsregelingen) ---
// Bron: officiële rekentool, update leefloon 03/2026 (https://www.mi-is.be/nl/tools-ocmw/bedragen-equivalent-leefloon)
const LEEFLOON_ALLEENSTAANDE = 1340.47;
const LEEFLOON_SAMENWONEND = 670.235;

const VERBLIJFSREGELINGEN = [{"id": "week_week", "label": "Week/week (50/50)", "ouder1_verblijfsgebonden": 0.5, "ouder2_verblijfsgebonden": 0.5, "ouder1_verblijfsoverstijgend": 0.5, "ouder2_verblijfsoverstijgend": 0.5}, {"id": "8_6", "label": "8/6 dagen + helft vakantie", "ouder1_verblijfsgebonden": 0.4465, "ouder2_verblijfsgebonden": 0.5535, "ouder1_verblijfsoverstijgend": 0.5, "ouder2_verblijfsoverstijgend": 0.5}, {"id": "9_5", "label": "9/5 dagen + helft vakantie", "ouder1_verblijfsgebonden": 0.3958, "ouder2_verblijfsgebonden": 0.6042, "ouder1_verblijfsoverstijgend": 0.15, "ouder2_verblijfsoverstijgend": 0.85}, {"id": "10_4", "label": "10/4 dagen + helft vakantie", "ouder1_verblijfsgebonden": 0.3452, "ouder2_verblijfsgebonden": 0.6548, "ouder1_verblijfsoverstijgend": 0.15, "ouder2_verblijfsoverstijgend": 0.85}, {"id": "11_3", "label": "11/3 dagen + helft vakantie", "ouder1_verblijfsgebonden": 0.2945, "ouder2_verblijfsgebonden": 0.7055, "ouder1_verblijfsoverstijgend": 0.05, "ouder2_verblijfsoverstijgend": 0.95}, {"id": "we_helft_vak", "label": "1 weekend/2 + helft vakantie", "ouder1_verblijfsgebonden": 0.2438, "ouder2_verblijfsgebonden": 0.7562, "ouder1_verblijfsoverstijgend": 0.05, "ouder2_verblijfsoverstijgend": 0.95}];

// Volledige Hobin-berekening (vereenvoudigde modus: forfaitaire niet-samendrukbare kost
// op basis van leefloon, geen rekening met voordelen in natura, huurinkomsten of
// afzonderlijke groeipakket-toewijzing per kind — voor complexe dossiers verwijzen we
// naar de officiële rekentool, zie link bovenaan dit bestand).
function berekenHobin(params) {
  const {
    inkomenOuder1, inkomenOuder2,
    ouder1Alleenstaand, ouder2Alleenstaand,
    kinderen, // array van leeftijden
    verblijfId,
    groeipakketOuder1 = 0, groeipakketOuder2 = 0
  } = params;

  const nietSamendrukbaar1 = ouder1Alleenstaand ? LEEFLOON_ALLEENSTAANDE : LEEFLOON_SAMENWONEND;
  const nietSamendrukbaar2 = ouder2Alleenstaand ? LEEFLOON_ALLEENSTAANDE : LEEFLOON_SAMENWONEND;

  const beschikbaar1 = Math.max(0, inkomenOuder1 - nietSamendrukbaar1);
  const beschikbaar2 = Math.max(0, inkomenOuder2 - nietSamendrukbaar2);
  const beschikbaarTotaal = beschikbaar1 + beschikbaar2;

  const aandeel1 = beschikbaarTotaal > 0 ? beschikbaar1 / beschikbaarTotaal : 0.5;
  const aandeel2 = 1 - aandeel1;

  const gezinsinkomen = inkomenOuder1 + inkomenOuder2 + groeipakketOuder1 + groeipakketOuder2;
  const aantalKinderen = kinderen.length;

  const brutoKostPerKind = kinderen.map(leeftijd => igakKindKost(leeftijd, aantalKinderen, gezinsinkomen));
  const brutoKostTotaal = brutoKostPerKind.reduce((a, b) => a + b, 0);
  const groeipakketTotaal = groeipakketOuder1 + groeipakketOuder2;
  const nettoKostTotaal = Math.max(0, brutoKostTotaal - groeipakketTotaal);

  const moetBetalen1 = nettoKostTotaal * aandeel1;
  const moetBetalen2 = nettoKostTotaal * aandeel2;

  const verblijf = VERBLIJFSREGELINGEN.find(v => v.id === verblijfId) || VERBLIJFSREGELINGEN[0];

  const natura1Pct = (verblijf.ouder1_verblijfsgebonden + verblijf.ouder1_verblijfsoverstijgend) / 2;
  const natura2Pct = (verblijf.ouder2_verblijfsgebonden + verblijf.ouder2_verblijfsoverstijgend) / 2;

  const naturaBruto1 = brutoKostTotaal * natura1Pct;
  const naturaBruto2 = brutoKostTotaal * natura2Pct;

  const naturaNetto1 = Math.max(0, naturaBruto1 - groeipakketOuder1);
  const naturaNetto2 = Math.max(0, naturaBruto2 - groeipakketOuder2);

  const verschil1 = moetBetalen1 - naturaNetto1;
  const verschil2 = moetBetalen2 - naturaNetto2;

  let onderhoudsbijdrage = 0;
  let betaler = null;
  if (verschil1 > 0) { onderhoudsbijdrage = verschil1; betaler = 1; }
  else if (verschil2 > 0) { onderhoudsbijdrage = verschil2; betaler = 2; }

  return {
    beschikbaar1, beschikbaar2, aandeel1, aandeel2,
    brutoKostTotaal, nettoKostTotaal, brutoKostPerKind,
    moetBetalen1, moetBetalen2,
    naturaNetto1, naturaNetto2,
    onderhoudsbijdrage, betaler,
    onderhoudsbijdragePerKind: aantalKinderen > 0 ? onderhoudsbijdrage / aantalKinderen : 0
  };
}

// --- Statbel gezondheidsindex (gladde/lisse index, januari van elk jaar, basis 2013=100) ---
// Bron: statbel.fgov.be — zelfde nationale reeks als op de FR-zustersite gebruikt.
const STATBEL_INDICES = { 2015:102.7, 2016:103.6, 2017:105.1, 2018:107.3, 2019:109.5, 2020:110.4, 2021:111.4, 2022:115.9, 2023:129.0, 2024:131.8, 2025:134.9, 2026:137.3 };

function berekenIndexering(bedrag, jaarBasis, jaarDoel) {
  const idxBasis = STATBEL_INDICES[jaarBasis];
  const idxDoel = STATBEL_INDICES[jaarDoel];
  const ratio = idxDoel / idxBasis;
  return { geindexeerd: Math.round(bedrag * ratio * 100) / 100, ratio };
}

// --- Fiscale aftrekbaarheid — versioned, NIET hardcoden elders. Bron: data/tax-deduction-rate.json ---
const TAX_DEDUCTION = { year: 2026, deductiblePct: 60 };

function fmt(n) { return Math.round(n).toLocaleString('nl-BE') + ' €'; }
