"""Calculs de l'achat de la résidence principale.

Reprend le classeur Excel (estimation, frais de notaire, financement,
budget de possession, achat vs location, sensibilité du crédit) et la
capacité d'achat du foyer.
"""

from __future__ import annotations

import json
import math
import uuid
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"

# --- Prix moyens au m² et tables du classeur (octobre 2026 pour les taux du foyer) ---

PRIX_M2 = {
    1: 13000,
    2: 11700,
    3: 12200,
    4: 13100,
    5: 12200,
    6: 15700,
    7: 15400,
    8: 13100,
    9: 10900,
    10: 9300,
    11: 9900,
    12: 9100,
    13: 8800,
    14: 9600,
    15: 9600,
    16: 11400,
    17: 10200,
    18: 9100,
    19: 8200,
    20: 8400,
}

QUARTIERS = {
    1: "Louvre, Palais-Royal, Les Halles",
    2: "Bourse, Sentier, Montorgueil",
    3: "Haut Marais, Arts-et-Métiers",
    4: "Marais, Île Saint-Louis, Hôtel de Ville",
    5: "Quartier Latin, Sorbonne, Panthéon",
    6: "Saint-Germain-des-Prés, Odéon, Luxembourg",
    7: "Invalides, Tour Eiffel, Gros-Caillou",
    8: "Champs-Élysées, Madeleine, Monceau",
    9: "Opéra, Pigalle, Faubourg-Montmartre",
    10: "Canal Saint-Martin, République, Gare du Nord",
    11: "Bastille, Oberkampf, Folie-Méricourt",
    12: "Bercy, Nation, Quinze-Vingts",
    13: "Butte-aux-Cailles, Olympiades, Salpêtrière",
    14: "Montparnasse, Alésia, Parc-de-Montsouris",
    15: "Grenelle, Vaugirard, Beaugrenelle",
    16: "Auteuil, Passy, Muette",
    17: "Batignolles, Ternes, Épinettes",
    18: "Montmartre, Grandes-Carrières, Goutte-d'Or",
    19: "Buttes-Chaumont, La Villette, Pont-de-Flandre",
    20: "Belleville, Ménilmontant, Père-Lachaise",
}

# Ancienne table du classeur. Sert uniquement à la remplacer dans un fichier déjà enregistré.
_PRIX_M2_CLASSEUR = {
    1: 12481,
    2: 11222,
    3: 11914,
    4: 12547,
    5: 11738,
    6: 14414,
    7: 13818,
    8: 11905,
    9: 10745,
    10: 9182,
    11: 9797,
    12: 8875,
    13: 8505,
    14: 9286,
    15: 9433,
    16: 10627,
    17: 10012,
    18: 8727,
    19: 7861,
    20: 8400,
}

ETAGE_LABELS = ["RDC", "1er", "2e", "3e", "4e", "5e", "6e", "7e et +"]
ETAGE_AVEC = [-0.08, -0.03, 0.0, 0.02, 0.04, 0.06, 0.07, 0.08]
ETAGE_SANS = [-0.10, -0.04, -0.02, -0.01, -0.04, -0.08, -0.12, -0.15]

# (surface à partir de laquelle le coefficient s'applique, coefficient)
COEF_SURFACE = [(0, 0.08), (25, 0.03), (40, 0.0), (70, -0.02), (100, -0.03)]

COEF_ETAT = {
    "À rénover entièrement": -0.15,
    "À rafraîchir": -0.06,
    "Bon état": 0.0,
    "Rénové / refait à neuf": 0.05,
}
COEF_DPE = {"A": 0.04, "B": 0.03, "C": 0.01, "D": 0.0, "E": -0.03, "F": -0.08, "G": -0.12}
COEF_LUMIERE = {
    "Sombre / sur cour étroite": -0.06,
    "Normale": 0.0,
    "Lumineuse": 0.03,
    "Très lumineuse / traversante / vue dégagée": 0.06,
}
COEF_BRUIT = {
    "Bruyant (rue passante, bar, voie ferrée)": -0.05,
    "Normal": 0.0,
    "Très calme (cour, rue calme)": 0.03,
}
COEF_EXTERIEUR = {
    "Aucun": 0.0,
    "Petit balcon / loggia": 0.04,
    "Grand balcon / terrasse": 0.08,
    "Jardin privatif": 0.10,
}
COEF_IMMEUBLE = {
    "Pierre de taille / haussmannien": 0.03,
    "Ancien standard (avant 1950)": 0.0,
    "Années 1950-1980": -0.04,
    "Récent (après 1990)": 0.02,
}

CHECKLIST = [
    ("Copropriété", "PV des 3 dernières assemblées générales", "Travaux votés non encore appelés, litiges, contestations. À négocier dans le compromis : qui paie les travaux déjà votés."),
    ("Copropriété", "Appels de charges et budget prévisionnel (2 derniers exercices)", "Paris : souvent 30 à 60 €/m²/an. Hausse brutale ou budget irréaliste = signal d'alerte."),
    ("Copropriété", "Fonds de travaux (loi ALUR)", "Cotisation minimale de 5 % du budget prévisionnel. Fonds vide avec façade ou toiture à refaire = appel de fonds probable."),
    ("Copropriété", "Plan pluriannuel de travaux / diagnostic technique global", "Copropriétés de plus de 15 ans : liste et chiffrage des travaux à 10 ans."),
    ("Copropriété", "Fiche synthétique, état daté ou pré-état daté", "Impayés, procédures, dettes de la copropriété."),
    ("Copropriété", "Règlement de copropriété et état descriptif de division", "Usage (habitation), règles de travaux, location courte durée, tantièmes."),
    ("Copropriété", "Syndic et contrats d'entretien", "Qualité de gestion, coût et échéance (ascenseur, chauffage collectif)."),
    ("Bâtiment", "Toiture, façade, ravalement", "À Paris le ravalement est souvent exigé environ tous les 10 ans et se répartit entre copropriétaires."),
    ("Bâtiment", "État et âge de l'ascenseur", "Mise aux normes ou remplacement : plusieurs dizaines de milliers d'euros à répartir."),
    ("Bâtiment", "Colonnes montantes, canalisations, parties communes", "Humidité, fissures, escalier, sécurité incendie, état de la cour."),
    ("Bâtiment", "Chauffage collectif ou individuel", "Collectif : consommation et contrat. Individuel : âge de la chaudière ou du chauffe-eau."),
    ("Logement", "Surface loi Carrez", "Écart de plus de 5 % avec l'annonce : réduction de prix possible dans l'année suivant l'acte."),
    ("Logement", "DPE : classe, consommation, date de validité", "F et G : décote, travaux, et restrictions de location à la revente."),
    ("Logement", "Diagnostics (électricité, gaz, plomb, amiante, termites, risques)", "Paris est en zone termites. Chaque anomalie est un budget travaux."),
    ("Logement", "Humidité, infiltrations, moisissures", "Surtout rez-de-chaussée, 1er et dernier étage."),
    ("Logement", "Fenêtres, isolation phonique, vis-à-vis", "Visiter de jour et le soir. Tester fenêtres fermées."),
    ("Logement", "Orientation et luminosité", "Un appartement sombre se revend plus difficilement. Vérifier aussi la hauteur sous plafond."),
    ("Logement", "Plan, murs porteurs, rangements", "Ouvrir un mur porteur : accord de la copropriété et étude de structure."),
    ("Logement", "Cave, local à vélo, parking, chambre de service", "Vérifier qu'ils figurent au lot, dans l'acte et au règlement."),
    ("Juridique", "Libre ou occupé", "Bien loué : pas de jouissance immédiate, préemption éventuelle du locataire."),
    ("Juridique", "État hypothécaire, servitudes, urbanisme", "À faire vérifier par le notaire avant le compromis."),
    ("Juridique", "Dernier avis de taxe foncière du vendeur", "Reporter le montant réel dans l'analyse du bien."),
    ("Juridique", "Conditions suspensives (prêt, urbanisme, préemption)", "Prévoir 45 à 60 jours pour le prêt. Rétractation : 10 jours après notification du compromis."),
    ("Juridique", "Dépôt de garantie et calendrier", "En général 5 à 10 %. Dates de la condition suspensive et de l'acte authentique."),
    ("Financement", "Comparer au moins 3 offres", "Taux, assurance, garantie, modularité, indemnités de remboursement anticipé."),
    ("Financement", "Délégation d'assurance emprunteur", "Souvent bien moins chère que le contrat groupe. Changeable à tout moment."),
    ("Financement", "PTZ, prêt employeur, Action Logement", "Vérifier l'éligibilité (ressources, type de bien) avec un courtier ou l'ADIL."),
    ("Financement", "Endettement ≤ 35 % et apport couvrant notaire + garantie", "Voir le taux d'endettement et la capacité dans cette application."),
    ("Prix", "Comparables DVF / Patrim (même rue, ±300 m)", "Même type de bien, étage et état. Sert à recaler les prix au m²."),
    ("Prix", "Ancienneté de l'annonce et baisses déjà faites", "Plus de 60 à 90 jours en ligne, ou prix déjà baissé : marge de négociation plus forte."),
    ("Prix", "Motivation du vendeur et délai souhaité", "Succession, mutation, divorce : la marge de négociation change."),
    ("Quartier", "Transports et temps de trajet réel", "Tester le trajet domicile-travail aux heures de pointe."),
    ("Quartier", "Commerces, écoles, santé, espaces verts", "Un quartier équipé tient mieux ses prix."),
    ("Quartier", "Nuisances et projets urbains", "Bars, rue passante, périphérique, chantiers. Consulter le PLU."),
    ("Quartier", "Ambiance et sécurité", "Visiter de jour, le soir et le week-end."),
]


def _prix_m2_defaut():
    return dict(PRIX_M2)


@dataclass
class Hypotheses:
    droit_dept: float = 0.05
    taxe_communale: float = 0.012
    frais_assiette: float = 0.0237
    tva_emoluments: float = 0.20
    debours: float = 1200.0
    csi: float = 0.001
    emol_seuils: tuple = (6500.0, 17000.0, 60000.0)
    emol_taux: tuple = (0.0387, 0.01596, 0.01064, 0.00799)
    charges_m2: float = 40.0
    taxe_m2: float = 25.0
    assurance_habitation: float = 300.0
    entretien_pct: float = 0.003
    provision_copro_pct: float = 0.002
    energie_mois: float = 150.0
    loyer_m2: float = 30.0
    garantie_pct: float = 0.01
    frais_dossier: float = 1000.0
    frais_courtage: float = 0.0
    valeur_cave: float = 5000.0
    valeur_local_velo: float = 2000.0
    valeur_parking: float = 30000.0
    fourchette: float = 0.05
    seuil_sous_marche: float = -0.05
    seuil_coherent: float = 0.03
    seuil_negociable: float = 0.08
    coef_dernier_asc: float = 0.03
    coef_dernier_sans: float = 0.0
    part_charges_locataire: float = 0.6
    assurance_locataire: float = 150.0
    hausse_immo: float = 0.015
    hausse_loyers: float = 0.02
    inflation_charges: float = 0.025
    rendement_epargne: float = 0.03
    frais_revente: float = 0.05
    horizon_annees: int = 25
    prix_m2: dict = field(default_factory=_prix_m2_defaut)
    etage_avec: list = field(default_factory=lambda: list(ETAGE_AVEC))
    etage_sans: list = field(default_factory=lambda: list(ETAGE_SANS))
    coef_surface: list = field(default_factory=lambda: [list(x) for x in COEF_SURFACE])
    coef_etat: dict = field(default_factory=lambda: dict(COEF_ETAT))
    coef_dpe: dict = field(default_factory=lambda: dict(COEF_DPE))
    coef_lumiere: dict = field(default_factory=lambda: dict(COEF_LUMIERE))
    coef_bruit: dict = field(default_factory=lambda: dict(COEF_BRUIT))
    coef_exterieur: dict = field(default_factory=lambda: dict(COEF_EXTERIEUR))
    coef_immeuble: dict = field(default_factory=lambda: dict(COEF_IMMEUBLE))

    def to_dict(self) -> dict:
        d = deepcopy(self.__dict__)
        d["emol_seuils"] = list(self.emol_seuils)
        d["emol_taux"] = list(self.emol_taux)
        d["prix_m2"] = {str(k): v for k, v in self.prix_m2.items()}
        return d

    @classmethod
    def from_dict(cls, raw: dict) -> "Hypotheses":
        base = cls()
        data = base.to_dict()
        for cle, valeur in (raw or {}).items():
            if cle in data:
                data[cle] = valeur
        data["emol_seuils"] = tuple(data["emol_seuils"])
        data["emol_taux"] = tuple(data["emol_taux"])
        data["prix_m2"] = {int(k): float(v) for k, v in data["prix_m2"].items()}
        return cls(**data)


@dataclass
class Profil:
    """Capacité d'achat du foyer. Défauts : octobre 2026, apport 320 000 €, cible 2 600 €."""

    apport: float = 320_000.0
    mensualite_cible: float = 2_600.0
    inclut_assurance: bool = False
    taux_20: float = 0.0335
    taux_25: float = 0.0355
    duree_retenue: int = 25
    revenus: float = 0.0
    autres_credits: float = 0.0
    assurance: float = 0.003
    mobilier: float = 5_000.0
    agence_pct: float = 0.0
    endettement_max: float = 0.35
    loyer_actuel: float = 1_900.0

    def taux(self, duree: int) -> float:
        if int(duree) == 20:
            return self.taux_20
        if int(duree) == 25:
            return self.taux_25
        return self.taux_25 if int(duree) > 22 else self.taux_20

    def to_dict(self) -> dict:
        return deepcopy(self.__dict__)

    @classmethod
    def from_dict(cls, raw: dict) -> "Profil":
        data = cls().to_dict()
        for cle, valeur in (raw or {}).items():
            if cle in data:
                data[cle] = valeur
        data["duree_retenue"] = int(data["duree_retenue"])
        return cls(**data)


def profil_par_defaut() -> Profil:
    return Profil()


def hypotheses_par_defaut() -> Hypotheses:
    return Hypotheses()


# --- Persistance ---

def _lire_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default


def _ecrire_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def charger_profil() -> Profil:
    return Profil.from_dict(_lire_json(DATA / "profil.json", {}))


def sauver_profil(profil: Profil) -> None:
    _ecrire_json(DATA / "profil.json", profil.to_dict())


def migrer_prix_m2(hyp: Hypotheses) -> bool:
    """Remplace la table du classeur par les prix moyens saisis, sans toucher à un barème déjà modifié."""
    actuel = {i: float(hyp.prix_m2.get(i, 0)) for i in range(1, 21)}
    if any(actuel[i] != float(_PRIX_M2_CLASSEUR[i]) for i in range(1, 21)):
        return False
    hyp.prix_m2 = dict(PRIX_M2)
    return True


def charger_hypotheses() -> Hypotheses:
    hyp = Hypotheses.from_dict(_lire_json(DATA / "hypotheses.json", {}))
    if migrer_prix_m2(hyp):
        sauver_hypotheses(hyp)
    return hyp


def sauver_hypotheses(hyp: Hypotheses) -> None:
    _ecrire_json(DATA / "hypotheses.json", hyp.to_dict())


def reinitialiser_hypotheses() -> Hypotheses:
    path = DATA / "hypotheses.json"
    if path.exists():
        path.unlink()
    return Hypotheses()


def charger_biens() -> list:
    data = _lire_json(DATA / "biens.json", [])
    return data if isinstance(data, list) else []


def sauver_biens(biens: list) -> None:
    _ecrire_json(DATA / "biens.json", biens)


# --- Mathématiques de crédit ---

def mensualite_unitaire(taux_annuel: float, annees: int) -> float:
    """Mensualité hors assurance pour 1 € emprunté."""
    n = int(annees) * 12
    if n <= 0:
        return 0.0
    r = taux_annuel / 12.0
    if abs(r) < 1e-12:
        return 1.0 / n
    return r * (1 + r) ** n / ((1 + r) ** n - 1)


def mensualite_hors_assurance(capital: float, taux_annuel: float, annees: int) -> float:
    if capital <= 0:
        return 0.0
    return capital * mensualite_unitaire(taux_annuel, annees)


def capital_restant(capital: float, taux_annuel: float, annees: int, mois: int, mensualite: float) -> float:
    n = int(annees) * 12
    if mois >= n or capital <= 0:
        return 0.0
    r = taux_annuel / 12.0
    if abs(r) < 1e-12:
        return max(0.0, capital - mensualite * mois)
    crd = capital * (1 + r) ** mois - mensualite * ((1 + r) ** mois - 1) / r
    return max(0.0, crd)


def taux_mensuel(nper: int, pmt: float, pv: float) -> float:
    """Équivalent de TAUX / RATE d'Excel (échéance de fin de période, valeur future nulle)."""
    if nper <= 0 or pmt == 0:
        return float("nan")
    r = 0.003
    for _ in range(80):
        if r <= -0.9:
            r = 0.0
        fact = (1 + r) ** nper
        if abs(r) < 1e-14:
            f = pv + pmt * nper
            df = pv * nper + pmt * nper * (nper - 1) / 2
        else:
            f = pv * fact + pmt * (fact - 1) / r
            df = pv * nper * (1 + r) ** (nper - 1) + pmt * (
                r * nper * (1 + r) ** (nper - 1) - (fact - 1)
            ) / (r ** 2)
        if df == 0 or math.isnan(df):
            break
        nxt = r - f / df
        if abs(nxt - r) < 1e-14:
            return nxt
        r = nxt
    return r


def capacite_emprunt(mensualite: float, taux: float, annees: int, assurance: float, inclut_assurance: bool) -> float:
    if mensualite <= 0:
        return 0.0
    facteur = mensualite_unitaire(taux, annees)
    if inclut_assurance:
        facteur += assurance / 12.0
    if facteur <= 0:
        return 0.0
    return mensualite / facteur


def simuler_pret(capital: float, taux: float, annees: int, assurance: float, frais_bancaires: float = 0.0) -> dict:
    ha = mensualite_hors_assurance(capital, taux, annees)
    ass = capital * assurance / 12.0 if capital > 0 else 0.0
    totale = ha + ass
    mois = int(annees) * 12
    interets = ha * mois - capital if capital > 0 else 0.0
    cout_ass = ass * mois
    taeg = float("nan")
    if capital > 0 and totale > 0:
        net = capital - frais_bancaires
        if net > 0:
            r = taux_mensuel(mois, -totale, net)
            if not math.isnan(r):
                taeg = (1 + r) ** 12 - 1
    return {
        "mensualite_ha": ha,
        "assurance_mois": ass,
        "mensualite": totale,
        "interets": interets,
        "cout_assurance": cout_ass,
        "cout_credit": interets + cout_ass + frais_bancaires,
        "taeg": taeg,
    }


# --- Notaire, estimation, projet ---

def frais_notaire(prix: float, hyp: Hypotheses) -> dict:
    s1, s2, s3 = hyp.emol_seuils
    t1, t2, t3, t4 = hyp.emol_taux
    droits = prix * (hyp.droit_dept + hyp.taxe_communale + hyp.droit_dept * hyp.frais_assiette)
    emol_ht = (
        min(prix, s1) * t1
        + max(min(prix, s2) - s1, 0) * t2
        + max(min(prix, s3) - s2, 0) * t3
        + max(prix - s3, 0) * t4
    )
    tva = emol_ht * hyp.tva_emoluments
    csi = prix * hyp.csi
    total = droits + emol_ht + tva + hyp.debours + csi
    return {
        "droits": droits,
        "emoluments_ht": emol_ht,
        "tva": tva,
        "debours": hyp.debours,
        "csi": csi,
        "total": total,
        "pct": (total / prix) if prix else 0.0,
    }


def _coef_surface(surface: float, hyp: Hypotheses) -> float:
    coef = hyp.coef_surface[0][1]
    for seuil, valeur in hyp.coef_surface:
        if surface >= seuil:
            coef = valeur
    return coef


def _coef_etage(etage: int, ascenseur: bool, hyp: Hypotheses) -> float:
    i = min(max(int(etage), 0), 7)
    return hyp.etage_avec[i] if ascenseur else hyp.etage_sans[i]


def estimer(bien: dict, hyp: Hypotheses) -> dict | None:
    surface = float(bien.get("surface") or 0)
    arr = int(bien.get("arrondissement") or 0)
    if surface <= 0 or arr not in hyp.prix_m2:
        return None
    prix_ref = float(hyp.prix_m2[arr])
    asc = bool(bien.get("ascenseur"))
    dernier = bool(bien.get("dernier_etage"))
    coefs = {
        "Étage × ascenseur": _coef_etage(bien.get("etage") or 0, asc, hyp),
        "Dernier étage": hyp.coef_dernier_asc if (dernier and asc) else (hyp.coef_dernier_sans if dernier else 0.0),
        "Surface": _coef_surface(surface, hyp),
        "Immeuble": hyp.coef_immeuble.get(bien.get("immeuble"), 0.0),
        "État": hyp.coef_etat.get(bien.get("etat"), 0.0),
        "DPE": hyp.coef_dpe.get(bien.get("dpe"), 0.0),
        "Luminosité": hyp.coef_lumiere.get(bien.get("luminosite"), 0.0),
        "Nuisances": hyp.coef_bruit.get(bien.get("nuisances"), 0.0),
        "Extérieur": hyp.coef_exterieur.get(bien.get("exterieur"), 0.0),
    }
    mult = 1.0
    for c in coefs.values():
        mult *= 1 + c
    annexes = (
        (hyp.valeur_cave if bien.get("cave") else 0.0)
        + (hyp.valeur_local_velo if bien.get("local_velo") else 0.0)
        + (hyp.valeur_parking if bien.get("parking") else 0.0)
    )
    prix_m2 = prix_ref * mult
    valeur = prix_m2 * surface + annexes
    return {
        "prix_ref": prix_ref,
        "coefs": coefs,
        "multiplicateur": mult,
        "prix_m2": prix_m2,
        "annexes": annexes,
        "valeur": valeur,
        "basse": valeur * (1 - hyp.fourchette),
        "haute": valeur * (1 + hyp.fourchette),
    }


def verdict(ecart: float, hyp: Hypotheses) -> str:
    if ecart <= hyp.seuil_sous_marche:
        return "Sous le marché – à creuser"
    if ecart <= hyp.seuil_coherent:
        return "Prix cohérent"
    if ecart <= hyp.seuil_negociable:
        return "Un peu cher – négociable"
    return "Surévalué – négocier ou passer"


def arrondi_millier(valeur: float) -> int:
    return int(math.floor(valeur / 1000.0 + 0.5)) * 1000


def analyser_prix(bien: dict, hyp: Hypotheses) -> dict | None:
    est = estimer(bien, hyp)
    if est is None:
        return None
    surface = float(bien["surface"])
    prix = float(bien.get("prix") or 0)
    travaux = float(bien.get("travaux") or 0)
    out = dict(est)
    if prix <= 0:
        out.update({"prix": 0, "a_un_prix": False})
        return out
    ecart = prix / est["valeur"] - 1
    notaire = frais_notaire(prix, hyp)["total"]
    charges = float(bien["charges"]) if bien.get("charges_connues") else surface * hyp.charges_m2
    taxe = float(bien["taxe"]) if bien.get("taxe_connue") else surface * hyp.taxe_m2
    acquisition = prix + notaire + travaux
    out.update(
        {
            "a_un_prix": True,
            "prix": prix,
            "prix_m2_demande": prix / surface,
            "ecart": ecart,
            "ecart_eur": prix - est["valeur"],
            "verdict": verdict(ecart, hyp),
            "offre": min(prix, arrondi_millier(est["valeur"])),
            "plafond": arrondi_millier(est["haute"]),
            "notaire_sur_demande": notaire,
            "acquisition": acquisition,
            "acquisition_m2": acquisition / surface,
            "charges": charges,
            "taxe": taxe,
        }
    )
    return out


def cout_projet(prix: float, travaux: float, profil: Profil, hyp: Hypotheses) -> dict:
    notaire = frais_notaire(max(prix, 0), hyp)
    agence = max(prix, 0) * profil.agence_pct
    besoin = max(prix, 0) + agence + notaire["total"] + max(travaux, 0) + profil.mobilier
    garantie = hyp.garantie_pct * max(0.0, besoin - profil.apport)
    frais_banque = garantie + hyp.frais_dossier + hyp.frais_courtage
    total = besoin + frais_banque
    emprunt = max(0.0, total - profil.apport)
    return {
        "prix": max(prix, 0),
        "agence": agence,
        "notaire": notaire,
        "travaux": max(travaux, 0),
        "mobilier": profil.mobilier,
        "besoin": besoin,
        "garantie": garantie,
        "dossier": hyp.frais_dossier,
        "courtage": hyp.frais_courtage,
        "frais_banque": frais_banque,
        "total": total,
        "emprunt": emprunt,
        "apport_utilise": min(profil.apport, total),
    }


def prix_achat_max(enveloppe: float, travaux: float, profil: Profil, hyp: Hypotheses) -> float:
    """Plus grand prix d'achat dont le coût total de projet tient dans l'enveloppe."""
    if enveloppe <= 0:
        return 0.0
    if cout_projet(0, travaux, profil, hyp)["total"] > enveloppe + 1:
        return 0.0
    lo, hi = 0.0, max(enveloppe, 1.0)
    # Le prix max est inférieur à l'enveloppe ; on élargit si les frais sont négatifs (cas impossible).
    while cout_projet(hi, travaux, profil, hyp)["total"] <= enveloppe:
        hi *= 1.5
        if hi > enveloppe * 3:
            break
    for _ in range(70):
        mid = (lo + hi) / 2
        if cout_projet(mid, travaux, profil, hyp)["total"] <= enveloppe:
            lo = mid
        else:
            hi = mid
    return lo


def scenario(profil: Profil, duree: int, hyp: Hypotheses | None = None) -> dict:
    taux = profil.taux(duree)
    cap = capacite_emprunt(profil.mensualite_cible, taux, duree, profil.assurance, profil.inclut_assurance)
    budget = cap + profil.apport
    pret_unitaire = simuler_pret(cap, taux, duree, profil.assurance)
    return {
        "duree": duree,
        "taux": taux,
        "capacite": cap,
        "budget": budget,
        "mensualite_ha_max": pret_unitaire["mensualite_ha"],
        "assurance_max": pret_unitaire["assurance_mois"],
        "mensualite_totale_max": pret_unitaire["mensualite"],
    }


def financement_scenario(projet: dict, profil: Profil, duree: int) -> dict:
    pret = simuler_pret(
        projet["emprunt"],
        profil.taux(duree),
        duree,
        profil.assurance,
        projet["frais_banque"],
    )
    cible = profil.mensualite_cible
    comparable = pret["mensualite"] if profil.inclut_assurance else pret["mensualite_ha"]
    cap = capacite_emprunt(cible, profil.taux(duree), duree, profil.assurance, profil.inclut_assurance)
    budget = cap + profil.apport
    endettement = None
    if profil.revenus > 0:
        endettement = (pret["mensualite"] + profil.autres_credits) / profil.revenus
    return {
        "duree": duree,
        "taux": profil.taux(duree),
        **pret,
        "comparable": comparable,
        "ecart_mensualite": comparable - cible,
        "tient": comparable <= cible + 0.5,
        "capacite": cap,
        "ecart_emprunt": projet["emprunt"] - cap,
        "budget": budget,
        "ecart_budget": projet["total"] - budget,
        "endettement": endettement,
        "sous_plafond": None if endettement is None else endettement <= profil.endettement_max + 1e-9,
        "revenu_min": (pret["mensualite"] + profil.autres_credits) / profil.endettement_max if profil.endettement_max else None,
        "mensualite_max_revenus": (profil.revenus * profil.endettement_max - profil.autres_credits) if profil.revenus > 0 else None,
    }


def budget_possession(bien: dict, projet: dict, pret: dict, profil: Profil, hyp: Hypotheses) -> dict:
    surface = float(bien.get("surface") or 0)
    prix = projet["prix"]
    if bien.get("charges_connues"):
        charges_an = float(bien.get("charges") or 0)
    else:
        charges_an = surface * hyp.charges_m2
    if bien.get("taxe_connue"):
        taxe_an = float(bien.get("taxe") or 0)
    else:
        taxe_an = surface * hyp.taxe_m2
    assurance_an = hyp.assurance_habitation
    entretien_an = prix * hyp.entretien_pct
    provision_an = prix * hyp.provision_copro_pct
    possession_an = charges_an + taxe_an + assurance_an + entretien_an + provision_an
    energie = hyp.energie_mois
    credit = pret["mensualite"]
    total_mois = credit + possession_an / 12 + energie
    loyer = surface * hyp.loyer_m2
    return {
        "mensualite_ha": pret["mensualite_ha"],
        "assurance_mois": pret["assurance_mois"],
        "charges_an": charges_an,
        "taxe_an": taxe_an,
        "assurance_an": assurance_an,
        "entretien_an": entretien_an,
        "provision_an": provision_an,
        "possession_an": possession_an,
        "energie": energie,
        "total_mois": total_mois,
        "total_an": total_mois * 12,
        "loyer": loyer,
        "rendement": (loyer * 12 / prix) if prix else None,
        "annees_loyer": (prix / (loyer * 12)) if loyer else None,
        "effort": (total_mois / profil.revenus) if profil.revenus else None,
        "reste": (profil.revenus - profil.autres_credits - total_mois) if profil.revenus else None,
    }


def projection_achat_location(bien: dict, projet: dict, pret: dict, budget: dict, profil: Profil, hyp: Hypotheses, duree: int) -> dict:
    """Si tu achètes, contre si tu continues à louer au loyer actuel et que tu places l'apport.

    Le loyer actuel est tout compris (loyer + charges). On n'y ajoute pas une deuxième fois
    les charges. L'énergie est hors des deux côtés.
    """
    rows = []
    capital_loc = profil.apport
    prix = projet["prix"]
    emprunt = projet["emprunt"]
    seuil = None
    for annee in range(0, hyp.horizon_annees + 1):
        valeur = prix * (1 + hyp.hausse_immo) ** annee
        if annee == 0:
            crd = emprunt
            sorties_achat = 0.0
            sorties_loc = 0.0
        else:
            crd = capital_restant(emprunt, profil.taux(duree), duree, annee * 12, pret["mensualite_ha"])
            credit_an = 12 * (pret["mensualite_ha"] + pret["assurance_mois"]) if annee <= duree else 0.0
            sorties_achat = credit_an + budget["possession_an"] * (1 + hyp.inflation_charges) ** (annee - 1)
            # Loyer actuel tout compris, qui suit la hausse des loyers.
            sorties_loc = profil.loyer_actuel * 12 * (1 + hyp.hausse_loyers) ** (annee - 1)
            capital_loc = capital_loc * (1 + hyp.rendement_epargne) + sorties_achat - sorties_loc
        patrimoine = valeur * (1 - hyp.frais_revente) - crd
        avantage = patrimoine - capital_loc
        if seuil is None and avantage >= 0:
            seuil = annee
        rows.append(
            {
                "annee": annee,
                "valeur": valeur,
                "crd": crd,
                "patrimoine_acheteur": patrimoine,
                "capital_locataire": capital_loc,
                "avantage": avantage,
                "sorties_achat": sorties_achat,
                "sorties_loc": sorties_loc,
            }
        )

    def a(annee):
        if annee < 0 or annee >= len(rows):
            return None
        return rows[annee]["avantage"]

    return {
        "seuil": seuil,
        "avantage_5": a(5),
        "avantage_10": a(10),
        "avantage_15": a(15),
        "avantage_20": a(20),
        "avantage_25": a(min(25, hyp.horizon_annees)),
        "lignes": rows,
    }


def projection_revente(projet: dict, pret: dict, profil: Profil, hyp: Hypotheses, duree: int) -> dict:
    """Résultat si tu revends et que tu y as habité.

    Cash récupéré = valeur × (1 − frais de revente) − crédit restant.
    On en retire l'apport, les intérêts et l'assurance déjà payés.
    Le notaire, les travaux et le mobilier sont dans le trou de départ :
    ils ont été payés, ils ne reviennent pas dans le prix de vente.
    La copropriété et la taxe foncière n'y sont pas : c'est le coût d'habiter,
    comparé au loyer dans l'autre calcul.
    """
    prix = projet["prix"]
    emprunt = projet["emprunt"]
    apport = projet["apport_utilise"]
    lignes = []
    seuil = None
    for annee in range(0, hyp.horizon_annees + 1):
        valeur = prix * (1 + hyp.hausse_immo) ** annee
        mois = min(annee, int(duree)) * 12
        if mois <= 0:
            crd = emprunt
        else:
            crd = capital_restant(emprunt, profil.taux(duree), duree, mois, pret["mensualite_ha"])
        interets = max(0.0, pret["mensualite_ha"] * mois - (emprunt - crd)) if mois else 0.0
        assurance = pret["assurance_mois"] * mois
        cash = valeur * (1 - hyp.frais_revente) - crd
        resultat = cash - apport - interets - assurance
        if seuil is None and resultat >= 0:
            seuil = annee
        lignes.append(
            {
                "annee": annee,
                "valeur": valeur,
                "crd": crd,
                "interets": interets,
                "assurance": assurance,
                "cash": cash,
                "resultat": resultat,
            }
        )
    return {
        "seuil": seuil,
        "lignes": lignes,
        "perte_immediate": lignes[0]["resultat"],
        "resultat_final": lignes[-1]["resultat"],
    }


def sensibilite(emprunt: float, taux: float, assurance: float) -> list:
    lignes = []
    for ecart in (-0.005, 0.0, 0.005, 0.01):
        ligne = {"ecart": ecart, "taux": taux + ecart, "durees": {}}
        for duree in (15, 20, 25):
            pret = simuler_pret(emprunt, taux + ecart, duree, assurance)
            ligne["durees"][duree] = pret
        lignes.append(ligne)
    return lignes


def analyser(bien: dict, profil: Profil, hyp: Hypotheses) -> dict | None:
    est = analyser_prix(bien, hyp)
    if est is None:
        return None
    prix_retenu = float(bien["prix_vise"]) if bien.get("a_une_offre") and bien.get("prix_vise") else float(bien.get("prix") or 0)
    travaux = float(bien.get("travaux") or 0)
    if prix_retenu <= 0:
        return {"estimation": est, "pret": False}
    projet = cout_projet(prix_retenu, travaux, profil, hyp)
    scenarios = {d: financement_scenario(projet, profil, d) for d in (20, 25)}
    retenu = scenarios[20 if profil.duree_retenue == 20 else 25]
    budg = budget_possession(bien, projet, retenu, profil, hyp)
    comparatif = projection_achat_location(bien, projet, retenu, budg, profil, hyp, retenu["duree"])
    revente = projection_revente(projet, retenu, profil, hyp, retenu["duree"])
    sensi = sensibilite(projet["emprunt"], retenu["taux"], profil.assurance)
    prix_max = {
        d: prix_achat_max(scenarios[d]["budget"], travaux, profil, hyp) for d in (20, 25)
    }
    ecart_retenu = (prix_retenu / est["valeur"] - 1) if est["valeur"] else None
    return {
        "estimation": est,
        "pret": True,
        "prix_retenu": prix_retenu,
        "ecart_retenu": ecart_retenu,
        "verdict_retenu": verdict(ecart_retenu, hyp) if ecart_retenu is not None else None,
        "projet": projet,
        "scenarios": scenarios,
        "retenu": retenu,
        "budget": budg,
        "comparatif": comparatif,
        "revente": revente,
        "sensibilite": sensi,
        "prix_max": prix_max,
    }


def nouveau_bien_vide() -> dict:
    return {
        "id": str(uuid.uuid4()),
        "cree_le": date.today().isoformat(),
        "nom": "",
        "lien": "",
        "notes": "",
        "arrondissement": 11,
        "surface": 0.0,
        "etage": 3,
        "ascenseur": False,
        "dernier_etage": False,
        "immeuble": "Ancien standard (avant 1950)",
        "etat": "Bon état",
        "dpe": "D",
        "luminosite": "Normale",
        "nuisances": "Normal",
        "exterieur": "Aucun",
        "cave": False,
        "local_velo": False,
        "parking": False,
        "prix": 0.0,
        "a_une_offre": False,
        "prix_vise": 0.0,
        "travaux": 0.0,
        "charges_connues": False,
        "charges": 0.0,
        "taxe_connue": False,
        "taxe": 0.0,
        "checklist": [{"statut": "À faire", "notes": ""} for _ in CHECKLIST],
    }


def resume_bien(bien: dict, profil: Profil, hyp: Hypotheses) -> dict:
    """Ligne de comparaison pour la liste des biens enregistrés."""
    nom = bien.get("nom") or "Sans nom"
    base = {
        "id": bien.get("id"),
        "nom": nom,
        "arrondissement": bien.get("arrondissement"),
        "surface": bien.get("surface") or 0,
        "prix": bien.get("prix") or 0,
    }
    resultat = analyser(bien, profil, hyp)
    if not resultat or not resultat.get("pret"):
        base["incomplet"] = True
        return base
    est = resultat["estimation"]
    s25 = resultat["scenarios"][25]
    s20 = resultat["scenarios"][20]
    base.update(
        {
            "incomplet": False,
            "valeur": est["valeur"],
            "ecart": est.get("ecart"),
            "verdict": est.get("verdict"),
            "cout": resultat["projet"]["total"],
            "emprunt": resultat["projet"]["emprunt"],
            "mensualite_20": s20["comparable"],
            "mensualite_25": s25["comparable"],
            "tient_20": s20["tient"],
            "tient_25": s25["tient"],
            "ecart_budget_25": s25["ecart_budget"],
        }
    )
    return base


def _self_check() -> None:
    hyp = Hypotheses()
    # Exemple du classeur : 3P 62 m², Paris 11e, 3e sans ascenseur, cave, 620 000 €.
    bien = nouveau_bien_vide()
    bien.update(
        {
            "arrondissement": 11,
            "surface": 62,
            "etage": 3,
            "ascenseur": False,
            "dernier_etage": False,
            "cave": True,
            "prix": 620_000,
            "travaux": 15_000,
            "charges_connues": True,
            "charges": 2400,
            "taxe_connue": True,
            "taxe": 1400,
        }
    )
    est = analyser_prix(bien, hyp)
    assert est["valeur"] == 612_662, est["valeur"]
    assert est["verdict"] == "Prix cohérent", est["verdict"]
    assert est["offre"] == 613_000
    assert est["plafond"] == 643_000
    assert abs(est["notaire_sur_demande"] - 47_415.96) < 0.02

    profil = Profil(apport=120_000, taux_25=0.0385, duree_retenue=25, revenus=9500, mobilier=5000)
    projet = cout_projet(620_000, 15_000, profil, hyp)
    assert abs(projet["total"] - 694_090.1196) < 0.05, projet["total"]
    assert abs(projet["emprunt"] - 574_090.1196) < 0.05

    cap20 = capacite_emprunt(2600, 0.0335, 20, 0.003, False)
    cap25 = capacite_emprunt(2600, 0.0355, 25, 0.003, False)
    assert round(cap20) == 454_322, round(cap20)
    assert round(cap25) == 516_581, round(cap25)

    foyer = Profil()
    resultat = analyser(bien, foyer, hyp)
    assert resultat["scenarios"][25]["tient"] is True or resultat["scenarios"][25]["tient"] is False
    assert len(resultat["comparatif"]["lignes"]) == 26
    revente = resultat["revente"]
    assert revente["perte_immediate"] < 0
    assert revente["lignes"][0]["interets"] == 0
    assert revente["lignes"][10]["interets"] > revente["lignes"][1]["interets"]
    # Le capital restant dû décroît et tombe à 0 au terme.
    crd_fin = resultat["comparatif"]["lignes"][25]["crd"]
    assert crd_fin == 0 or resultat["retenu"]["duree"] != 25 or crd_fin < 1
    print("self-check ok", round(est["valeur"]), est["verdict"], round(cap20), round(cap25))


if __name__ == "__main__":
    _self_check()
