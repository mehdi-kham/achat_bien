"""Achat de la résidence principale — saisie d'un bien, calculs du classeur Excel."""

from __future__ import annotations

import html
import math

import pandas as pd
import streamlit as st

from moteur import (
    CHECKLIST,
    COEF_BRUIT,
    COEF_DPE,
    COEF_ETAT,
    COEF_EXTERIEUR,
    COEF_IMMEUBLE,
    COEF_LUMIERE,
    ETAGE_LABELS,
    QUARTIERS,
    Hypotheses,
    Profil,
    analyser,
    arrondi_millier,
    capacite_emprunt,
    charger_biens,
    charger_hypotheses,
    charger_profil,
    cout_projet,
    frais_notaire,
    migrer_prix_m2,
    nouveau_bien_vide,
    prix_achat_max,
    projection_achat_location,
    reinitialiser_hypotheses,
    resume_bien,
    sauver_biens,
    sauver_hypotheses,
    sauver_profil,
    scenario,
)

st.set_page_config(
    page_title="Résidence principale",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1180px; }
    div[data-testid="stMetric"] {
        background: #fffcf8;
        border: 1px solid #e6dfd4;
        border-radius: 12px;
        padding: 0.55rem 0.75rem 0.35rem 0.75rem;
    }
    div[data-testid="stMetric"] label { color: #57534e; }
    h1 { letter-spacing: -0.03em; font-weight: 650; }
    h2, h3 { letter-spacing: -0.02em; }
    </style>
    """,
    unsafe_allow_html=True,
)

CHAMPS_BIEN = [
    "nom",
    "lien",
    "notes",
    "arrondissement",
    "surface",
    "etage",
    "ascenseur",
    "dernier_etage",
    "immeuble",
    "etat",
    "dpe",
    "luminosite",
    "nuisances",
    "exterieur",
    "cave",
    "local_velo",
    "parking",
    "prix",
    "a_une_offre",
    "prix_vise",
    "travaux",
    "charges_connues",
    "charges",
    "taxe_connue",
    "taxe",
]
PAGES = ["Capacité d'achat", "Analyser un bien", "Biens enregistrés", "Hypothèses"]
STATUTS = ["À faire", "OK", "Alerte", "N/A"]


def eur(valeur, decimales=0) -> str:
    if valeur is None or (isinstance(valeur, float) and (math.isnan(valeur) or math.isinf(valeur))):
        return "—"
    signe = "−" if valeur < 0 else ""
    nombre = abs(float(valeur))
    texte = f"{nombre:,.{decimales}f}"
    texte = texte.replace(",", " ").replace(".", ",")
    return f"{signe}{texte} €"


def pct(valeur, decimales=1) -> str:
    if valeur is None or (isinstance(valeur, float) and (math.isnan(valeur) or math.isinf(valeur))):
        return "—"
    signe = "−" if valeur < 0 else ""
    nombre = abs(float(valeur)) * 100
    texte = f"{nombre:,.{decimales}f}".replace(",", " ").replace(".", ",")
    return f"{signe}{texte} %"


def ordinal(arrondissement: int) -> str:
    return "1er" if int(arrondissement) == 1 else f"{int(arrondissement)}e"


def tableau(entetes: list[str], lignes: list[list[str]]) -> None:
    head = "".join(
        f"<th style='text-align:left;padding:0.55rem 0.7rem;border-bottom:1px solid #e6dfd4;'>{html.escape(h)}</th>"
        for h in entetes
    )
    corps = []
    for ligne in lignes:
        cellules = "".join(
            f"<td style='padding:0.55rem 0.7rem;border-bottom:1px solid #f3eee6;white-space:nowrap;'>{html.escape(c)}</td>"
            for c in ligne
        )
        corps.append(f"<tr>{cellules}</tr>")
    st.markdown(
        "<div style='overflow-x:auto;border:1px solid #e6dfd4;border-radius:12px;background:#fffcf8;margin:0.2rem 0 0.8rem 0;'>"
        "<table style='border-collapse:collapse;width:100%;font-size:0.95rem;'>"
        f"<thead><tr>{head}</tr></thead><tbody>{''.join(corps)}</tbody></table></div>",
        unsafe_allow_html=True,
    )


def bandeau(genre: str, titre: str, detail: str = "") -> None:
    styles = {
        "ok": "background:#ecfdf3;color:#14532d;border:1px solid #bbf7d0;",
        "warn": "background:#fffbeb;color:#78350f;border:1px solid #fde68a;",
        "bad": "background:#fef2f2;color:#7f1d1d;border:1px solid #fecaca;",
        "info": "background:#f8f5f0;color:#292524;border:1px solid #e7e0d6;",
    }
    contenu = (
        f"<div style='{styles[genre]}border-radius:12px;padding:0.8rem 1rem;margin:0.2rem 0 0.9rem 0;'>"
        f"<strong>{html.escape(titre)}</strong>"
    )
    if detail:
        contenu += f"<div style='font-weight:500;margin-top:0.25rem;line-height:1.45;'>{html.escape(detail)}</div>"
    contenu += "</div>"
    st.markdown(contenu, unsafe_allow_html=True)


def disclaimer() -> None:
    st.caption(
        "Aide à la décision, pas un conseil financier, fiscal ou juridique. "
        "Frais de notaire, taux et situation de la copropriété sont à faire confirmer par un notaire et un courtier. "
        "Prix au m² : moyennes par arrondissement, pas le prix de la rue. "
        "Les taux du foyer sont ceux d'octobre 2026 (3,35 % sur 20 ans, 3,55 % sur 25 ans). "
        "Le classeur Excel, lui, était calé sur les moyennes courtiers de novembre (3,76 % et 3,85 %)."
    )


def normaliser_checklist(saved) -> list:
    saved = saved or []
    lignes = []
    for i in range(len(CHECKLIST)):
        if i < len(saved) and isinstance(saved[i], dict):
            statut = saved[i].get("statut") or "À faire"
            if statut not in STATUTS:
                statut = "À faire"
            lignes.append({"statut": statut, "notes": saved[i].get("notes") or ""})
        else:
            lignes.append({"statut": "À faire", "notes": ""})
    return lignes


def cadre_checklist(lignes) -> pd.DataFrame:
    rows = []
    for (theme, point, pourquoi), reponse in zip(CHECKLIST, normaliser_checklist(lignes)):
        rows.append(
            {
                "Thème": theme,
                "Point": point,
                "Pourquoi": pourquoi,
                "Statut": reponse["statut"],
                "Notes": reponse["notes"],
            }
        )
    return pd.DataFrame(rows)


FLOTTANTS_BIEN = {"surface", "prix", "prix_vise", "travaux", "charges", "taxe"}
ENTIERS_BIEN = {"arrondissement", "etage"}
BOOLEENS_BIEN = {
    "ascenseur",
    "dernier_etage",
    "cave",
    "local_velo",
    "parking",
    "a_une_offre",
    "charges_connues",
    "taxe_connue",
}


def formulaire_depuis(bien: dict) -> dict:
    vide = nouveau_bien_vide()
    formulaire = {}
    for cle in CHAMPS_BIEN:
        valeur = bien.get(cle, vide[cle])
        if cle in FLOTTANTS_BIEN:
            valeur = float(valeur or 0)
        elif cle in ENTIERS_BIEN:
            valeur = int(valeur or 0)
        elif cle in BOOLEENS_BIEN:
            valeur = bool(valeur)
        elif valeur is None:
            valeur = ""
        formulaire[cle] = valeur
    return formulaire


def injecter_bien(bien: dict, nouveau: bool = False) -> None:
    st.session_state.current_id = None if nouveau else bien.get("id")
    st.session_state.current_cree_le = None if nouveau else bien.get("cree_le")
    formulaire = formulaire_depuis(bien)
    st.session_state.formulaire = formulaire
    # Posé dans le même tour que les widgets : Streamlit pousse ces valeurs à l'écran.
    for cle, valeur in formulaire.items():
        st.session_state[f"b_{cle}"] = valeur
    st.session_state.checklist_rows = normaliser_checklist(bien.get("checklist"))
    st.session_state.pop("editeur_checklist", None)


def profil_reference(loyer: float = 1_900.0) -> Profil:
    """Tableau de capacité à conserver : 20 ans retenu, 2 600 €, apport 320 000 €."""
    return Profil(duree_retenue=20, mobilier=0.0, loyer_actuel=loyer or 1_900.0)


def profil_a_ete_efface(profil: Profil) -> bool:
    return profil.apport < 1 and profil.mensualite_cible < 1 and profil.taux_20 < 1e-8 and profil.taux_25 < 1e-8


def tenir_profil() -> None:
    """Remet le tableau de capacité s'il a été écrasé par des champs encore vides."""
    profil = st.session_state.profil
    if not profil_a_ete_efface(profil):
        return
    profil = profil_reference(profil.loyer_actuel)
    st.session_state.profil = profil
    sauver_profil(profil)
    for cle in list(st.session_state.keys()):
        if str(cle).startswith("p_") and cle != "p_loyer_actuel":
            del st.session_state[cle]


def semer_profil(profil: Profil) -> None:
    """À appeler dans le même tour que les widgets, jamais au démarrage."""
    valeurs = {
        "p_apport": float(profil.apport),
        "p_mensualite": float(profil.mensualite_cible),
        "p_inclut": bool(profil.inclut_assurance),
        "p_taux20": float(profil.taux_20) * 100,
        "p_taux25": float(profil.taux_25) * 100,
        "p_duree": int(profil.duree_retenue),
        "p_revenus": float(profil.revenus),
        "p_autres": float(profil.autres_credits),
        "p_assurance": float(profil.assurance) * 100,
        "p_mobilier": float(profil.mobilier),
        "p_agence": float(profil.agence_pct) * 100,
        "p_endettement": float(profil.endettement_max) * 100,
    }
    for cle, valeur in valeurs.items():
        if cle not in st.session_state:
            st.session_state[cle] = valeur


def profil_courant() -> Profil:
    return Profil(
        apport=float(st.session_state.p_apport),
        mensualite_cible=float(st.session_state.p_mensualite),
        inclut_assurance=bool(st.session_state.p_inclut),
        taux_20=float(st.session_state.p_taux20) / 100,
        taux_25=float(st.session_state.p_taux25) / 100,
        duree_retenue=int(st.session_state.p_duree),
        revenus=float(st.session_state.p_revenus),
        autres_credits=float(st.session_state.p_autres),
        assurance=float(st.session_state.p_assurance) / 100,
        mobilier=float(st.session_state.p_mobilier),
        agence_pct=float(st.session_state.p_agence) / 100,
        endettement_max=float(st.session_state.p_endettement) / 100,
        loyer_actuel=float(
            st.session_state["p_loyer_actuel"]
            if "p_loyer_actuel" in st.session_state
            else getattr(st.session_state.get("profil"), "loyer_actuel", 1900)
        ),
    )


def lire_bien() -> dict:
    bien = dict(st.session_state.formulaire)
    bien["arrondissement"] = int(bien["arrondissement"])
    bien["etage"] = int(bien["etage"])
    bien["checklist"] = st.session_state.get("checklist_rows") or normaliser_checklist([])
    bien["id"] = st.session_state.get("current_id")
    bien["cree_le"] = st.session_state.get("current_cree_le")
    return bien


def init() -> None:
    if not st.session_state.get("ready"):
        profil = charger_profil()
        st.session_state.profil = profil
        st.session_state.hyp = charger_hypotheses()
        st.session_state.biens = charger_biens()
        st.session_state.nav = "Capacité d'achat"
        st.session_state.hyp_nonce = 0
        st.session_state.formulaire = formulaire_depuis(nouveau_bien_vide())
        st.session_state.checklist_rows = normaliser_checklist([])
        st.session_state.current_id = None
        st.session_state.current_cree_le = None
        st.session_state.ready = True
        st.session_state.form_version = 2
    tenir_profil()
    if migrer_prix_m2(st.session_state.hyp):
        sauver_hypotheses(st.session_state.hyp)
        st.session_state.hyp_nonce = st.session_state.get("hyp_nonce", 0) + 1
        for cle in list(st.session_state.keys()):
            if str(cle).startswith("ed_prix_"):
                del st.session_state[cle]
    if st.session_state.get("form_version") != 2:
        # Les anciennes clés de formulaire avaient été posées avant l'affichage
        # des widgets : Streamlit les ignorait et montrait la première option.
        for cle in list(st.session_state.keys()):
            if str(cle).startswith("b_"):
                del st.session_state[cle]
        if "formulaire" not in st.session_state:
            st.session_state.formulaire = formulaire_depuis(nouveau_bien_vide())
        st.session_state.form_version = 2


def bien_exemple() -> dict:
    exemple = nouveau_bien_vide()
    exemple.update(
        {
            "nom": "Exemple — 3 pièces, Paris 11e",
            "arrondissement": 11,
            "surface": 62,
            "etage": 3,
            "ascenseur": False,
            "dernier_etage": False,
            "cave": True,
            "parking": False,
            "prix": 620_000,
            "travaux": 15_000,
            "charges_connues": True,
            "charges": 2_400,
            "taxe_connue": True,
            "taxe": 1_400,
            "notes": "Exemple du fichier Excel, à remplacer par un vrai bien.",
        }
    )
    return exemple


def preparer() -> None:
    if "formulaire" not in st.session_state:
        st.session_state.formulaire = formulaire_depuis(nouveau_bien_vide())
    if "checklist_rows" not in st.session_state:
        st.session_state.checklist_rows = normaliser_checklist([])
    if st.session_state.pop("pending_new", False):
        injecter_bien(nouveau_bien_vide(), nouveau=True)
    if st.session_state.pop("pending_exemple", False):
        injecter_bien(bien_exemple(), nouveau=True)
    bien_id = st.session_state.pop("pending_bien_id", None)
    if bien_id:
        trouve = next((b for b in st.session_state.biens if b.get("id") == bien_id), None)
        if trouve:
            injecter_bien(trouve, nouveau=False)
    if st.session_state.get("goto"):
        st.session_state.nav = st.session_state.pop("goto")
    vide = formulaire_depuis(nouveau_bien_vide())
    for cle, valeur in vide.items():
        st.session_state.formulaire.setdefault(cle, valeur)
    if "p_apport" in st.session_state and "p_duree" in st.session_state:
        st.session_state.profil = profil_courant()
        sauver_profil(st.session_state.profil)


def barre_laterale() -> None:
    profil = st.session_state.profil
    st.sidebar.markdown("### Résidence principale")
    st.sidebar.caption("Paris · foyer")
    st.sidebar.markdown(f"**Apport**  \n{eur(profil.apport)}")
    libelle_cible = "mensualité tout compris" if profil.inclut_assurance else "hors assurance"
    st.sidebar.markdown(f"**Cible**  \n{eur(profil.mensualite_cible)} / mois ({libelle_cible})")
    for duree in (20, 25):
        s = scenario(profil, duree)
        st.sidebar.metric(f"Budget {duree} ans", eur(s["budget"]))
        st.sidebar.caption(f"{pct(s['taux'], 2)} · emprunt {eur(s['capacite'])}")
    st.sidebar.divider()
    st.sidebar.radio("Aller à", PAGES, key="nav", label_visibility="collapsed")
    n = len(st.session_state.biens)
    st.sidebar.caption(f"{n} bien{'s' if n > 1 else ''} enregistré{'s' if n > 1 else ''}".replace("0 bien enregistré", "Aucun bien enregistré"))


def profil_est_le_repere(profil: Profil) -> bool:
    return (
        abs(profil.apport - 320_000) < 1
        and abs(profil.mensualite_cible - 2_600) < 1
        and abs(profil.taux_20 - 0.0335) < 1e-6
        and abs(profil.taux_25 - 0.0355) < 1e-6
        and not profil.inclut_assurance
    )


def page_capacite() -> None:
    profil = st.session_state.profil
    hyp = st.session_state.hyp
    st.title("Capacité d'achat")
    st.markdown(
        "Deux durées pour la même mensualité. "
        "Le **budget global** est l'enveloppe du projet entier : prix, frais de notaire, travaux, mobilier et frais de prêt. "
        "Le **prix d'achat maximum** est ce qui reste pour le bien une fois ces frais retirés."
    )
    st.caption("Ce tableau reste fixe. Il ne change que si vous modifiez les paramètres du foyer, plus bas.")
    if profil_est_le_repere(profil):
        st.caption(
            "Repère d'octobre 2026, calcul exact hors assurance : "
            "votre tableau arrondissait à environ 454 500 € d'emprunt et 774 500 € de budget sur 20 ans, "
            "516 000 € et 836 000 € sur 25 ans."
        )

    s20 = scenario(profil, 20)
    s25 = scenario(profil, 25)
    retenu = s20 if profil.duree_retenue == 20 else s25
    cible = "tout compris" if profil.inclut_assurance else "hors assurance"
    tableau(
        ["Durée", "Taux", "Mensualité", "Emprunt", "Budget", "Prix max"],
        [
            [
                f"{s['duree']} ans" + (" · retenu" if s["duree"] == retenu["duree"] else ""),
                pct(s["taux"], 2),
                eur(profil.mensualite_cible),
                eur(s["capacite"]),
                eur(s["budget"]),
                eur(prix_achat_max(s["budget"], 0, profil, hyp)),
            ]
            for s in (s20, s25)
        ],
    )
    st.caption(
        f"Mensualité cible {cible}. Budget détaillé des biens : scénario {retenu['duree']} ans. "
        f"Assurance emprunteur {pct(profil.assurance, 2)} du capital par an, "
        f"soit {eur(retenu['assurance_max'])} / mois à l'emprunt maximum "
        f"(mensualité totale {eur(retenu['mensualite_totale_max'])}). "
        f"Mobilier et déménagement : {eur(profil.mobilier)}."
    )

    st.subheader("Prix d'achat maximum selon les travaux")
    st.caption("Même enveloppe, mais une partie part en travaux. Frais de notaire, garantie, dossier et mobilier sont déjà déduits.")
    niveaux = [0, 10_000, 20_000, 40_000, 60_000]
    lignes_travaux = []
    for travaux in niveaux:
        ligne = {"Travaux": eur(travaux)}
        for s in (s20, s25):
            ligne[f"Prix max {s['duree']} ans"] = eur(prix_achat_max(s["budget"], travaux, profil, hyp))
        lignes_travaux.append(ligne)
    st.dataframe(pd.DataFrame(lignes_travaux), hide_index=True, width="stretch")

    if profil.revenus > 0:
        st.subheader("Plafond d'endettement")
        lignes = []
        plus_strict = False
        for s in (s20, s25):
            mensualite_banque = profil.revenus * profil.endettement_max - profil.autres_credits
            cap_banque = capacite_emprunt(
                max(mensualite_banque, 0), s["taux"], s["duree"], profil.assurance, True
            )
            effort = (s["mensualite_totale_max"] + profil.autres_credits) / profil.revenus
            if cap_banque + 1 < s["capacite"]:
                plus_strict = True
            lignes.append(
                {
                    "Durée": f"{s['duree']} ans",
                    "Mensualité max au plafond": eur(mensualite_banque),
                    "Emprunt max au plafond": eur(cap_banque),
                    "Effort du crédit à la cible": pct(effort),
                }
            )
        st.dataframe(pd.DataFrame(lignes), hide_index=True, width="stretch")
        if plus_strict:
            bandeau(
                "warn",
                "Le plafond de 35 % autorise moins que la mensualité cible.",
                "La banque s'arrêtera au plus petit des deux. L'assurance est incluse dans ce plafond, comme dans la recommandation du HCSF.",
            )
        else:
            bandeau(
                "ok",
                "La mensualité cible reste en dessous du plafond d'endettement.",
                "C'est elle qui limite le budget. Le plafond bancaire laisserait emprunter davantage.",
            )
    else:
        st.info("Renseignez les revenus nets du foyer plus bas pour comparer la cible de 2 600 € au plafond d'endettement de 35 %.")

    st.subheader("Paramètres du foyer")
    st.caption("C'est ici que la capacité se modifie. L'analyse d'un appartement ne la change pas.")
    semer_profil(profil)
    a, b, c = st.columns(3)
    a.number_input("Apport (€)", min_value=0.0, step=5_000.0, key="p_apport")
    b.number_input("Mensualité cible (€ / mois)", min_value=0.0, step=50.0, key="p_mensualite")
    c.selectbox("Durée utilisée pour le budget détaillé", [20, 25], key="p_duree", format_func=lambda n: f"{n} ans")
    a, b, c = st.columns(3)
    a.number_input("Taux 20 ans (% / an)", min_value=0.0, max_value=15.0, step=0.01, format="%.2f", key="p_taux20")
    b.number_input("Taux 25 ans (% / an)", min_value=0.0, max_value=15.0, step=0.01, format="%.2f", key="p_taux25")
    c.number_input("Assurance emprunteur (% du capital / an)", min_value=0.0, max_value=2.0, step=0.01, format="%.2f", key="p_assurance")
    st.checkbox(
        "La mensualité cible inclut déjà l'assurance emprunteur",
        key="p_inclut",
        help="Laisser décoché pour retrouver le tableau d'octobre : 2 600 € de crédit, assurance en plus.",
    )
    with st.expander("Revenus, mobilier, frais d'agence"):
        a, b, c = st.columns(3)
        a.number_input("Revenus nets mensuels du foyer (€)", min_value=0.0, step=100.0, key="p_revenus", help="0 si vous ne voulez pas le taux d'effort.")
        b.number_input("Autres mensualités de crédit (€)", min_value=0.0, step=50.0, key="p_autres")
        c.number_input("Plafond d'endettement (%)", min_value=1.0, max_value=50.0, step=0.5, format="%.1f", key="p_endettement")
        a, b = st.columns(2)
        a.number_input("Mobilier, déménagement, équipement (€)", min_value=0.0, step=500.0, key="p_mobilier")
        b.number_input("Honoraires d'agence à votre charge (% du prix)", min_value=0.0, max_value=10.0, step=0.1, format="%.1f", key="p_agence", help="À Paris le prix affiché est en général FAI : laisser 0.")
    # Relire après les widgets de cette page pour que le fichier suive tout de suite.
    st.session_state.profil = profil_courant()
    sauver_profil(st.session_state.profil)
    disclaimer()


def _noter(champ: str, valeur):
    st.session_state.formulaire[champ] = valeur
    return valeur


def saisie_texte(label: str, champ: str, zone: bool = False, **kwargs):
    cle = f"b_{champ}"
    if cle not in st.session_state:
        kwargs["value"] = st.session_state.formulaire.get(champ) or ""
    widget = st.text_area if zone else st.text_input
    return _noter(champ, widget(label, key=cle, **kwargs))


def saisie_nombre(label: str, champ: str, **kwargs):
    cle = f"b_{champ}"
    if cle not in st.session_state:
        kwargs["value"] = st.session_state.formulaire[champ]
    return _noter(champ, st.number_input(label, key=cle, **kwargs))


def saisie_coche(label: str, champ: str, **kwargs):
    cle = f"b_{champ}"
    if cle not in st.session_state:
        kwargs["value"] = bool(st.session_state.formulaire[champ])
    return _noter(champ, st.checkbox(label, key=cle, **kwargs))


def saisie_liste(label: str, champ: str, options: list, **kwargs):
    cle = f"b_{champ}"
    if not options:
        return _noter(champ, None)
    if cle not in st.session_state:
        actuel = st.session_state.formulaire.get(champ)
        kwargs["index"] = options.index(actuel) if actuel in options else 0
    elif st.session_state[cle] not in options:
        st.session_state[cle] = options[0]
    return _noter(champ, st.selectbox(label, options, key=cle, **kwargs))


def formulaire_bien(hyp: Hypotheses) -> None:
    st.subheader("Le bien")
    a, b = st.columns([1.2, 1])
    with a:
        saisie_texte("Nom ou adresse", "nom", placeholder="3 pièces rue Oberkampf")
    with b:
        saisie_texte("Lien de l'annonce", "lien", placeholder="https://…")
    a, b, c, d = st.columns(4)
    with a:
        saisie_liste(
            "Arrondissement",
            "arrondissement",
            list(range(1, 21)),
            format_func=lambda i: f"{ordinal(i)} — {eur(hyp.prix_m2.get(i, 0))} / m² · {QUARTIERS.get(i, '')}",
        )
    with b:
        saisie_nombre("Surface Carrez (m²)", "surface", min_value=0.0, step=0.5)
    with c:
        saisie_nombre("Étage (0 = rez-de-chaussée)", "etage", min_value=0, max_value=30, step=1)
    with d:
        saisie_liste("Immeuble", "immeuble", list(hyp.coef_immeuble))
    a, b, c, d, e = st.columns(5)
    with a:
        saisie_coche("Ascenseur", "ascenseur")
    with b:
        saisie_coche("Dernier étage", "dernier_etage")
    with c:
        saisie_coche("Cave", "cave")
    with d:
        saisie_coche("Local à vélo", "local_velo")
    with e:
        saisie_coche("Parking ou box", "parking")
    a, b, c = st.columns(3)
    with a:
        saisie_liste("État", "etat", list(hyp.coef_etat))
    with b:
        saisie_liste("DPE", "dpe", list(hyp.coef_dpe))
    with c:
        saisie_liste("Luminosité", "luminosite", list(hyp.coef_lumiere))
    a, b = st.columns(2)
    with a:
        saisie_liste("Nuisances", "nuisances", list(hyp.coef_bruit))
    with b:
        saisie_liste("Extérieur", "exterieur", list(hyp.coef_exterieur))

    st.subheader("Prix et coûts annoncés")
    a, b, c = st.columns(3)
    with a:
        saisie_nombre("Prix demandé, frais d'agence inclus (€)", "prix", min_value=0.0, step=5_000.0)
    with b:
        saisie_nombre(
            "Travaux à prévoir (€)",
            "travaux",
            min_value=0.0,
            step=1_000.0,
            help="Privatifs et quote-part de copropriété.",
        )
    with c:
        saisie_coche("J'ai une offre ou un prix négocié", "a_une_offre")
    if st.session_state.formulaire["a_une_offre"]:
        saisie_nombre("Prix retenu (€)", "prix_vise", min_value=0.0, step=5_000.0)
    a, b = st.columns(2)
    with a:
        connu = saisie_coche("Je connais les charges de copropriété", "charges_connues")
        if connu:
            saisie_nombre("Charges annuelles (€)", "charges", min_value=0.0, step=100.0)
        else:
            st.caption(f"Sinon : {eur(hyp.charges_m2)} / m² / an.")
    with b:
        connu = saisie_coche("Je connais la taxe foncière", "taxe_connue")
        if connu:
            saisie_nombre("Taxe foncière annuelle (€)", "taxe", min_value=0.0, step=50.0)
        else:
            st.caption(f"Sinon : {eur(hyp.taxe_m2)} / m² / an. L'avis du vendeur est plus fiable.")
    saisie_texte("Notes", "notes", zone=True, height=80, placeholder="Visite, feeling, points à revérifier…")


def enregistrer_bien(bien: dict) -> None:
    from datetime import date

    if float(bien.get("surface") or 0) <= 0 or float(bien.get("prix") or 0) <= 0:
        st.warning("Indiquez au moins la surface et le prix demandé pour enregistrer le bien.")
        return
    payload = dict(bien)
    payload["checklist"] = normaliser_checklist(st.session_state.get("checklist_rows"))
    payload["nom"] = (payload.get("nom") or "").strip() or "Sans nom"
    aujourd_hui = date.today().isoformat()
    biens = list(st.session_state.biens)
    if payload.get("id"):
        for i, existant in enumerate(biens):
            if existant.get("id") == payload["id"]:
                payload["cree_le"] = existant.get("cree_le") or aujourd_hui
                payload["modifie_le"] = aujourd_hui
                biens[i] = payload
                break
        else:
            payload["cree_le"] = aujourd_hui
            biens.append(payload)
    else:
        frais = nouveau_bien_vide()
        payload["id"] = frais["id"]
        payload["cree_le"] = aujourd_hui
        biens.append(payload)
        st.session_state.current_id = payload["id"]
        st.session_state.current_cree_le = payload["cree_le"]
    st.session_state.biens = biens
    sauver_biens(biens)
    st.toast("Bien enregistré")


def afficher_estimation(est: dict, hyp: Hypotheses) -> None:
    if not est.get("a_un_prix"):
        bandeau("info", "Estimation prête", "Ajoutez le prix demandé pour le verdict, les frais et le financement.")
    else:
        genres = {
            "Sous le marché – à creuser": "info",
            "Prix cohérent": "ok",
            "Un peu cher – négociable": "warn",
            "Surévalué – négocier ou passer": "bad",
        }
        bandeau(
            genres.get(est["verdict"], "info"),
            est["verdict"],
            f"Le prix demandé est à {pct(est['ecart'])} de la valeur estimée, soit {eur(est['ecart_eur'])}.",
        )
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Valeur estimée", eur(est["valeur"]))
        c2.metric("Prix demandé", eur(est["prix"]))
        c3.metric("Écart", pct(est["ecart"]))
        c4.metric("Prix demandé au m²", eur(est["prix_m2_demande"]))
        st.markdown(
            f"Fourchette de cohérence ({pct(hyp.fourchette, 0)}) : **{eur(est['basse'])}** à **{eur(est['haute'])}**.  \n"
            f"Offre de départ suggérée : **{eur(est['offre'])}**. "
            f"Plafond conseillé (haut de fourchette, arrondi) : **{eur(est['plafond'])}**."
        )
    with st.expander("D'où vient l'estimation"):
        st.caption(
            "Prix moyen de l'arrondissement × chaque coefficient, puis cave, local à vélo et parking en euros. "
            "Un coefficient de 0 % est le bien « moyen » de la table."
        )
        c1, c2, c3 = st.columns(3)
        c1.metric("Prix moyen de l'arrondissement", eur(est["prix_ref"]) + " / m²")
        c2.metric("Multiplicateur", f"{est['multiplicateur']:.3f}".replace(".", ","))
        c3.metric("Prix ajusté", eur(est["prix_m2"]) + " / m²")
        lignes = [{"Critère": nom, "Coefficient": pct(coef)} for nom, coef in est["coefs"].items()]
        lignes.append({"Critère": "Cave, local à vélo et parking", "Coefficient": eur(est["annexes"])})
        st.dataframe(pd.DataFrame(lignes), hide_index=True, width="stretch")


def afficher_financement(resultat: dict, profil: Profil, hyp: Hypotheses) -> None:
    projet = resultat["projet"]
    est = resultat["estimation"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Coût total du projet", eur(projet["total"]))
    c2.metric("Montant à emprunter", eur(projet["emprunt"]))
    c3.metric("Apport mobilisé", eur(projet["apport_utilise"]))
    if resultat["prix_retenu"] != est.get("prix"):
        st.caption(
            f"Calculé sur le prix retenu de {eur(resultat['prix_retenu'])} "
            f"({pct(resultat['ecart_retenu'])} par rapport à l'estimation : {resultat['verdict_retenu']})."
        )

    lignes = [
        ("Prix d'achat retenu", projet["prix"]),
        ("Honoraires d'agence à votre charge", projet["agence"]),
        ("Frais de notaire", projet["notaire"]["total"]),
        ("Travaux", projet["travaux"]),
        ("Mobilier et déménagement", projet["mobilier"]),
        ("Garantie du prêt", projet["garantie"]),
        ("Frais de dossier", projet["dossier"]),
        ("Frais de courtage", projet["courtage"]),
        ("Coût total", projet["total"]),
    ]
    st.dataframe(
        pd.DataFrame([{"Poste": nom, "Montant": eur(montant)} for nom, montant in lignes]),
        hide_index=True,
        width="stretch",
    )
    n = projet["notaire"]
    with st.expander(f"Détail des frais de notaire · {eur(n['total'])} · {pct(n['pct'])} du prix"):
        detail = [
            ("Droits de mutation (département, commune, frais d'assiette)", n["droits"]),
            ("Émoluments du notaire HT", n["emoluments_ht"]),
            ("TVA sur émoluments", n["tva"]),
            ("Débours et formalités", n["debours"]),
            ("Contribution de sécurité immobilière", n["csi"]),
        ]
        st.dataframe(
            pd.DataFrame([{"Poste": a, "Montant": eur(b, 0)} for a, b in detail]),
            hide_index=True,
            width="stretch",
        )
        st.caption("Ordre de grandeur attendu dans l'ancien à Paris : 7 à 8 % du prix. Le notaire donnera le chiffre exact.")

    st.subheader("Face à la mensualité de " + eur(profil.mensualite_cible))
    cartes = []
    for duree in (20, 25):
        s = resultat["scenarios"][duree]
        cartes.append(s)
    col20, col25 = st.columns(2)
    for colonne, s in zip((col20, col25), cartes):
        with colonne:
            if s["tient"]:
                bandeau("ok", f"{s['duree']} ans · ça tient", f"Mensualité {eur(s['comparable'])}, soit {eur(abs(s['ecart_mensualite']))} sous la cible.")
            else:
                bandeau("bad", f"{s['duree']} ans · au-dessus de la cible", f"Mensualité {eur(s['comparable'])}, soit {eur(s['ecart_mensualite'])} de plus.")
            st.markdown(
                f"Taux **{pct(s['taux'], 2)}**  \n"
                f"Mensualité hors assurance **{eur(s['mensualite_ha'])}**  \n"
                f"Assurance **{eur(s['assurance_mois'])}**  \n"
                f"Mensualité totale **{eur(s['mensualite'])}**  \n"
                f"Intérêts sur la durée **{eur(s['interets'])}**  \n"
                f"Coût du crédit (intérêts, assurance, frais) **{eur(s['cout_credit'])}**  \n"
                f"TAEG indicatif **{pct(s['taeg'], 2)}**  \n"
                f"Prix d'achat max avec ces travaux **{eur(resultat['prix_max'][s['duree']])}**"
            )
            if s["endettement"] is not None:
                genre = "ok" if s["sous_plafond"] else "bad"
                texte = "sous le plafond" if s["sous_plafond"] else "au-dessus du plafond"
                st.markdown(f"Endettement **{pct(s['endettement'])}**, {texte} de {pct(profil.endettement_max, 0)}.")
                st.caption(f"Revenu net minimum pour rester au plafond : {eur(s['revenu_min'])} / mois.")
            ecart_prix = resultat["prix_retenu"] - resultat["prix_max"][s["duree"]]
            if ecart_prix > 1:
                st.caption(f"Le prix retenu dépasse le prix maximum de {eur(ecart_prix)}.")
            else:
                st.caption(f"Marge sous le prix maximum : {eur(-ecart_prix)}.")


def afficher_budget(resultat: dict, profil: Profil) -> None:
    b = resultat["budget"]
    s = resultat["retenu"]
    st.caption(f"Budget détaillé du scénario {s['duree']} ans à {pct(s['taux'], 2)}. Changez la durée dans Capacité d'achat.")
    c1, c2, c3 = st.columns(3)
    c1.metric("Coût mensuel du logement", eur(b["total_mois"]))
    c2.metric("Dont crédit, assurance incluse", eur(s["mensualite"]))
    c3.metric("Possession hors crédit", eur(b["possession_an"] / 12))
    lignes = [
        ("Mensualité du crédit hors assurance", b["mensualite_ha"], b["mensualite_ha"] * 12),
        ("Assurance emprunteur", b["assurance_mois"], b["assurance_mois"] * 12),
        ("Charges de copropriété", b["charges_an"] / 12, b["charges_an"]),
        ("Taxe foncière", b["taxe_an"] / 12, b["taxe_an"]),
        ("Assurance habitation", b["assurance_an"] / 12, b["assurance_an"]),
        ("Entretien et petits travaux", b["entretien_an"] / 12, b["entretien_an"]),
        ("Provision travaux de copropriété", b["provision_an"] / 12, b["provision_an"]),
        ("Énergie, eau chaude, internet", b["energie"], b["energie"] * 12),
        ("Total logement", b["total_mois"], b["total_an"]),
    ]
    st.dataframe(
        pd.DataFrame([{"Poste": p, "Mensuel": eur(m), "Annuel": eur(a)} for p, m, a in lignes]),
        hide_index=True,
        width="stretch",
    )
    if profil.revenus > 0:
        c1, c2 = st.columns(2)
        c1.metric("Taux d'effort (logement / revenus)", pct(b["effort"]))
        c2.metric("Reste à vivre", eur(b["reste"]))
        st.caption("Le reste à vivre est après le logement et les autres crédits. Il doit couvrir le quotidien et l'épargne de précaution.")
    else:
        st.caption("Revenus non renseignés : le taux d'effort et le reste à vivre apparaîtront dès que vous les saisirez dans Capacité d'achat.")
    if b["loyer"] and b["annees_loyer"] is not None and b["rendement"] is not None:
        annees = f"{b['annees_loyer']:.1f}".replace(".", ",")
        st.markdown(
            f"Loyer de marché équivalent : **{eur(b['loyer'])} / mois** hors charges. "
            f"Rendement locatif brut implicite : **{pct(b['rendement'])}**. "
            f"Le prix représente **{annees} années de loyer**."
        )
    else:
        st.caption("Le loyer au m² est à zéro : le rendement et le nombre d'années de loyer ne sont pas calculés.")


def afficher_comparatif(resultat: dict, profil: Profil, hyp: Hypotheses) -> None:
    if "p_loyer_actuel" not in st.session_state:
        st.session_state.p_loyer_actuel = float(profil.loyer_actuel or 0)
    st.number_input(
        "Ton loyer actuel, tout compris (€ / mois)",
        min_value=0.0,
        step=50.0,
        key="p_loyer_actuel",
        help="Loyer + charges, comme ils sortent de ton compte. L'électricité n'est pas dedans : elle est la même si tu achètes.",
    )
    # Le champ est lu au tour suivant par le profil. On aligne tout de suite le calcul affiché.
    loyer_actuel = float(st.session_state.p_loyer_actuel)
    if abs(loyer_actuel - profil.loyer_actuel) > 0.5:
        profil = Profil(**{**profil.to_dict(), "loyer_actuel": loyer_actuel})
        st.session_state.profil = profil
        sauver_profil(profil)
        resultat = dict(resultat)
        resultat["comparatif"] = projection_achat_location(
            {},
            resultat["projet"],
            resultat["retenu"],
            resultat["budget"],
            profil,
            hyp,
            resultat["retenu"]["duree"],
        )

    credit = resultat["retenu"]["mensualite"]
    possession_mois = resultat["budget"]["possession_an"] / 12
    cout_achat = credit + possession_mois
    ecart_mois = cout_achat - loyer_actuel
    loyer_hors = resultat["budget"]["loyer"]
    charges_loc = resultat["budget"]["charges_an"] * hyp.part_charges_locataire / 12
    loyer_bien = loyer_hors + charges_loc
    emprunt = resultat["projet"]["emprunt"]
    interet_mois = emprunt * resultat["retenu"]["taux"] / 12
    assurance_mois = resultat["retenu"]["assurance_mois"]
    capital_mois = max(0.0, resultat["retenu"]["mensualite_ha"] - interet_mois)

    st.subheader("1. Continuer à louer, ou acheter celui-ci")
    st.markdown(
        f"Aujourd'hui tu paies **{eur(loyer_actuel)} / mois** tout compris. "
        f"Pour habiter **ce** bien, tu paierais **{eur(cout_achat)} / mois** "
        f"(crédit {eur(credit)} + copropriété, taxe foncière, assurance et entretien {eur(possession_mois)})."
    )
    if ecart_mois > 1:
        bandeau(
            "info",
            f"Acheter coûte {eur(ecart_mois)} de plus par mois que ton loyer.",
            "Dans le scénario « je reste locataire », cet écart est placé chaque année, en plus de l'apport qui n'a pas été dépensé.",
        )
    elif ecart_mois < -1:
        bandeau(
            "ok",
            f"Acheter coûte {eur(abs(ecart_mois))} de moins par mois que ton loyer.",
            "Tu dépenses moins en étant propriétaire. Le graphique crédite cette économie au scénario achat.",
        )
    else:
        bandeau("info", "Le coût mensuel d'acheter est très proche de ton loyer actuel.")

    comp = resultat["comparatif"]
    if comp["seuil"] is None:
        bandeau(
            "warn",
            f"Sur {hyp.horizon_annees} ans, rester locataire à {eur(loyer_actuel)} te laisse plus d'argent.",
            "La ligne claire reste au-dessus : l'apport placé, plus l'écart de dépenses, grossit plus vite que ce que tu récupérerais en revendant.",
        )
    elif comp["seuil"] == 0:
        bandeau("ok", "Acheter te laisse plus d'argent dès la première année, avec ces hypothèses.")
    else:
        bandeau(
            "info",
            f"Acheter devient plus intéressant à partir de l'année {comp['seuil']}.",
            "Avant cette date, le capital placé du locataire est plus élevé que ce que tu récupérerais en revendant.",
        )
    c1, c2, c3, c4 = st.columns(4)
    for colonne, annee, valeur in (
        (c1, 5, comp["avantage_5"]),
        (c2, 10, comp["avantage_10"]),
        (c3, 15, comp["avantage_15"]),
        (c4, 20, comp["avantage_20"]),
    ):
        colonne.metric(f"Achat − location à {annee} ans", eur(valeur))
    st.caption(
        "Négatif : en continuant à louer et en plaçant l'argent, tu as plus de capital à cette date. "
        f"Prix du bien +{pct(hyp.hausse_immo, 1)} / an, ton loyer +{pct(hyp.hausse_loyers, 1)} / an, "
        f"placement {pct(hyp.rendement_epargne, 1)} / an. Modifiable dans Hypothèses."
    )
    df = pd.DataFrame(comp["lignes"]).set_index("annee")
    courbe = df[["patrimoine_acheteur", "capital_locataire"]].rename(
        columns={
            "patrimoine_acheteur": "J'achète, puis je revends",
            "capital_locataire": "Je reste locataire",
        }
    )
    st.line_chart(courbe, height=320)
    st.caption(
        "La ligne foncée part plus bas : le jour de l'achat, le notaire est payé et revendre coûterait encore "
        f"{pct(hyp.frais_revente, 0)}. La ligne claire part de ton apport ({eur(profil.apport)}), qui reste placé si tu n'achètes pas."
    )

    st.subheader("2. Ce que cet appartement se louerait, face au crédit")
    st.markdown(
        "Ce n'est pas ton loyer actuel. C'est le loyer de marché de **ce** bien "
        f"({eur(hyp.loyer_m2)} / m² hors charges, dans Hypothèses), plus la part des charges qu'un locataire paie."
    )
    c1, c2, c3 = st.columns(3)
    c1.metric("Loyer potentiel, charges comprises", eur(loyer_bien))
    c2.metric("Mensualité du crédit, assurance incluse", eur(credit))
    c3.metric("Crédit − loyer potentiel", eur(credit - loyer_bien))
    if credit > loyer_bien + 1:
        st.markdown(
            f"Le crédit coûte **{eur(credit - loyer_bien)} de plus par mois** que si quelqu'un louait ce logement. "
            f"Sur cette mensualité, environ **{eur(capital_mois)}** remboursent le capital "
            f"(tu le récupères à la revente) et **{eur(interet_mois + assurance_mois)}** sont des intérêts et de l'assurance "
            "(ça, c'est le coût de la banque)."
        )
    elif loyer_bien > credit + 1:
        st.markdown(
            f"Le crédit coûte **{eur(loyer_bien - credit)} de moins par mois** que le loyer potentiel de ce bien. "
            f"Environ **{eur(capital_mois)}** de la mensualité remboursent le capital, "
            f"**{eur(interet_mois + assurance_mois)}** partent en intérêts et assurance."
        )
    else:
        st.markdown("Le crédit et le loyer potentiel de ce bien sont presque au même niveau.")
    st.caption(f"Détail du loyer potentiel : {eur(loyer_hors)} de loyer + {eur(charges_loc)} de charges locatives.")

    st.subheader("3. Au bout de combien d'années tu peux revendre sans perte")
    rev = resultat["revente"]
    st.markdown(
        "Tu y habites. On regarde seulement l'opération : si tu revends, est-ce que tu récupères "
        "ton apport, le notaire, les intérêts et l'assurance déjà payés ? "
        "La copropriété et la taxe foncière ne sont pas là-dedans, c'est le coût d'y vivre, déjà comparé à ton loyer plus haut."
    )
    bandeau(
        "warn" if rev["perte_immediate"] < 0 else "ok",
        f"Le jour de l'achat, revendre te ferait perdre {eur(abs(rev['perte_immediate']))}.",
        "C'est le notaire, les frais de prêt, les travaux, le mobilier, et les frais de revente. "
        "Le graphique montre en combien d'années la hausse du prix et le capital remboursé comblent ce trou, intérêts compris.",
    )
    if rev["seuil"] is None:
        bandeau(
            "warn",
            f"Sur {hyp.horizon_annees} ans, une revente reste perdante de {eur(abs(rev['resultat_final']))}.",
            f"Avec une hausse des prix de {pct(hyp.hausse_immo, 1)} par an, le bien ne reprend pas assez de valeur pour couvrir le notaire et les intérêts.",
        )
    elif rev["seuil"] == 0:
        bandeau("ok", "Tu peux revendre dès la première année sans perte, avec ces hypothèses.")
    else:
        bandeau(
            "ok",
            f"Tu peux revendre sans perte à partir de l'année {rev['seuil']}.",
            "Avant, le prix de vente ne couvre pas encore le crédit restant, ton apport, les intérêts et l'assurance.",
        )
    courbe_revente = pd.DataFrame(rev["lignes"]).set_index("annee")
    courbe_revente["Ni gain ni perte"] = 0.0
    st.line_chart(
        courbe_revente[["resultat", "Ni gain ni perte"]].rename(
            columns={"resultat": "Résultat de la revente"}
        ),
        height=320,
    )
    st.caption(
        "Au-dessus de zéro : tu ressors avec plus que ce que l'opération t'a coûté. "
        "En dessous : le notaire et les intérêts ne sont pas encore rattrapés. "
        "Le changement de pente, c'est la fin du crédit : tu ne paies plus d'intérêts."
    )
    with st.expander("Détail année par année"):
        st.markdown("**Rester locataire ou acheter**")
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Année": r["annee"],
                        "Si je revends, argent récupéré": eur(r["patrimoine_acheteur"]),
                        "Si je continue à louer, capital placé": eur(r["capital_locataire"]),
                        "Achat − location": eur(r["avantage"]),
                    }
                    for r in comp["lignes"]
                ]
            ),
            hide_index=True,
            width="stretch",
        )
        st.markdown("**Revente sans perte**")
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Année": r["annee"],
                        "Valeur du bien": eur(r["valeur"]),
                        "Crédit restant": eur(r["crd"]),
                        "Intérêts déjà payés": eur(r["interets"]),
                        "Assurance déjà payée": eur(r["assurance"]),
                        "Argent récupéré": eur(r["cash"]),
                        "Résultat": eur(r["resultat"]),
                    }
                    for r in rev["lignes"]
                ]
            ),
            hide_index=True,
            width="stretch",
        )
    st.caption(
        "Ça ne compte pas le fait d'être chez toi, ni la liberté de faire des travaux. "
        "La plus-value d'une résidence principale est exonérée d'impôt."
    )


def afficher_sensibilite(resultat: dict) -> None:
    emprunt = resultat["projet"]["emprunt"]
    st.caption(
        f"Même montant emprunté ({eur(emprunt)}), autour du taux du scénario "
        f"{resultat['retenu']['duree']} ans ({pct(resultat['retenu']['taux'], 2)}). "
        "Allonger la durée baisse la mensualité et augmente le coût total."
    )
    lignes_mens = []
    lignes_cout = []
    for ligne in resultat["sensibilite"]:
        mens = {"Taux": pct(ligne["taux"], 2)}
        cout = {"Taux": pct(ligne["taux"], 2)}
        for duree, pret in ligne["durees"].items():
            mens[f"{duree} ans"] = eur(pret["mensualite"])
            cout[f"{duree} ans"] = eur(pret["interets"] + pret["cout_assurance"])
        lignes_mens.append(mens)
        lignes_cout.append(cout)
    st.markdown("**Mensualité totale, assurance incluse**")
    st.dataframe(pd.DataFrame(lignes_mens), hide_index=True, width="stretch")
    st.markdown("**Intérêts + assurance sur toute la durée**")
    st.dataframe(pd.DataFrame(lignes_cout), hide_index=True, width="stretch")


def afficher_checklist() -> None:
    st.caption("35 points repris du classeur. Le statut est enregistré avec le bien.")
    edited = st.data_editor(
        cadre_checklist(st.session_state.checklist_rows),
        hide_index=True,
        width="stretch",
        num_rows="fixed",
        height=460,
        disabled=["Thème", "Point", "Pourquoi"],
        column_config={
            "Statut": st.column_config.SelectboxColumn("Statut", options=STATUTS, required=True),
            "Pourquoi": st.column_config.TextColumn(width="large"),
            "Notes": st.column_config.TextColumn(width="medium"),
        },
        key="editeur_checklist",
    )
    st.session_state.checklist_rows = [
        {"statut": row["Statut"], "notes": row["Notes"] or ""} for row in edited.to_dict("records")
    ]
    comptes = {statut: 0 for statut in STATUTS}
    for ligne in st.session_state.checklist_rows:
        comptes[ligne["statut"]] = comptes.get(ligne["statut"], 0) + 1
    st.caption(
        f"OK {comptes['OK']} · Alerte {comptes['Alerte']} · À faire {comptes['À faire']} · N/A {comptes['N/A']}"
    )


def page_analyser() -> None:
    hyp = st.session_state.hyp
    profil = st.session_state.profil
    st.title("Analyser un bien")
    st.caption(
        f"Apport {eur(profil.apport)} · cible {eur(profil.mensualite_cible)} / mois · "
        f"{pct(profil.taux_20, 2)} sur 20 ans et {pct(profil.taux_25, 2)} sur 25 ans."
    )
    actions = st.columns([1, 1, 2])
    if actions[0].button("Nouveau bien", width="stretch"):
        st.session_state.pending_new = True
        st.rerun()
    if actions[1].button("Exemple du classeur", width="stretch"):
        st.session_state.pending_exemple = True
        st.rerun()

    formulaire_bien(hyp)
    bien = lire_bien()
    # La checklist de l'onglet peut déjà avoir été éditée au tour précédent.
    bien["checklist"] = normaliser_checklist(st.session_state.get("checklist_rows"))

    if st.button("Enregistrer ce bien", type="primary"):
        enregistrer_bien(lire_bien())

    resultat = analyser(bien, profil, hyp)
    st.divider()
    if resultat is None:
        bandeau("info", "Indiquez l'arrondissement et la surface Carrez.", "Le reste des calculs se remplit ensuite tout seul.")
        disclaimer()
        return

    onglets = st.tabs(["Estimation", "Financement", "Budget mensuel", "Louer ou acheter", "Sensibilité", "Checklist visite"])
    with onglets[0]:
        afficher_estimation(resultat["estimation"], hyp)
    if not resultat["pret"]:
        for onglet in onglets[1:5]:
            with onglet:
                st.info("Ajoutez le prix demandé pour calculer le financement, le budget et la comparaison avec la location.")
    else:
        with onglets[1]:
            afficher_financement(resultat, profil, hyp)
        with onglets[2]:
            afficher_budget(resultat, profil)
        with onglets[3]:
            afficher_comparatif(resultat, profil, hyp)
        with onglets[4]:
            afficher_sensibilite(resultat)
    with onglets[5]:
        afficher_checklist()
    disclaimer()


def page_biens() -> None:
    st.title("Biens enregistrés")
    biens = st.session_state.biens
    profil = st.session_state.profil
    hyp = st.session_state.hyp
    if not biens:
        bandeau("info", "Aucun bien enregistré.", "Analysez un appartement, puis enregistrez-le pour le comparer aux suivants.")
        disclaimer()
        return

    resumes = [resume_bien(b, profil, hyp) for b in biens]
    resumes.sort(key=lambda r: (r.get("ecart") is None, r.get("ecart") if r.get("ecart") is not None else 0))
    lignes = []
    for r in resumes:
        lignes.append(
            {
                "Bien": r["nom"],
                "Arr.": ordinal(r["arrondissement"]) if r.get("arrondissement") else "—",
                "Surface": f"{float(r['surface']):.0f} m²" if r.get("surface") else "—",
                "Prix": eur(r["prix"]) if r.get("prix") else "—",
                "Valeur estimée": "—" if r.get("incomplet") else eur(r["valeur"]),
                "Écart": "—" if r.get("incomplet") else pct(r["ecart"]),
                "Verdict": "À compléter" if r.get("incomplet") else r["verdict"],
                "Coût total": "—" if r.get("incomplet") else eur(r["cout"]),
                "Mensualité 20 ans": "—" if r.get("incomplet") else eur(r["mensualite_20"]),
                "Mensualité 25 ans": "—" if r.get("incomplet") else eur(r["mensualite_25"]),
                "Tient sur 25 ans": "—" if r.get("incomplet") else ("Oui" if r["tient_25"] else "Non"),
            }
        )
    st.dataframe(pd.DataFrame(lignes), hide_index=True, width="stretch")
    st.download_button(
        "Télécharger le comparatif (CSV)",
        data=pd.DataFrame(lignes).to_csv(index=False).encode("utf-8"),
        file_name="biens_residence.csv",
        mime="text/csv",
    )

    complets = [r for r in resumes if not r.get("incomplet")]
    if complets:
        courbe = pd.DataFrame(
            {
                "Bien": [r["nom"] for r in complets],
                "Prix demandé": [r["prix"] for r in complets],
                "Valeur estimée": [r["valeur"] for r in complets],
            }
        )
        st.bar_chart(courbe, x="Bien", y=["Prix demandé", "Valeur estimée"], height=280)

    st.subheader("Ouvrir ou supprimer")
    libelles = {b["id"]: b.get("nom") or "Sans nom" for b in biens}
    choix = st.selectbox("Bien", list(libelles), format_func=lambda i: libelles[i])
    a, b = st.columns(2)
    if a.button("Ouvrir dans l'analyse", type="primary", width="stretch"):
        st.session_state.pending_bien_id = choix
        st.session_state.goto = "Analyser un bien"
        st.rerun()
    if b.checkbox("Je confirme la suppression"):
        if st.button("Supprimer ce bien", width="stretch"):
            st.session_state.biens = [x for x in biens if x.get("id") != choix]
            sauver_biens(st.session_state.biens)
            if st.session_state.get("current_id") == choix:
                st.session_state.pending_new = True
            st.rerun()
    disclaimer()


def semer_hypotheses(hyp: Hypotheses) -> None:
    valeurs = {
        "h_droit": hyp.droit_dept * 100,
        "h_communal": hyp.taxe_communale * 100,
        "h_assiette": hyp.frais_assiette * 100,
        "h_tva": hyp.tva_emoluments * 100,
        "h_debours": float(hyp.debours),
        "h_csi": hyp.csi * 100,
        "h_charges": float(hyp.charges_m2),
        "h_taxe": float(hyp.taxe_m2),
        "h_assur_hab": float(hyp.assurance_habitation),
        "h_entretien": hyp.entretien_pct * 100,
        "h_provision": hyp.provision_copro_pct * 100,
        "h_energie": float(hyp.energie_mois),
        "h_loyer": float(hyp.loyer_m2),
        "h_garantie": hyp.garantie_pct * 100,
        "h_dossier": float(hyp.frais_dossier),
        "h_courtage": float(hyp.frais_courtage),
        "h_cave": float(hyp.valeur_cave),
        "h_local_velo": float(hyp.valeur_local_velo),
        "h_parking": float(hyp.valeur_parking),
        "h_fourchette": hyp.fourchette * 100,
        "h_sous": hyp.seuil_sous_marche * 100,
        "h_coherent": hyp.seuil_coherent * 100,
        "h_neg": hyp.seuil_negociable * 100,
        "h_dernier_asc": hyp.coef_dernier_asc * 100,
        "h_dernier_sans": hyp.coef_dernier_sans * 100,
        "h_part_loc": hyp.part_charges_locataire * 100,
        "h_assur_loc": float(hyp.assurance_locataire),
        "h_hausse_immo": hyp.hausse_immo * 100,
        "h_hausse_loyer": hyp.hausse_loyers * 100,
        "h_inflation": hyp.inflation_charges * 100,
        "h_rendement": hyp.rendement_epargne * 100,
        "h_revente": hyp.frais_revente * 100,
    }
    for cle, valeur in valeurs.items():
        if cle not in st.session_state:
            st.session_state[cle] = valeur


def editeur_map(titre: str, mapping: dict, cle: str) -> dict:
    st.markdown(f"**{titre}**")
    base = pd.DataFrame({"Libellé": list(mapping), "Coefficient": [float(v) for v in mapping.values()]})
    edited = st.data_editor(
        base,
        key=cle,
        hide_index=True,
        disabled=["Libellé"],
        num_rows="fixed",
        column_config={"Coefficient": st.column_config.NumberColumn(format="%.3f", step=0.01, help="0,03 = +3 %")},
    )
    return {row["Libellé"]: float(row["Coefficient"]) for row in edited.to_dict("records")}


def page_hypotheses() -> None:
    hyp = st.session_state.hyp
    semer_hypotheses(hyp)
    nonce = st.session_state.hyp_nonce
    st.title("Hypothèses")
    st.markdown(
        "Ce sont les paramètres du classeur. Un changement ici recalcule tous les biens, y compris ceux déjà enregistrés. "
        "Les taux et l'apport du foyer sont dans Capacité d'achat."
    )
    if st.button("Revenir aux hypothèses d'origine"):
        for cle in list(st.session_state):
            if str(cle).startswith("h_") or str(cle).startswith("ed_"):
                del st.session_state[cle]
        st.session_state.hyp = reinitialiser_hypotheses()
        st.session_state.hyp_nonce = nonce + 1
        st.rerun()

    st.subheader("Prix moyens au m²")
    st.caption("Moyenne de l'arrondissement. À recaler avec 3 à 5 ventes récentes de la même rue (base DVF).")
    df_prix = pd.DataFrame(
        [
            {
                "Arrondissement": ordinal(i),
                "Quartiers": QUARTIERS.get(i, ""),
                "Prix €/m²": float(hyp.prix_m2.get(i, 0)),
            }
            for i in range(1, 21)
        ]
    )
    ed_prix = st.data_editor(
        df_prix,
        key=f"ed_prix_{nonce}",
        hide_index=True,
        disabled=["Arrondissement", "Quartiers"],
        num_rows="fixed",
        column_config={"Prix €/m²": st.column_config.NumberColumn(step=50)},
    )

    st.subheader("Frais d'acquisition")
    a, b, c = st.columns(3)
    a.number_input("Droit départemental (%)", min_value=0.0, step=0.01, format="%.2f", key="h_droit")
    b.number_input("Taxe communale (%)", min_value=0.0, step=0.01, format="%.2f", key="h_communal")
    c.number_input("Frais d'assiette (% du droit départemental)", min_value=0.0, step=0.01, format="%.2f", key="h_assiette")
    a, b, c = st.columns(3)
    a.number_input("TVA sur émoluments (%)", min_value=0.0, step=0.1, format="%.1f", key="h_tva")
    b.number_input("Débours et formalités (€)", min_value=0.0, step=50.0, key="h_debours")
    c.number_input("Contribution de sécurité immobilière (%)", min_value=0.0, step=0.01, format="%.2f", key="h_csi")
    st.caption("Émoluments du notaire : barème dégressif 3,870 % / 1,596 % / 1,064 % / 0,799 % sur les tranches 6 500 €, 17 000 € et 60 000 €.")

    st.subheader("Coûts de possession, quand l'annonce ne les donne pas")
    a, b, c = st.columns(3)
    a.number_input("Charges de copropriété (€/m²/an)", min_value=0.0, step=1.0, key="h_charges")
    b.number_input("Taxe foncière (€/m²/an)", min_value=0.0, step=1.0, key="h_taxe")
    c.number_input("Assurance habitation (€/an)", min_value=0.0, step=10.0, key="h_assur_hab")
    a, b, c = st.columns(3)
    a.number_input("Entretien privatif (% du prix / an)", min_value=0.0, step=0.05, format="%.2f", key="h_entretien")
    b.number_input("Provision travaux copropriété (% du prix / an)", min_value=0.0, step=0.05, format="%.2f", key="h_provision")
    c.number_input("Énergie, eau, internet (€/mois)", min_value=0.0, step=10.0, key="h_energie")
    st.number_input("Loyer de marché équivalent (€/m²/mois, hors charges)", min_value=0.0, step=0.5, key="h_loyer")

    st.subheader("Frais de prêt et lecture du prix")
    a, b, c = st.columns(3)
    a.number_input("Frais de garantie (% du besoin financé)", min_value=0.0, step=0.1, format="%.2f", key="h_garantie")
    b.number_input("Frais de dossier (€)", min_value=0.0, step=100.0, key="h_dossier")
    c.number_input("Frais de courtage (€)", min_value=0.0, step=100.0, key="h_courtage")
    a, b, c = st.columns(3)
    a.number_input("Valeur d'une cave (€)", min_value=0.0, step=500.0, key="h_cave")
    b.number_input("Valeur d'un local à vélo (€)", min_value=0.0, step=500.0, key="h_local_velo")
    c.number_input("Valeur d'un parking (€)", min_value=0.0, step=1_000.0, key="h_parking")
    a, b, c, d = st.columns(4)
    a.number_input("Fourchette autour de l'estimation (± %)", min_value=0.0, step=0.5, key="h_fourchette")
    b.number_input("Seuil « sous le marché » (écart ≤ %)", step=0.5, key="h_sous")
    c.number_input("Seuil « prix cohérent » (écart ≤ %)", step=0.5, key="h_coherent")
    d.number_input("Seuil « négociable » (écart ≤ %)", step=0.5, key="h_neg")
    a, b = st.columns(2)
    a.number_input("Supplément dernier étage avec ascenseur (%)", step=0.5, key="h_dernier_asc")
    b.number_input("Supplément dernier étage sans ascenseur (%)", step=0.5, key="h_dernier_sans")

    with st.expander("Coefficients de valorisation"):
        st.caption("Les libellés restent ceux du formulaire. 0,03 signifie +3 % sur la valeur.")
        df_etage = pd.DataFrame(
            {
                "Étage": ETAGE_LABELS,
                "Avec ascenseur": [float(x) for x in hyp.etage_avec],
                "Sans ascenseur": [float(x) for x in hyp.etage_sans],
            }
        )
        ed_etage = st.data_editor(
            df_etage,
            key=f"ed_etage_{nonce}",
            hide_index=True,
            disabled=["Étage"],
            num_rows="fixed",
            column_config={
                "Avec ascenseur": st.column_config.NumberColumn(format="%.3f", step=0.01),
                "Sans ascenseur": st.column_config.NumberColumn(format="%.3f", step=0.01),
            },
        )
        df_surf = pd.DataFrame(
            {
                "À partir de (m²)": [float(s) for s, _ in hyp.coef_surface],
                "Coefficient": [float(c) for _, c in hyp.coef_surface],
            }
        )
        st.markdown("**Surface** — le m² des petites surfaces est plus cher. Le coefficient appliqué est celui de la plus grande tranche atteinte.")
        ed_surf = st.data_editor(
            df_surf,
            key=f"ed_surf_{nonce}",
            hide_index=True,
            num_rows="fixed",
            column_config={"Coefficient": st.column_config.NumberColumn(format="%.3f", step=0.01)},
        )
        maps = {
            "coef_etat": editeur_map("État", hyp.coef_etat, f"ed_etat_{nonce}"),
            "coef_dpe": editeur_map("DPE", hyp.coef_dpe, f"ed_dpe_{nonce}"),
            "coef_lumiere": editeur_map("Luminosité", hyp.coef_lumiere, f"ed_lum_{nonce}"),
            "coef_bruit": editeur_map("Nuisances", hyp.coef_bruit, f"ed_bruit_{nonce}"),
            "coef_exterieur": editeur_map("Extérieur", hyp.coef_exterieur, f"ed_ext_{nonce}"),
            "coef_immeuble": editeur_map("Immeuble", hyp.coef_immeuble, f"ed_imm_{nonce}"),
        }
    if "maps" not in locals():
        maps = None
        ed_etage = None
        ed_surf = None

    st.subheader("Achat contre location")
    a, b, c = st.columns(3)
    a.number_input("Hausse annuelle des prix (%)", step=0.1, format="%.1f", key="h_hausse_immo")
    b.number_input("Hausse annuelle des loyers (%)", step=0.1, format="%.1f", key="h_hausse_loyer")
    c.number_input("Inflation des charges et taxes (%)", step=0.1, format="%.1f", key="h_inflation")
    a, b, c = st.columns(3)
    a.number_input("Rendement net du capital placé (%)", step=0.1, format="%.1f", key="h_rendement")
    b.number_input("Frais de revente (%)", step=0.5, format="%.1f", key="h_revente")
    c.number_input("Part des charges aussi payée par un locataire (%)", min_value=0.0, max_value=100.0, step=5.0, key="h_part_loc")
    st.number_input("Assurance habitation du locataire (€/an)", min_value=0.0, step=10.0, key="h_assur_loc")

    nouveau = Hypotheses(
        droit_dept=float(st.session_state.h_droit) / 100,
        taxe_communale=float(st.session_state.h_communal) / 100,
        frais_assiette=float(st.session_state.h_assiette) / 100,
        tva_emoluments=float(st.session_state.h_tva) / 100,
        debours=float(st.session_state.h_debours),
        csi=float(st.session_state.h_csi) / 100,
        charges_m2=float(st.session_state.h_charges),
        taxe_m2=float(st.session_state.h_taxe),
        assurance_habitation=float(st.session_state.h_assur_hab),
        entretien_pct=float(st.session_state.h_entretien) / 100,
        provision_copro_pct=float(st.session_state.h_provision) / 100,
        energie_mois=float(st.session_state.h_energie),
        loyer_m2=float(st.session_state.h_loyer),
        garantie_pct=float(st.session_state.h_garantie) / 100,
        frais_dossier=float(st.session_state.h_dossier),
        frais_courtage=float(st.session_state.h_courtage),
        valeur_cave=float(st.session_state.h_cave),
        valeur_local_velo=float(st.session_state.h_local_velo),
        valeur_parking=float(st.session_state.h_parking),
        fourchette=float(st.session_state.h_fourchette) / 100,
        seuil_sous_marche=float(st.session_state.h_sous) / 100,
        seuil_coherent=float(st.session_state.h_coherent) / 100,
        seuil_negociable=float(st.session_state.h_neg) / 100,
        coef_dernier_asc=float(st.session_state.h_dernier_asc) / 100,
        coef_dernier_sans=float(st.session_state.h_dernier_sans) / 100,
        part_charges_locataire=float(st.session_state.h_part_loc) / 100,
        assurance_locataire=float(st.session_state.h_assur_loc),
        hausse_immo=float(st.session_state.h_hausse_immo) / 100,
        hausse_loyers=float(st.session_state.h_hausse_loyer) / 100,
        inflation_charges=float(st.session_state.h_inflation) / 100,
        rendement_epargne=float(st.session_state.h_rendement) / 100,
        frais_revente=float(st.session_state.h_revente) / 100,
        prix_m2={i: float(ed_prix.iloc[i - 1]["Prix €/m²"]) for i in range(1, 21)},
    )
    # Les coefficients ne sont dans la page que si l'expandeur a été ouvert au moins une fois.
    # Streamlit exécute le contenu des expanders à chaque run : les éditeurs existent toujours.
    if maps is not None and ed_etage is not None and ed_surf is not None:
        nouveau.etage_avec = [float(x) for x in ed_etage["Avec ascenseur"]]
        nouveau.etage_sans = [float(x) for x in ed_etage["Sans ascenseur"]]
        paires = sorted(
            (
                (float(s), float(c))
                for s, c in zip(ed_surf["À partir de (m²)"], ed_surf["Coefficient"])
            ),
            key=lambda x: x[0],
        )
        nouveau.coef_surface = [[s, c] for s, c in paires]
        nouveau.coef_etat = maps["coef_etat"]
        nouveau.coef_dpe = maps["coef_dpe"]
        nouveau.coef_lumiere = maps["coef_lumiere"]
        nouveau.coef_bruit = maps["coef_bruit"]
        nouveau.coef_exterieur = maps["coef_exterieur"]
        nouveau.coef_immeuble = maps["coef_immeuble"]
    else:
        nouveau.etage_avec = list(hyp.etage_avec)
        nouveau.etage_sans = list(hyp.etage_sans)
        nouveau.coef_surface = [list(x) for x in hyp.coef_surface]
        nouveau.coef_etat = dict(hyp.coef_etat)
        nouveau.coef_dpe = dict(hyp.coef_dpe)
        nouveau.coef_lumiere = dict(hyp.coef_lumiere)
        nouveau.coef_bruit = dict(hyp.coef_bruit)
        nouveau.coef_exterieur = dict(hyp.coef_exterieur)
        nouveau.coef_immeuble = dict(hyp.coef_immeuble)

    st.session_state.hyp = nouveau
    sauver_hypotheses(nouveau)
    # Prix de contrôle, pour voir tout de suite l'effet d'un taux de notaire modifié.
    exemple = frais_notaire(620_000, nouveau)
    st.caption(f"Contrôle : frais de notaire sur 620 000 € = {eur(exemple['total'])} ({pct(exemple['pct'])}).")
    disclaimer()


def main() -> None:
    init()
    preparer()
    barre_laterale()
    page = st.session_state.nav
    if page == "Capacité d'achat":
        page_capacite()
    elif page == "Analyser un bien":
        page_analyser()
    elif page == "Biens enregistrés":
        page_biens()
    else:
        page_hypotheses()


main()
