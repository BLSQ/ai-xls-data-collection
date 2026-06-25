DONOR_IDS = range(1, 8)

PROVINCE_MAPPING = {
    "nu Nord Ubangi Province": "nord_ubangi",
    "sn Sankuru Province": "sankuru",
    "lm Lomami Province": "lomami",
    "tn Tanganyika Province": "tanganyika",
    "hl Haut Lomami Province": "haut_lomami",
    "sk Sud Kivu Province": "sud_kivu",
    "bu Bas Uele Province": "bas_uele",
    "hk Haut Katanga Province": "haut_katanga",
    "it Ituri Province": "ituri",
    "eq Equateur Province": "equateur",
    "kc Kongo Central Province": "congo_central",
    "ke Kasai Oriental Province": "kasai_oriental",
    "ll Lualaba Province": "lualaba",
    "kg Kwango Province": "kwango",
    "mg Mongala Province": "mongala",
    "mn Maniema Province": "maniema",
    "nk Nord Kivu Province": "nord_kivu",
    "su Sud Ubangi Province": "sud_ubangi",
    "hu Haut Uele Province": "haut_uele",
    "kn Kinshasa Province": "kinshasa",
    "tu Tshuapa Province": "tshuapa",
    "md Maindombe Province": "mai_ndombe",
    "tp Tshopo Province": "tshopo",
    "kl Kwilu Province": "kwilu",
    "kr Kasai Central Province": "kasai_central",
    "ks Kasai Province": "kasai",
    # Special cases
    "République Démocratique du Congo": "national",
}

FRENCH_MONTHS = {
    1: "Janvier",
    2: "Février",
    3: "Mars",
    4: "Avril",
    5: "Mai",
    6: "Juin",
    7: "Juillet",
    8: "Août",
    9: "Septembre",
    10: "Octobre",
    11: "Novembre",
    12: "Décembre",
}

META_JOIN_COLS = [
    "entity_id",
    "nom_du_projet_programme_intitule_du_budget",
    "sigle",
    "date_debut_du_projet",
    "date_fin_du_projet",
]
