# -*- coding: utf-8 -*-
# App: Deprovisioning Consip – robusta ai file EN/IT
# Autore: adattato per funzionare con intestazioni colonne in Italiano/English

import re
import csv
import io
from datetime import datetime
from typing import List, Set, Tuple, Optional

import streamlit as st
import pandas as pd

# ==============
# Costanti CSV
# ==============

# Header CSV per lo Step 1 (Utente)
HEADER_MODIFICA = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

# Header CSV per Device
HEADER_DEVICE = [
    "Computer", "OU", "add_mail", "remove_mail",
    "add_mobile", "remove_mobile",
    "add_userprincipalname", "remove_userprincipalname",
    "disable", "moveToOU"
]

# =========================
# Utility per colonne EN/IT
# =========================

_SPACES_UNDERS = re.compile(r"[\s_]+")


def _norm_key(s: str) -> str:
    """Normalizza una chiave colonna (minuscolo, senza spazi/underscore)."""
    return _SPACES_UNDERS.sub("", str(s).strip().lower())


def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Ritorna il nome della prima colonna di df che corrisponde ad almeno
    uno dei 'candidates' (case/space/underscore insensitive).
    Tenta prima match esatto normalizzato, poi 'contains'.
    """
    if df is None or df.empty:
        return None
    norm_map = {_norm_key(c): c for c in df.columns}
    # tentativo 1: match esatto normalizzato
    for cand in candidates:
        key = _norm_key(cand)
        if key in norm_map:
            return norm_map[key]
    # tentativo 2: contains
    for cand in candidates:
        ckey = _norm_key(cand)
        for col in df.columns:
            if ckey in _norm_key(col):
                return col
    return None


def _get_any(df: pd.DataFrame, candidates: List[str]) -> pd.Series:
    """Ritorna la Series della prima colonna trovata tra i candidates; se non c'è, solleva KeyError."""
    col = _find_col(df, candidates)
    if not col:
        raise KeyError(f"Columns not found: {candidates}")
    return df[col]


def _require_any(df: pd.DataFrame, required_map: dict, context: str) -> Tuple[bool, List[str]]:
    """
    Verifica che per ogni chiave di required_map almeno una delle colonne candidate esista.
    Ritorna (ok, missing_list).
    """
    missing = []
    for logical_name, cand_list in required_map.items():
        if _find_col(df, cand_list) is None:
            missing.append(logical_name)
    return (len(missing) == 0, missing)


def _clean_series_to_list(series: pd.Series) -> List[str]:
    """
    Converte in lista di stringhe uniche/ordinate, rimuovendo vuoti.
    """
    if series is None or series.empty:
        return []
    vals = [str(x).strip() for x in series.dropna().astype(str).tolist()]
    vals = [v for v in vals if v != ""]
    # unici e ordinati (case-insensitive)
    return sorted(list(set(vals)), key=lambda x: x.lower())


def _read_excel_or_empty(uploaded_file) -> pd.DataFrame:
    """Legge un Excel con engine openpyxl; se non presente o errore, ritorna DF vuoto."""
    if not uploaded_file:
        return pd.DataFrame()
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception:
        # fallback a default (nel caso openpyxl non sia disponibile nell'ambiente)
        try:
            return pd.read_excel(uploaded_file)
        except Exception:
            return pd.DataFrame()


# =========================
# Candidati colonne EN/IT
# =========================

# Membri / gruppi (export AD – "Estr_MembriGruppi")
CAND_MG_MEMBER = [
    "Member", "Membro",
    "MemberSamAccountName", "MembroSamAccountName",
    "SamAccountName", "sAMAccountName",
    "MemberUserPrincipalName", "UserPrincipalNameMembro",
    "userPrincipalName", "UPN",
]
CAND_MG_GROUP = [
    "Group", "Gruppo",
    "GroupName", "NomeGruppo",
    "DisplayName", "NomeVisualizzato",
    "cn", "CN", "Name", "Nome"
]

# Distribution List (DL) file
CAND_DL_GROUP = [
    "Distribution Group", "Gruppo di distribuzione",
    "Group", "Gruppo",
    "GroupName", "NomeGruppo",
    "DisplayName", "Nome", "NomeVisualizzato"
]
CAND_DL_EMAIL = [
    "SMTP", "PrimarySmtpAddress",
    "Email", "E-mail", "Mail", "Posta", "Indirizzo email", "Indirizzo SMTP"
]

# Shared Mailbox / SM file
CAND_SM_MEMBER_EMAIL = [
    "MemberUserPrincipalName", "UserPrincipalNameMembro",
    "userPrincipalName", "UPN",
    "MemberEmail", "EmailMembro", "Member Mail", "Email"
]
CAND_SM_GROUP_NAME = [
    "Group", "Gruppo",
    "DisplayName", "Nome", "NomeVisualizzato",
    "Mailbox", "SharedMailbox", "Cassetta postale", "Casella condivisa",
    "PrimarySmtpAddress", "SMTP", "Email"
]

# Entra (gruppi utente)
CAND_ENTRA_MEMBER_UPN = [
    "MemberUserPrincipalName", "UserPrincipalNameMembro",
    "userPrincipalName", "UPN",
    "Email", "E-mail"
]
CAND_ENTRA_GROUP_NAME = [
    "GroupName", "NomeGruppo",
    "DisplayName", "Nome", "NomeVisualizzato",
    "Group", "Gruppo"
]

# Device export
CAND_DEV_ENABLED = ["Enabled", "Abilitato"]
CAND_DEV_DESC = ["Description", "Descrizione"]
CAND_DEV_NAME = ["Name", "Nome", "Computer", "NomeComputer"]
CAND_DEV_MAIL = ["Mail", "Email", "E-mail", "Posta"]
CAND_DEV_MOBILE = ["Mobile", "Cellulare", "Telefono", "Phone"]
CAND_DEV_UPN = ["userPrincipalName", "UPN", "NomePrincipaleUtente"]


# ====================================================
# Funzione per comporre la stringa di rimozione gruppi
# ====================================================

def estrai_rimozione_gruppi(sam_lower: str, mg_df: pd.DataFrame) -> str:
    """
    Ritorna stringa gruppi AD (da "Estr_MembriGruppi") da rimuovere, separata da ';',
    con esclusioni note (Domain Users/Utenti del dominio).
    Supporta intestazioni EN/IT.
    """
    if mg_df is None or mg_df.empty:
        return ""

    required = {
        "member": CAND_MG_MEMBER,
        "group": CAND_MG_GROUP
    }
    ok, missing = _require_any(mg_df, required, "Estr_MembriGruppi")
    if not ok:
        # Se mancano colonne, non blocchiamo l'esecuzione: ritorniamo stringa vuota
        st.warning(f"Nel file 'Estr_MembriGruppi' non ho trovato i campi: {', '.join(missing)}")
        return ""

    member_col = _get_any(mg_df, CAND_MG_MEMBER)
    group_col = _get_any(mg_df, CAND_MG_GROUP)

    # Confronta sia sAMAccountName sia UPN (per tolleranza)
    user_email = f"{sam_lower}@consip.it"
    member_vals = member_col.astype(str).str.strip().str.lower()
    mask = (member_vals == sam_lower) | (member_vals == user_email)

    groups_series = group_col[mask].dropna().astype(str)

    # Escludi gruppi generici/di default
    EXCLUDE = {"domain users", "utenti del dominio"}
    filtered = [g for g in groups_series if g.strip() and g.strip().lower() not in EXCLUDE]

    filtered = sorted(set(filtered), key=lambda x: x.lower())
    if not filtered:
        return ""

    joined = ";".join(filtered)
    # Se ci sono spazi, racchiudi l'intera stringa nelle virgolette per sicurezza
    return f"\"{joined}\"" if any(" " in g for g in filtered) else joined


# ======================================================
# Helper: estrai nomi gruppi "generici" da un DataFrame
# ======================================================

def extract_group_names_from_df(df: pd.DataFrame) -> Set[str]:
    """
    Prova a estrarre nomi di gruppi da dataframe generici (DL/SM/MG).
    Cerca prima colonne 'GROUP/GROUPNAME/DISPLAYNAME', altrimenti 'NAME/NOME'.
    """
    if df is None or df.empty:
        return set()

    # 1) tenta group name esplicito
    group_col_name = _find_col(df, CAND_DL_GROUP + CAND_MG_GROUP + CAND_SM_GROUP_NAME)
    if group_col_name:
        series = df[group_col_name]
    else:
        # 2) fallback su 'name' / 'nome'
        fallback = _find_col(df, ["Name", "Nome", "DisplayName", "NomeVisualizzato"])
        if not fallback:
            return set()
        series = df[fallback]

    result = set([str(x).strip() for x in series.dropna().astype(str).tolist() if str(x).strip() != ""])
    return result


# =======================================================
# Helper: estrai gruppi Entra per user_email/UPN (EN/IT)
# =======================================================

def extract_entra_groups_for_user(entra_df: pd.DataFrame, user_email_or_upn: str) -> Set[str]:
    """
    Cerca i gruppi Entra a cui l'utente (UPN/email) appartiene.
    Supporta intestazioni EN/IT: MemberUserPrincipalName/UserPrincipalNameMembro + GroupName/NomeGruppo/DisplayName
    """
    if entra_df is None or entra_df.empty:
        return set()

    required = {
        "member_upn": CAND_ENTRA_MEMBER_UPN,
        "group_name": CAND_ENTRA_GROUP_NAME
    }
    ok, missing = _require_any(entra_df, required, "Entra")
    if not ok:
        st.warning(f"Nel file 'Entra' mancano i campi: {', '.join(missing)}")
        return set()

    member_col = _get_any(entra_df, CAND_ENTRA_MEMBER_UPN).astype(str).str.strip().str.lower()
    group_series = _get_any(entra_df, CAND_ENTRA_GROUP_NAME)

    target = user_email_or_upn.strip().lower()
    mask = member_col == target
    if not mask.any():
        return set()

    groups = _clean_series_to_list(group_series[mask])
    return set(groups)


# =========================================================
# Funzione testuale di deprovisioning (Step 2) – EN/IT safe
# =========================================================

def genera_deprovisioning(
    sam: str,
    dl_df: pd.DataFrame,
    sm_df: pd.DataFrame,
    mg_df: pd.DataFrame,
    entra_df: pd.DataFrame
) -> List[str]:
    sam_lower = sam.lower().strip()
    user_email = f"{sam_lower}@consip.it"

    clean = sam_lower.replace(".ext", "")
    parts = clean.split('.', 1)
    if sam_lower.endswith(".ext") and len(parts) == 2:
        nome, cognome = parts
        title = f"[Consip – SR] Casella di posta - Deprovisioning - {cognome.capitalize()} {nome.capitalize()} (esterno)"
    elif len(parts) == 2:
        nome, cognome = parts
        title = f"[Consip – SR] Casella di posta - Deprovisioning - {cognome.capitalize()} {nome.capitalize()}"
    else:
        title = f"[Consip – SR] Casella di posta - Deprovisioning - {clean}"

    st.subheader(title)
    lines = [f"Ciao,\nper {user_email} :"]
    warnings: List[str] = []
    step = 1

    fixed = [
        "Disabilitare invio ad utente (Message Delivery Restrictions)",
        "Impostare Hide dalla Rubrica",
        "Disabilitare accesso Mailbox (Mailbox features – Disable Protocolli/OWA)",
        f"Estrarre il PST (O365 eDiscovery) da archiviare in \\nasconsip2....\\backuppst\\03 - backup email cancellate\\{user_email} (in z7 con psw condivisa)",
        "Rimuovere i Ruoli assegnati",
        "Rimuovere le applicazioni dall’utenza Azure"
    ]
    for desc in fixed:
        lines.append(f"{step}. {desc}")
        step += 1

    # --- DL (Distribution Lists): rimozione abilitazione
    dl_list: List[str] = []
    if dl_df is not None and not dl_df.empty:
        col_group = _find_col(dl_df, CAND_DL_GROUP)
        col_mail = _find_col(dl_df, CAND_DL_EMAIL)
        if col_group and col_mail:
            mask = dl_df[col_mail].astype(str).str.strip().str.lower() == user_email
            if mask.any():
                dl_list = _clean_series_to_list(dl_df.loc[mask, col_group])
        else:
            warnings.append("Nel file DL non ho trovato colonne di 'Gruppo' o 'Email/SMTP'.")
    if dl_list:
        lines.append(f"{step}. Rimozione abilitazione dalle DL")
        for dl in dl_list:
            lines.append(f"   - {dl}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate DL per l'utente indicato")

    # --- Disabilita account Azure
    lines.append(f"{step}. Disabilitare l’account di Azure")
    step += 1

    # --- SM (Shared Mailboxes): rimozione abilitazioni
    sm_list: List[str] = []
    if sm_df is not None and not sm_df.empty:
        col_member_email = _find_col(sm_df, CAND_SM_MEMBER_EMAIL)
        col_group_name = _find_col(sm_df, CAND_SM_GROUP_NAME)
        if col_member_email and col_group_name:
            mvals = sm_df[col_member_email].astype(str).str.strip().str.lower()
            mask_any = mvals == user_email
            if mask_any.any():
                sm_list = _clean_series_to_list(sm_df.loc[mask_any, col_group_name])
        else:
            warnings.append("Nel file SM non ho trovato colonne di 'Member UPN/Email' o 'Nome mailbox/gruppo'.")
    if sm_list:
        lines.append(f"{step}. Rimozione abilitazione da SM")
        for sm in sm_list:
            lines.append(f"   - {sm}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate SM profilate all'utente indicato")

    # --- Azure (Entra) gruppi da rimuovere dopo scrematura DL/MG/SM
    entra_groups = extract_entra_groups_for_user(entra_df, user_email)
    dl_groups_all = extract_group_names_from_df(dl_df)
    mg_groups_all: Set[str] = set()
    if mg_df is not None and not mg_df.empty:
        mg_group_col = _find_col(mg_df, CAND_MG_GROUP)
        if mg_group_col:
            mg_groups_all = set([str(x).strip() for x in mg_df[mg_group_col].dropna().astype(str).tolist()])
    sm_groups_all = extract_group_names_from_df(sm_df)
    other_groups_union = set(dl_groups_all) | set(mg_groups_all) | set(sm_groups_all)

    def normalize_set(s):
        return {str(x).strip() for x in s if str(x).strip() != ""}

    entra_norm = normalize_set(entra_groups)
    other_norm = normalize_set(other_groups_union)
    other_lower = {g.lower() for g in other_norm}
    subset = {g for g in entra_norm if g.lower() not in other_lower}

    # Esclusioni specifiche note
    exclude_specific = {"o365 copilot plus", "o365 teams premium"}
    subset = {g for g in subset if g.lower() not in exclude_specific}

    if subset:
        lines.append(f"{step}. Rimozione gruppi Azure:")
        for g in sorted(subset, key=lambda x: x.lower()):
            lines.append(f"   - {g}")
        step += 1
    else:
        warnings.append("⚠️ Nessun gruppo Azure (file Entra) da rimuovere dopo lo scremamento con DL/MG/SM")

    final_rest = [
        "Cancellare la foto da Azure (se applicabile)",
        "Rimozione Wi-Fi"
    ]
    for desc in final_rest:
        lines.append(f"{step}. {desc}")
        step += 1

    if warnings:
        lines.append("\n⚠️ Avvisi:")
        lines.extend(warnings)

    return lines


# ==========================================
# Funzione per generare CSV Device – EN/IT
# ==========================================

def genera_device_csv(sam: str, device_df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    """
    Ritorna (contenuto_csv, nome_file) oppure (None, None) se non applicabile.
    Supporta colonne EN/IT tipiche degli export Device.
    """
    if device_df is None or device_df.empty:
        return None, None

    col_enabled = _find_col(device_df, CAND_DEV_ENABLED)
    if not col_enabled or col_enabled not in device_df.columns:
        return None, None

    # Filtra solo enabled == True/vero
    enabled_series = device_df[col_enabled]
    # Gestione valori tipo "True"/"Yes"/"Sì"
    enabled_mask = enabled_series.astype(str).str.strip().str.lower().isin(["true", "1", "yes", "si", "sì"])
    # se la colonna è nativa booleana, preserva
    if enabled_series.dtype == bool:
        enabled_mask = enabled_series == True

    device_df = device_df[enabled_mask]
    if device_df.empty:
        return None, None

    col_desc = _find_col(device_df, CAND_DEV_DESC)
    col_name = _find_col(device_df, CAND_DEV_NAME)
    col_mail = _find_col(device_df, CAND_DEV_MAIL)
    col_mobile = _find_col(device_df, CAND_DEV_MOBILE)
    col_upn = _find_col(device_df, CAND_DEV_UPN)

    if not col_desc or not col_name:
        return None, None

    # cerca riga del device collegata all'utente (in Description)
    sam_clean = sam.strip()
    desc_vals = device_df[col_desc].astype(str)
    mask = desc_vals.str.contains(fr"\s-\s{re.escape(sam_clean)}\s-\s", case=False, na=False)
    if not mask.any():
        return None, None

    row_data = device_df.loc[mask].iloc[0]
    computer_name = str(row_data[col_name]).strip()

    def _has_val(colname: Optional[str]) -> str:
        if not colname or colname not in row_data.index:
            return ""
        return "SI" if str(row_data.get(colname, "")).strip() else ""

    remove_mail = _has_val(col_mail)
    remove_mobile = _has_val(col_mobile)
    remove_upn = _has_val(col_upn)

    if not any([remove_mail, remove_mobile, remove_upn]):
        return None, None

    # Nome file: AAAAMMGG_Computer_riferimenti_remove[Cognome].csv
    today = datetime.now().strftime("%Y%m%d")
    clean = sam.replace(".ext", "")
    parts = clean.split('.')
    cognome = parts[1].capitalize() if len(parts) >= 2 else clean.capitalize()
    file_name = f"{today}_Computer_riferimenti_remove[{cognome}].csv"

    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar='\\')
    writer.writerow(HEADER_DEVICE)
    writer.writerow([computer_name, "", "", remove_mail, "", remove_mobile, "", remove_upn, "", ""])
    # Riga EOF
    writer.writerow(["EOF-riga lasciata appositamente scritta così per verificare che nessun PC sia andato nella OU/dismessi/computer"])
    buf.seek(0)
    return buf.getvalue(), file_name


# ===============
# Streamlit UI
# ===============

def main():
    st.set_page_config(page_title="Deprovisioning Consip", layout="centered")
    st.title("Deprovisioning Utente")

    sam = st.text_input("Nome utente (sAMAccountName)", "").strip().lower()
    st.markdown("---")

    if sam:
        clean = sam.replace(".ext", "")
        parts = clean.split('.')
        if len(parts) == 2:
            nome, cognome = parts
            csv_name = f"Deprovisioning_{cognome.capitalize()}_{nome[0].upper()}.csv"
        else:
            csv_name = f"Deprovisioning_{clean}.csv"
    else:
        csv_name = "Deprovisioning_.csv"
    st.write(f"**File CSV generato:** {csv_name}")

    dl_file = st.file_uploader("Carica file DL (Excel)", type="xlsx")
    sm_file = st.file_uploader("Carica file SM (Excel)", type="xlsx")
    mg_file = st.file_uploader("Carica file Estr_MembriGruppi (Excel)", type="xlsx")
    entra_file = st.file_uploader("Carica file Entra (Excel)", type="xlsx")
    device_file = st.file_uploader("Carica file Estr_Device (Excel)", type="xlsx")

    if st.button("Genera Template e CSV per Deprovisioning"):
        if not sam:
            st.error("Inserisci lo sAMAccountName")
            return

        dl_df = _read_excel_or_empty(dl_file)
        sm_df = _read_excel_or_empty(sm_file)
        mg_df = _read_excel_or_empty(mg_file)
        entra_df = _read_excel_or_empty(entra_file)
        device_df = _read_excel_or_empty(device_file)

        st.write("Colonne DL file:", dl_df.columns.tolist() if not dl_df.empty else "—")
        st.write("Colonne SM file:", sm_df.columns.tolist() if not sm_df.empty else "—")
        st.write("Colonne MG file:", mg_df.columns.tolist() if not mg_df.empty else "—")
        st.write("Colonne Entra file:", entra_df.columns.tolist() if not entra_df.empty else "—")
        st.write("Colonne Device file:", device_df.columns.tolist() if not device_df.empty else "—")

        # CSV Utente (Step 1)
        rimozione = estrai_rimozione_gruppi(sam, mg_df)
        row_map = {h: "" for h in HEADER_MODIFICA}
        row_map["sAMAccountName"] = sam
        row_map["RimozioneGruppo"] = rimozione
        row_map["disable"] = "SI"
        row_map["moveToOU"] = "SI"
        row = [row_map.get(h, "") for h in HEADER_MODIFICA]

        buf = io.StringIO()
        writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar='\\')
        writer.writerow(HEADER_MODIFICA)
        writer.writerow(row)
        buf.seek(0)

        preview_df = pd.read_csv(io.StringIO(buf.getvalue()), sep=",")
        st.subheader("Anteprima CSV Utente")
        st.dataframe(preview_df, use_container_width=True)
        st.download_button(label="📥 Scarica CSV Utente", data=buf.getvalue(), file_name=csv_name, mime="text/csv")

        # CSV Device (se presente)
        if device_file:
            device_csv, device_filename = genera_device_csv(sam, device_df)
            if device_csv:
                st.subheader("Anteprima CSV PC")
                preview_device_df = pd.read_csv(io.StringIO(device_csv), sep=",", header=None)
                st.dataframe(preview_device_df, use_container_width=True)
                st.download_button(label="📥 Scarica CSV PC", data=device_csv, file_name=device_filename, mime="text/csv")
            else:
                st.warning("Nessun dato valido per generare il CSV Device.")

        # Testo Deprovisioning (Step 2)
        steps = genera_deprovisioning(sam, dl_df, sm_df, mg_df, entra_df)
        st.subheader("Istruzioni Deprovisioning")
        st.text("\n".join(steps))


if __name__ == "__main__":
    main()
