import streamlit as st
import pandas as pd
import csv
import io
from datetime import datetime

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

# Funzione per comporre la stringa di rimozione gruppi per CSV
def estrai_rimozione_gruppi(sam_lower: str, mg_df: pd.DataFrame) -> str:
    if mg_df.empty:
        return ""
    member_col = next((c for c in mg_df.columns if "member" in c.lower()), None)
    group_col  = next((c for c in mg_df.columns if "group"  in c.lower()), None)
    if not member_col or not group_col:
        return ""
    mask = mg_df[member_col].astype(str).str.lower() == sam_lower
    groups_series = mg_df.loc[mask, group_col].dropna().astype(str)
    filtered = [g for g in groups_series if g.strip() and g.lower() != "domain users"]
    filtered = sorted(set(filtered), key=lambda x: x.lower())
    if not filtered:
        return ""
    joined = ";".join(filtered)
    return f"\"{joined}\"" if any(" " in g for g in filtered) else joined
# Helper: estrai nomi gruppi "generici" da un DataFrame
def extract_group_names_from_df(df: pd.DataFrame) -> set:
    if df.empty:
        return set()
    cols = df.columns.tolist()
    group_cols = [c for c in cols if 'group' in c.lower() and 'member' not in c.lower() and 'email' not in c.lower()]
    if not group_cols:
        group_cols = [c for c in cols if ('display' in c.lower() or 'name' in c.lower()) and 'member' not in c.lower() and 'email' not in c.lower()]
    result = set()
    for c in group_cols:
        result.update(df[c].dropna().astype(str).tolist())
    return set([r for r in result if str(r).strip() != ""])

# Helper: estrai gruppi dall'excel Entra per user_email
def extract_entra_groups_for_user(entra_df: pd.DataFrame, user_email: str) -> set:
    if entra_df.empty:
        return set()
    member_email_col = next((c for c in entra_df.columns if 'email' in c.lower()), None)
    group_col = next((c for c in entra_df.columns if 'group' in c.lower()), None)
    if not member_email_col or not group_col:
        return set()
    mask = entra_df[member_email_col].astype(str).str.lower() == user_email.lower()
    groups = entra_df.loc[mask, group_col].dropna().astype(str).tolist()
    return set([g for g in groups if g.strip() != ""])

# Funzione testuale di deprovisioning (Step 2)
def genera_deprovisioning(sam: str, dl_df: pd.DataFrame, sm_df: pd.DataFrame, mg_df: pd.DataFrame, entra_df: pd.DataFrame) -> list:
    sam_lower = sam.lower()
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
    warnings = []
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

    # DL
    dl_list = []
    if not dl_df.empty:
        group_col = next((c for c in dl_df.columns if 'distribution group' in c.lower()), None)
        smtp_col  = next((c for c in dl_df.columns if 'smtp' in c.lower() or 'email' in c.lower()), None)
        if group_col and smtp_col:
            mask = dl_df[smtp_col].astype(str).str.lower() == user_email
            dl_list = dl_df.loc[mask, group_col].dropna().unique().tolist()
    if dl_list:
        lines.append(f"{step}. Rimozione abilitazione dalle DL")
        for dl in sorted(dl_list):
            lines.append(f"   - {dl}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate DL all'utente indicato")

    # Disabilita account Azure
    lines.append(f"{step}. Disabilitare l’account di Azure")
    step += 1

    # SM
    sm_list = []
    if not sm_df.empty:
        member_col_sm = next((c for c in sm_df.columns if "member" in c.lower()), None)
        smtp_col_sm   = next((c for c in sm_df.columns if "email" in c.lower()), None)
        if member_col_sm and smtp_col_sm:
            sm_list = sm_df.loc[
                sm_df[member_col_sm].astype(str).str.lower() == user_email,
                smtp_col_sm
            ].dropna().unique().tolist()
        else:
            email_cols = [c for c in sm_df.columns if 'email' in c.lower()]
            group_cols = [c for c in sm_df.columns if 'group' in c.lower() or 'name' in c.lower() or 'display' in c.lower()]
            if email_cols and group_cols:
                mask_any = None
                for ec in email_cols:
                    mask = sm_df[ec].astype(str).str.lower() == user_email
                    mask_any = mask if mask_any is None else (mask_any | mask)
                if mask_any is not None and mask_any.any():
                    sm_list = sm_df.loc[mask_any, group_cols[0]].dropna().unique().tolist()
    if sm_list:
        lines.append(f"{step}. Rimozione abilitazione da SM")
        for sm in sorted(sm_list):
            lines.append(f"   - {sm}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate SM profilate all'utente indicato")

    entra_groups = extract_entra_groups_for_user(entra_df, user_email)
    dl_groups_all = extract_group_names_from_df(dl_df)
    mg_groups_all = set()
    if not mg_df.empty:
        group_col_mg_all = next((c for c in mg_df.columns if "group" in c.lower()), None)
        if group_col_mg_all:
            mg_groups_all = set(mg_df[group_col_mg_all].dropna().astype(str).tolist())
    sm_groups_all = extract_group_names_from_df(sm_df)
    other_groups_union = set(dl_groups_all) | set(mg_groups_all) | set(sm_groups_all)

    def normalize_set(s):
        return {str(x).strip() for x in s if str(x).strip() != ""}
    entra_norm = {g for g in normalize_set(entra_groups)}
    other_norm = {g for g in normalize_set(other_groups_union)}
    other_lower = {g.lower() for g in other_norm}
    subset = {g for g in entra_norm if g.lower() not in other_lower}
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

# Funzione per generare CSV Device
def genera_device_csv(sam: str, device_df: pd.DataFrame) -> (str, str):
    if device_df.empty or "Enabled" not in device_df.columns:
        return None, None
    device_df = device_df[device_df["Enabled"] == True]
    if device_df.empty:
        return None, None

    mask = device_df["Description"].astype(str).str.contains(f" - {sam} - ", case=False, na=False)
    if not mask.any():
        return None, None

    row_data = device_df.loc[mask].iloc[0]
    computer_name = str(row_data["Name"]).strip()
    remove_mail = "SI" if str(row_data.get("Mail", "")).strip() else ""
    remove_mobile = "SI" if str(row_data.get("Mobile", "")).strip() else ""
    remove_upn = "SI" if str(row_data.get("userPrincipalName", "")).strip() else ""

    if not any([remove_mail, remove_mobile, remove_upn]):
        return None, None

    # Nome file: AAAAMMGG_Computer_riferimenti_remuve[Cognome].csv
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

# Streamlit UI
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

        dl_df = pd.read_excel(dl_file) if dl_file else pd.DataFrame()
        sm_df = pd.read_excel(sm_file) if sm_file else pd.DataFrame()
        mg_df = pd.read_excel(mg_file) if mg_file else pd.DataFrame()
        entra_df = pd.read_excel(entra_file) if entra_file else pd.DataFrame()
        device_df = pd.read_excel(device_file) if device_file else pd.DataFrame()

        st.write("Colonne DL file:", dl_df.columns.tolist())
        st.write("Colonne SM file:", sm_df.columns.tolist())
        st.write("Colonne MG file:", mg_df.columns.tolist())
        st.write("Colonne Entra file:", entra_df.columns.tolist())
        st.write("Colonne Device file:", device_df.columns.tolist())

        # CSV Utente
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
        st.dataframe(preview_df)
        st.download_button(label="📥 Scarica CSV Utente", data=buf.getvalue(), file_name=csv_name, mime="text/csv")

        # CSV Device
        if device_file:
            device_csv, device_filename = genera_device_csv(sam, device_df)
            if device_csv:
                st.subheader("Anteprima CSV PC")
                preview_device_df = pd.read_csv(io.StringIO(device_csv), sep=",", header=None)
                st.dataframe(preview_device_df)
                st.download_button(label="📥 Scarica CSV PC", data=device_csv, file_name=device_filename, mime="text/csv")
            else:
                st.warning("Nessun dato valido per generare il CSV Device.")

        # Testo Deprovisioning
        steps = genera_deprovisioning(sam, dl_df, sm_df, mg_df, entra_df)
        st.text("\n".join(steps))

if __name__ == "__main__":
    main()
