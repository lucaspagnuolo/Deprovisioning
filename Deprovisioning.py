import streamlit as st
import pandas as pd
import csv
import io

# Header CSV per lo Step 1
HEADER_MODIFICA = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

# Funzione per comporre la stringa di rimozione gruppi per CSV:
# >>> ORA: include tutti i gruppi trovati in MembriGruppi per l'utente, ESCLUSO solo "Domain Users"
def estrai_rimozione_gruppi(sam_lower: str, mg_df: pd.DataFrame) -> str:
    if mg_df.empty:
        return ""
    # Trova colonne di membership e group
    member_col = next((c for c in mg_df.columns if "member" in c.lower()), None)
    group_col  = next((c for c in mg_df.columns if "group"  in c.lower()), None)
    if not member_col or not group_col:
        return ""
    # Filtra gruppi dell'utente
    mask = mg_df[member_col].astype(str).str.lower() == sam_lower
    groups_series = mg_df.loc[mask, group_col].dropna().astype(str)

    # Escludi soltanto "Domain Users" (case-insensitive); rimuovi vuoti e deduplica
    filtered = [g for g in groups_series if g.strip() and g.lower() != "domain users"]
    filtered = sorted(set(filtered), key=lambda x: x.lower())

    if not filtered:
        return ""
    joined = ";".join(filtered)
    # Manteniamo la logica di quoting se esistono gruppi con spazi (compatibile con il writer QUOTE_NONE)
    return f"\"{joined}\"" if any(" " in g for g in filtered) else joined

# Helper: estrai nomi gruppi "generici" da un DataFrame (usa colonne probabili)
def extract_group_names_from_df(df: pd.DataFrame) -> set:
    if df.empty:
        return set()
    cols = df.columns.tolist()
    # prefer columns containing 'group' but not 'member'/'email'
    group_cols = [c for c in cols if 'group' in c.lower() and 'member' not in c.lower() and 'email' not in c.lower()]
    if not group_cols:
        # fallback: 'display' or 'name' escludendo member/email
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

    # Generazione titolo
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

    # Passi fissi iniziali
    fixed = [
        "Disabilitare invio ad utente (Message Delivery Restrictions)",
        "Impostare Hide dalla Rubrica",
        "Disabilitare accesso Mailbox (Mailbox features – Disable Protocolli/OWA)",
        f"Estrarre il PST (O365 eDiscovery) da archiviare in \\nasconsip2....\\backuppst\\03 - backup email cancellate\\{user_email} (in z7 con psw condivisa)",
        "Rimuovere le appartenenze dall’utenza Azure",
        "Rimuovere le applicazioni dall’utenza Azure"
    ]
    for desc in fixed:
        lines.append(f"{step}. {desc}")
        step += 1

    # Step: estrazione delle DL da dl_df
    dl_list = []
    if not dl_df.empty:
        display_col = next((c for c in dl_df.columns if 'display' in c.lower()), None) or next((c for c in dl_df.columns if 'name' in c.lower()), None)
        smtp_col    = next((c for c in dl_df.columns if 'smtp' in c.lower() or 'email' in c.lower()), None)
        if display_col and smtp_col:
            mask = dl_df[smtp_col].astype(str).str.lower() == user_email
            dl_list = dl_df.loc[mask, display_col].dropna().unique().tolist()
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

    # --- Prepara insiemi per il confronto con Entra ---
    entra_groups = extract_entra_groups_for_user(entra_df, user_email)

    # gruppi presenti nei 3 file (DL, MG, SM) per escluderli dall'Entra
    dl_groups_all = extract_group_names_from_df(dl_df)
    mg_groups_all = set()
    if not mg_df.empty:
        group_col_mg_all = next((c for c in mg_df.columns if "group" in c.lower()), None)
        if group_col_mg_all:
            mg_groups_all = set(mg_df[group_col_mg_all].dropna().astype(str).tolist())
    sm_groups_all = extract_group_names_from_df(sm_df)

    other_groups_union = set([g for g in dl_groups_all if str(g).strip() != ""]) | set([g for g in mg_groups_all if str(g).strip() != ""]) | set([g for g in sm_groups_all if str(g).strip() != ""])

    def normalize_set(s):
        return {str(x).strip() for x in s if str(x).strip() != ""}
    entra_norm = {g for g in normalize_set(entra_groups)}
    other_norm = {g for g in normalize_set(other_groups_union)}
    other_lower = {g.lower() for g in other_norm}
    subset = {g for g in entra_norm if g.lower() not in other_lower}

    # Escludi specifici gruppi da Azure (come prima, solo per il punto "Rimozione gruppi Azure")
    exclude_specific = {"o365 copilot plus", "o365 teams premium"}
    subset = {g for g in subset if g.lower() not in exclude_specific}

    # Disabilitazione utenza di dominio
    lines.append(f"{step}. Disabilitazione utenza di dominio")
    step += 1

    # Nuovo punto: Rimozione gruppi Azure (se presenti)
    if subset:
        lines.append(f"{step}. Rimozione gruppi Azure:")
        for g in sorted(subset, key=lambda x: x.lower()):
            lines.append(f"   - {g}")
        step += 1
    else:
        warnings.append("⚠️ Nessun gruppo Azure (file Entra) da rimuovere dopo lo scremamento con DL/MG/SM")

    # *** RIMOSSO: "Rimozione in AD del gruppo" + O365 Copilot/Teams Premium + O365 Utenti ***
    # *** RIMOSSO: "Spostamento in dismessi/utenti" ***

    # Altri passi finali (senza lo spostamento)
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

# Streamlit UI
def main():
    st.set_page_config(page_title="Deprovisioning Consip", layout="centered")
    st.title("Deprovisioning Utente")

    sam = st.text_input("Nome utente (sAMAccountName)", "").strip().lower()
    st.markdown("---")

    # Genera il nome del CSV
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
    entra_file = st.file_uploader("Carica file Entra (Excel) - colonne: GroupName | MemberName | MemberEmail", type="xlsx")

    if st.button("Genera Template e CSV per Deprovisioning"):
        if not sam:
            st.error("Inserisci lo sAMAccountName")
            return

        dl_df = pd.read_excel(dl_file) if dl_file else pd.DataFrame()
        sm_df = pd.read_excel(sm_file) if sm_file else pd.DataFrame()
        mg_df = pd.read_excel(mg_file) if mg_file else pd.DataFrame()
        entra_df = pd.read_excel(entra_file) if entra_file else pd.DataFrame()

        # DEBUG: mostra colonne per verifica
        st.write("Colonne DL file:", dl_df.columns.tolist())
        st.write("Colonne SM file:", sm_df.columns.tolist())
        st.write("Colonne MG file:", mg_df.columns.tolist())
        st.write("Colonne Entra file:", entra_df.columns.tolist())

        # Step 1: genera CSV
        rimozione = estrai_rimozione_gruppi(sam, mg_df)

        # Costruiamo la riga in modo esplicito per chiarezza
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
        st.subheader("Anteprima CSV")
        st.dataframe(preview_df)

        st.download_button(
            label="📥 Scarica CSV",
            data=buf.getvalue(),
            file_name=csv_name,
            mime="text/csv"
        )

        # Step 2: testo di deprovisioning (con modifiche richieste)
        steps = genera_deprovisioning(sam, dl_df, sm_df, mg_df, entra_df)
        st.text("\n".join(steps))

if __name__ == "__main__":
    main()
