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

# Funzione per comporre la stringa di rimozione gruppi (esclude O365 Utenti per CSV)
def estrai_rimozione_gruppi(sam_lower: str, mg_df: pd.DataFrame) -> str:
    if mg_df.empty:
        return ""
    member_col = next((c for c in mg_df.columns if "member" in c.lower()), None)
    group_col  = next((c for c in mg_df.columns if "group"  in c.lower()), None)
    if not member_col or not group_col:
        return ""

    mask = mg_df[member_col].astype(str).str.lower() == sam_lower
    all_groups = mg_df.loc[mask, group_col].dropna().tolist()
    exclude = {"o365 copilot plus", "o365 teams premium", "domain users"}
    base_groups = []
    for g in all_groups:
        gl = g.lower()
        if gl.startswith("o365 utenti") or gl in exclude:
            continue
        base_groups.append(g)
    if not base_groups:
        return ""
    joined = ";".join(base_groups)
    return f"\"{joined}\"" if any(" " in g for g in base_groups) else joined

# Funzione testuale di deprovisioning (Step 2)
def genera_deprovisioning(sam: str, dl_df: pd.DataFrame, sm_df: pd.DataFrame, mg_df: pd.DataFrame) -> list:
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

    # DL: prova prima match diretto su SMTP, poi ricerca in colonna 'member(s)'
    dl_list = []
    if not dl_df.empty:
        # individua colonna SMTP e DisplayName e Members
        smtp_col   = next((c for c in dl_df.columns if "smtp" in c.lower()), None)
        name_col   = next((c for c in dl_df.columns if "displayname" in c.lower()), None)
        members_col= next((c for c in dl_df.columns if "member" in c.lower()), None)
        # match diretto su SMTP
        if smtp_col:
            dl_list = dl_df.loc[
                dl_df[smtp_col].astype(str).str.lower() == user_email,
                name_col or smtp_col
            ].dropna().tolist()
        # se nulla trovato, cerca in stringa membri
        if not dl_list and members_col:
            mask = dl_df[members_col].astype(str).str.lower().str.contains(user_email)
            dl_list = dl_df.loc[mask, name_col or smtp_col].dropna().tolist()
    if dl_list:
        lines.append(f"{step}. Rimozione abilitazione dalle DL")
        for dl in dl_list:
            lines.append(f"   - {dl}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate DL all'utente indicato")

    # Disabilita account Azure
    lines.append(f"{step}. Disabilitare l’account di Azure")
    step += 1

    # SM
    sm_list = []
    if not sm_df.empty and "Member" in sm_df.columns and "EmailAddress" in sm_df.columns:
        mask = sm_df["Member"].astype(str).str.lower() == user_email
        sm_list = sm_df.loc[mask, "EmailAddress"].dropna().tolist()
    if sm_list:
        lines.append(f"{step}. Rimozione abilitazione da SM")
        for sm in sm_list:
            lines.append(f"   - {sm}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate SM profilate all'utente indicato")

    # MG
    lines.append(f"{step}. Rimozione in AD del gruppo")
    lines.append("   - O365 Copilot Plus")
    lines.append("   - O365 Teams Premium")
    member_col = next((c for c in mg_df.columns if "member" in c.lower()), None)
    group_col  = next((c for c in mg_df.columns if "group"  in c.lower()), None)
    utenti_groups = []
    if member_col and group_col and not mg_df.empty:
        utenti_groups = [
            g for g in mg_df.loc[
                    mg_df[member_col].astype(str).str.lower() == sam_lower,
                    group_col
                ].dropna().tolist()
            if g.lower().startswith("o365 utenti")
        ]
    if utenti_groups:
        for g in utenti_groups:
            lines.append(f"   - {g}")
    else:
        warnings.append("⚠️ Non è stato trovato nessun gruppo O365 Utenti per l'utente")
    step += 1

    # Passi finali
    final = [
        "Disabilitazione utenza di dominio",
        "Spostamento in dismessi/utenti",
        "Cancellare la foto da Azure (se applicabile)",
        "Rimozione Wi-Fi"
    ]
    for desc in final:
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

    if st.button("Genera Template e CSV per Deprovisioning"):
        if not sam:
            st.error("Inserisci lo sAMAccountName")
            return

        dl_df = pd.read_excel(dl_file) if dl_file else pd.DataFrame()
        sm_df = pd.read_excel(sm_file) if sm_file else pd.DataFrame()
        mg_df = pd.read_excel(mg_file) if mg_file else pd.DataFrame()

        # DEBUG: mostra colonne per verifica
        st.write("Colonne DL file:", dl_df.columns.tolist())
        st.write("Colonne SM file:", sm_df.columns.tolist())
        st.write("Colonne MG file:", mg_df.columns.tolist())

        # Step 1: genera CSV
        rimozione = estrai_rimozione_gruppi(sam, mg_df)
        row = [
            sam if i == HEADER_MODIFICA.index("sAMAccountName") else
            rimozione if i == HEADER_MODIFICA.index("RimozioneGruppo") else
            ""
            for i in range(len(HEADER_MODIFICA))
        ]

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

        # Step 2: testo di deprovisioning
        steps = genera_deprovisioning(sam, dl_df, sm_df, mg_df)
        st.text("\n".join(steps))

if __name__ == "__main__":
    main()
