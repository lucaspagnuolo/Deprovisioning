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

# Funzione per comporre la stringa di rimozione gruppi
def estrai_rimozione_gruppi(sam_lower: str, mg_df: pd.DataFrame) -> str:
    if mg_df.empty or mg_df.shape[1] <= 3:
        return ""
    mask = mg_df.iloc[:, 3].astype(str).str.lower() == sam_lower
    all_groups = mg_df.loc[mask, mg_df.columns[0]].dropna().tolist()
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
    if any(" " in g for g in base_groups):
        return f"\"{joined}\""
    return joined

# Funzione testuale di deprovisioning (Step 2)
def genera_deprovisioning(sam: str, dl_df: pd.DataFrame, sm_df: pd.DataFrame, mg_df: pd.DataFrame) -> list:
    sam_lower = sam.lower()

    # Generazione titolo
    clean = sam_lower.replace(".ext", "")
    if sam_lower.endswith(".ext"):
        parts = clean.split('.', 1)
        if len(parts) == 2:
            nome, cognome = parts
            title = f"[Consip – SR] Casella di posta - Deprovisioning - {cognome.capitalize()} {nome.capitalize()} (esterno)"
        else:
            title = f"[Consip – SR] Casella di posta - Deprovisioning - {clean} (esterno)"
    else:
        parts = clean.split('.', 1)
        if len(parts) == 2:
            nome, cognome = parts
            title = f"[Consip – SR] Casella di posta - Deprovisioning - {cognome.capitalize()} {nome.capitalize()}"
        else:
            title = f"[Consip – SR] Casella di posta - Deprovisioning - {clean}"

    dl_list = []
    if not dl_df.empty and dl_df.shape[1] > 5:
        mask = dl_df.iloc[:, 1].astype(str).str.lower() == sam_lower
        dl_list = dl_df.loc[mask, dl_df.columns[5]].dropna().tolist()
    sm_list = []
    if not sm_df.empty and sm_df.shape[1] > 2:
        target = f"{sam_lower}@consip.it"
        mask = sm_df.iloc[:, 2].astype(str).str.lower() == target
        sm_list = sm_df.loc[mask, sm_df.columns[0]].dropna().tolist()
    grp = []
    if not mg_df.empty and mg_df.shape[1] > 3:
        mask = mg_df.iloc[:, 3].astype(str).str.lower() == sam_lower
        grp = mg_df.loc[mask, mg_df.columns[0]].dropna().tolist()

    lines = [title, f"Ciao,\nper {sam_lower}@consip.it :"]
    warnings = []
    step = 2
    fixed = [
        "Disabilitare invio ad utente (Message Delivery Restrictions)",
        "Impostare Hide dalla Rubrica",
        "Disabilitare accesso Mailbox (Mailbox features – Disable Protocolli/OWA)",
        f"Estrarre il PST (O365 eDiscovery)...\\{sam_lower}@consip.it)",
        "Rimuovere le appartenenze dall’utenza Azure",
        "Rimuovere le applicazioni dall’utenza Azure"
    ]
    for desc in fixed:
        lines.append(f"{step}. {desc}")
        step += 1
    if dl_list:
        lines.append(f"{step}. Rimozione abilitazione dalle DL")
        for dl in dl_list:
            lines.append(f"   - {dl}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate DL all'utente indicato")

    lines.append(f"{step}. Disabilitare l’account di Azure")
    step += 1

    if sm_list:
        lines.append(f"{step}. Rimozione abilitazione da SM")
        for sm in sm_list:
            lines.append(f"   - {sm}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate SM profilate all'utente indicato")

    lines.append(f"{step}. Rimozione in AD del gruppo")
    lines.append("   - O365 Copilot Plus")
    lines.append("   - O365 Teams Premium")
    utenti_groups = [g for g in grp if g.lower().startswith("o365 utenti")]
    if utenti_groups:
        for g in utenti_groups:
            lines.append(f"   - {g}")
    else:
        warnings.append("⚠️ Non è stato trovato nessun gruppo O365 Utenti per l'utente")
    step += 1

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

    # Input sAMAccountName
    sam = st.text_input("Nome utente (sAMAccountName)", "").strip().lower()
    st.markdown("---")

    # Generazione automatica nome CSV
    if sam:
        clean = sam.replace(".ext", "")
        parts = clean.split('.')
        if len(parts) == 2:
            nome, cognome = parts
            cognome_fmt = cognome.capitalize()
            iniziale_nome = nome[0].upper() if nome else ''
            csv_name = f"Deprovisioning_{cognome_fmt}_{iniziale_nome}.csv"
        else:
            # formato senza punto: nomecognome
            csv_name = f"Deprovisioning_{clean}.csv"
    else:
        csv_name = "Deprovisioning_.csv"
    st.write(f"**File CSV generato:** {csv_name}")

    # File uploader
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

        # Anteprima CSV
        preview_df = pd.read_csv(io.StringIO(buf.getvalue()), sep=",")
        st.subheader("Anteprima CSV")
        st.dataframe(preview_df)

        # Download
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
