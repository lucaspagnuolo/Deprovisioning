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
    if mg_df.empty or mg_df.shape <= 3:
        return ""
    mask = mg_df.iloc[:, 3].astype(str).str.lower() == sam_lower
    all_groups = mg_df.loc[mask, mg_df.columns].dropna().tolist()
    exclude = {"o365 copilot plus", "o365 teams premium", "domain users"}
    base_groups = [g for g in all_groups if not (g.lower().startswith("o365 utenti") or g.lower() in exclude)]
    if not base_groups:
        return ""
    joined = ";".join(base_groups)
    return f"\"{joined}\"" if any(" " in g for g in base_groups) else joined

# Funzione testuale di deprovisioning (Step 2)
def genera_deprovisioning(sam: str, dl_df: pd.DataFrame, sm_df: pd.DataFrame, mg_df: pd.DataFrame) -> list:
    sam_lower = sam.lower()
    
    # Generazione del titolo condizionale
    if '.ext' in sam:
        title = f"[Consip – SR] Casella di posta - Deprovisioning - {sam.replace('.ext', '').split('.').capitalize()} {sam.replace('.ext', '').split('.').capitalize()} (esterno)"
    elif '.' in sam:
        title = f"[Consip – SR] Casella di posta - Deprovisioning - {sam.split('.').capitalize()} {sam.split('.').capitalize()}"
    else:
        title = f"[Consip – SR] Casella di posta - Deprovisioning - {sam.capitalize()}"
    
    dl_list = dl_df.loc[dl_df.iloc[:, 1].astype(str).str.lower() == sam_lower, dl_df.columns].dropna().tolist() if not dl_df.empty and dl_df.shape > 5 else []
    sm_list = sm_df.loc[sm_df.iloc[:, 2].astype(str).str.lower() == f"{sam_lower}@consip.it", sm_df.columns].dropna().tolist() if not sm_df.empty and sm_df.shape > 2 else []
    grp = mg_df.loc[mg_df.iloc[:, 3].astype(str).str.lower() == sam_lower, mg_df.columns].dropna().tolist() if not mg_df.empty and mg_df.shape > 3 else []

    lines = [title, f"Ciao,\nper {sam_lower}@consip.it :"]
    warnings = []
    step = 2
    fixed_steps = [
        "Disabilitare invio ad utente (Message Delivery Restrictions)",
        "Impostare Hide dalla Rubrica",
        "Disabilitare accesso Mailbox (Mailbox features – Disable Protocolli/OWA)",
        f"Estrarre il PST (O365 eDiscovery)...\\{sam_lower}@consip.it)",
        "Rimuovere le appartenenze dall’utenza Azure",
        "Rimuovere le applicazioni dall’utenza Azure"
    ]
    
    for desc in fixed_steps:
        lines.append(f"{step}. {desc}")
        step += 1
    
    if dl_list:
        lines.append(f"{step}. Rimozione abilitazione dalle DL")
        lines.extend([f"   - {dl}" for dl in dl_list])
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate DL all'utente indicato")

    lines.append(f"{step}. Disabilitare l’account di Azure")
    step += 1

    if sm_list:
        lines.append(f"{step}. Rimozione abilitazione da SM")
        lines.extend([f"   - {sm}" for sm in sm_list])
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate SM profilate all'utente indicato")

    lines.append(f"{step}. Rimozione in AD del gruppo")
    lines.extend(["   - O365 Copilot Plus", "   - O365 Teams Premium"])
    
    utenti_groups = [g for g in grp if g.lower().startswith("o365 utenti")]
    if utenti_groups:
        lines.extend([f"   - {g}" for g in utenti_groups])
    else:
        warnings.append("⚠️ Non è stato trovato nessun gruppo O365 Utenti per l'utente")
    
    step += 1

    final_steps = [
        "Disabilitazione utenza di dominio",
        "Spostamento in dismessi/utenti",
        "Cancellare la foto da Azure (se applicabile)",
        "Rimozione Wi-Fi"
    ]
    
    for desc in final_steps:
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
    csv_name = f"Deprovisioning_{sam.replace('.ext', '').split('.').capitalize()}_{sam.replace('.ext', '').split('.').upper()}.csv" if sam and '.' in sam else f"Deprovisioning_{sam.replace('.ext', '')}.csv"
    
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
