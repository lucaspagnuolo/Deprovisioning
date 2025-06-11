# app_deprovisioning.py
import streamlit as st
import pandas as pd

# Funzione principale che genera la lista dei passaggi di deprovisioning
def genera_deprovisioning(sam: str, dl_df: pd.DataFrame, sm_df: pd.DataFrame, mg_df: pd.DataFrame) -> list:
    sam_lower = sam.lower()
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

    lines = [f"Ciao,\nper {sam_lower}@consip.it :"]
    warnings = []
    step = 1

    # Passaggi fissi
    lines.extend([
        f"{step}. Disabilitare invio ad utente (Message Delivery Restrictions)",
        f"{step+1}. Impostare Hide dalla Rubrica",
        f"{step+2}. Disabilitare accesso Mailbox (Mailbox features – Disable Protocolli/OWA)",
        f"{step+3}. Estrarre il PST (O365 eDiscovery) da archiviare in \\nasconsip2....\\backuppst\\03 - backup email cancellate\\{sam_lower}@consip.it (in z7 con psw condivisa)",
        f"{step+4}. Rimuovere le appartenenze dall’utenza Azure",
        f"{step+5}. Rimuovere le applicazioni dall’utenza Azure"
    ])
    step += 6

    # Rimozione DL
    if dl_list:
        lines.append(f"{step}. Rimozione abilitazione dalle DL")
        for dl in dl_list:
            lines.append(f"   - {dl}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate DL all'utente indicato")

    lines.append(f"{step}. Disabilitare l’account di Azure")
    step += 1

    # Rimozione SM
    if sm_list:
        lines.append(f"{step}. Rimozione abilitazione da SM")
        for sm in sm_list:
            lines.append(f"   - {sm}")
        step += 1
    else:
        warnings.append("⚠️ Non sono state trovate SM profilate all'utente indicato")

    # Rimozione gruppi in AD
    lines.append(f"{step}. Rimozione in AD del gruppo")
    lines.append("   - O365 Copilot Plus")
    lines.append("   - O365 Teams Premium")
    # Aggiungo VivaEngage se non utente esterno
    if not sam_lower.endswith(".ext"):
        lines.append("   - O365 VivaEngage")

    # Gruppi O365 Utenti personalizzati
    utenti_groups = [g for g in grp if g.lower().startswith("o365 utenti")]
    if utenti_groups:
        for g in utenti_groups:
            lines.append(f"   - {g}")
    else:
        warnings.append("⚠️ Non è stato trovato nessun gruppo O365 Utenti per l'utente")
    step += 1

    # Ultimi step
    lines.extend([
        f"{step}. Disabilitazione utenza di dominio",
        f"{step+1}. Spostamento in dismessi/utenti",
        f"{step+2}. Cancellare la foto da Azure (se applicabile)",
        f"{step+3}. Rimozione Wi-Fi"
    ])

    # Avvisi finali
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

    dl_file = st.file_uploader("Carica file DL (Excel)", type="xlsx")
    sm_file = st.file_uploader("Carica file SM (Excel)", type="xlsx")
    mg_file = st.file_uploader("Carica file Membri Gruppi (Excel)", type="xlsx")

    if st.button("Genera Deprovisioning"):
        if not sam:
            st.error("Inserisci lo sAMAccountName")
            return

        dl_df = pd.read_excel(dl_file) if dl_file else pd.DataFrame()
        sm_df = pd.read_excel(sm_file) if sm_file else pd.DataFrame()
        mg_df = pd.read_excel(mg_file) if mg_file else pd.DataFrame()

        steps = genera_deprovisioning(sam, dl_df, sm_df, mg_df)
        st.text("\n".join(steps))

if __name__ == "__main__":
    main()
