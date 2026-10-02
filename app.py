import math
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# KONTAKTDATEN (Hier deine Daten anpassen)
# ==============================================================================
PROVIDE_BY_NAME = "Alexander Paul"
CONTACT_EMAIL = "alexander.paul.edelmetalle@gmail.com"
CONTACT_PHONE = "+49 1573 3729040"  


# ==============================================================================
# SEITEN-KONFIGURATION & LUXURIÖSES GOLDHÄNDLER-DESIGN (CSS)
# ==============================================================================

st.set_page_config(
    page_title="Edelmetall- & Ertragsrechner",
    page_icon="👑",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS für ein exklusives, matt-schwarzes Edelmetall-Design
CUSTOM_CSS = """
<style>
    /* Haupthintergrund: Deep Obsidian / Mattes Tiefschwarz */
    .stApp {
        background-color: #121212;
        color: #F3E5AB; /* Sanftes Champagner-Gold als Text */
        font-family: 'Inter', sans-serif;
    }
    
    /* Sub-Header & Kontaktbereich Styling */
    .provider-container {
        display: flex;
        align-items: center;
        gap: 15px;
        flex-wrap: wrap;
        margin-top: 15px; /* Mehr Abstand zum Titel oben */
        margin-bottom: 30px;
        padding: 12px 20px;
        background-color: #1A1A1A;
        border-left: 3px solid #D4AF37;
        border-radius: 6px;
    }
    .provider-text {
        color: #C0C0C0;
        font-size: 0.95rem;
    }
    .provider-link {
        color: #D4AF37 !important;
        text-decoration: none;
        font-weight: 600;
        transition: color 0.2s ease;
    }
    .provider-link:hover {
        text-decoration: underline;
        color: #F5E6AD !important;
    }

    /* Eingabe-Container / Karten in edlem Graphit/Schwarz mit Goldrand */
    div[data-testid="stForm"], div.css-1r6slbk, .financial-card {
        background-color: #1A1A1A;
        border: 1px solid #D4AF37; /* Echtes Gold */
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 20px rgba(212, 175, 55, 0.12);
    }

    /* KPI-Karten (Metrics) im dunklen Premium-Look */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #222222 0%, #161616 100%);
        border: 1px solid #9A7B38;
        border-radius: 10px;
        padding: 18px 22px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.6);
    }
    div[data-testid="stMetricLabel"] {
        color: #C0C0C0 !important; /* Silberne Beschriftung */
        font-size: 0.85rem !important;
        font-weight: 500 !important;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }
    div[data-testid="stMetricValue"] {
        color: #F3E5AB !important; /* Champagner-Gold für Beträge */
        font-weight: 700 !important;
        font-size: 1.8rem !important;
    }

    /* Input-Felder mit neutraler dunkler Basis */
    .stNumberInput input, .stSelectbox div[data-baseweb="select"] {
        background-color: #121212 !important;
        color: #F8FAFC !important;
        border-radius: 8px !important;
        border: 1px solid #333333 !important;
    }
    .stNumberInput input:focus, .stSelectbox div[data-baseweb="select"]:focus {
        border-color: #D4AF37 !important;
    }

    /* Slider Akzentfarbe */
    div[data-baseweb="slider"] div {
        background-color: #D4AF37 !important;
    }

    /* Styling für Download-Button (Goldene Optik) */
    div[data-testid="stDownloadButton"] button {
        background-color: #1A1A1A !important;
        color: #D4AF37 !important;
        border: 1px solid #D4AF37 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease-in-out;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #D4AF37 !important;
        color: #121212 !important;
        border-color: #D4AF37 !important;
    }

    /* Überschriften */
    h1, h2, h3 {
        color: #F5E6AD !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px;
    }
    
    /* Subtitle / Captions */
    .stCaption {
        color: #999999 !important;
    }

    /* Trennlinie */
    hr {
        border-color: #332A15 !important;
    }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==============================================================================
# BERECHNUNGSLOGIK & EXPORT-HELFER
# ==============================================================================

def calculate_compound_interest(
    initial: float, 
    monthly: float, 
    rate_annual: float, 
    total_months: int,
    time_unit: str
) -> pd.DataFrame:
    """
    Berechnet die Vermögensentwicklung über die angegebene Laufzeit in Monaten.
    """
    rate_monthly = rate_annual / 100 / 12
    data = []

    for m in range(total_months + 1):
        if rate_monthly > 0:
            total = initial * (1 + rate_monthly) ** m + monthly * (
                ((1 + rate_monthly) ** m - 1) / rate_monthly
            )
        else:
            total = initial + monthly * m

        deposits = initial + monthly * m
        interest = total - deposits

        years = round(m / 12, 2)
        time_label = years if time_unit == "Jahre" else m

        data.append({
            "Zeitschritt": time_label,
            "Laufzeit (Jahre)": years,
            "Laufzeit (Monate)": int(m),
            "Eingezahltes Kapital": float(deposits),
            "Zinsgewinn": float(interest),
            "Gesamtguthaben": float(total)
        })

    return pd.DataFrame(data)


@st.cache_data
def convert_df_to_excel_csv(df: pd.DataFrame) -> bytes:
    """
    Konvertiert das DataFrame in eine CSV-Datei mit deutschem Excel-Standard:
    - Semikolon (;) als Spaltentrenner
    - Komma (,) als Dezimaltrenner
    - UTF-8 mit BOM (utf-8-sig), damit Umlaute in Excel direkt stimmen.
    """
    export_df = df[[
        "Laufzeit (Jahre)", 
        "Laufzeit (Monate)", 
        "Eingezahltes Kapital", 
        "Zinsgewinn", 
        "Gesamtguthaben"
    ]].copy()

    return export_df.to_csv(
        sep=";", 
        decimal=",", 
        index=False, 
        encoding="utf-8-sig"
    ).encode("utf-8-sig")


# ==============================================================================
# VISUALISIERUNGEN (Plotly Chart & Plotly Custom Table)
# ==============================================================================

def create_financial_chart(df: pd.DataFrame, time_unit: str) -> go.Figure:
    """
    Erstellt ein elegantes Plotly-Area-Chart in Gold und Silber auf tiefschwarzem Grund.
    """
    fig = go.Figure()

    # Gesamtguthaben / Zinsgewinn -> GOLD
    fig.add_trace(go.Scatter(
        x=df["Zeitschritt"],
        y=df["Gesamtguthaben"],
        name="Gesamtwert (inkl. Ertrag)",
        mode="lines",
        fill="tozeroy",
        line=dict(color="#D4AF37", width=3),
        fillcolor="rgba(212, 175, 55, 0.15)",
        hovertemplate="<b>Gesamtwert:</b> %{y:,.2f} €<extra></extra>"
    ))

    # Eigenes eingezahltes Kapital -> SILBER
    fig.add_trace(go.Scatter(
        x=df["Zeitschritt"],
        y=df["Eingezahltes Kapital"],
        name="Eigenkapital (Investition)",
        mode="lines",
        fill="tozeroy",
        line=dict(color="#C0C0C0", width=2),
        fillcolor="rgba(192, 192, 192, 0.08)",
        hovertemplate="<b>Investiert:</b> %{y:,.2f} €<extra></extra>"
    ))

    fig.update_layout(
        paper_bgcolor="#1A1A1A",
        plot_bgcolor="#1A1A1A",
        font=dict(color="#A0A0A0", family="Inter, sans-serif"),
        title=dict(
            text="Prognostizierte Wertentwicklung",
            font=dict(size=18, color="#F3E5AB")
        ),
        xaxis=dict(
            title=f"Laufzeit ({time_unit})",
            gridcolor="#2A2A2A",
            showline=True,
            linecolor="#4A3E22",
            zeroline=False
        ),
        yaxis=dict(
            title="Wert in Euro (€)",
            gridcolor="#2A2A2A",
            showline=True,
            linecolor="#4A3E22",
            zeroline=False
        ),
        hovermode="x unified",
        margin=dict(l=20, r=20, t=50, b=30),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color="#F3E5AB")
        )
    )

    return fig


def create_gold_table(df: pd.DataFrame) -> go.Figure:
    """
    Erstellt eine komplett benutzerdefinierte Plotly-Tabelle im perfekten Gold/Silber-Look.
    """
    def fmt_euro(val):
        return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

    years_col = [f"{v:.2f} Jahre" for v in df["Laufzeit (Jahre)"]]
    months_col = [f"{int(v)} Mon." for v in df["Laufzeit (Monate)"]]
    capital_col = [fmt_euro(v) for v in df["Eingezahltes Kapital"]]
    interest_col = [fmt_euro(v) for v in df["Zinsgewinn"]]
    total_col = [fmt_euro(v) for v in df["Gesamtguthaben"]]

    num_rows = len(df)
    row_colors = ["#1A1A1A" if i % 2 == 0 else "#161616" for i in range(num_rows)]

    fig = go.Figure(data=[go.Table(
        header=dict(
            values=[
                "<b>Laufzeit (Jahre)</b>", 
                "<b>Laufzeit (Monate)</b>", 
                "<b>Eigenkapital</b>", 
                "<b>Zinsgewinn</b>", 
                "<b>Gesamtguthaben</b>"
            ],
            fill_color="#222222",
            align=["center", "center", "right", "right", "right"],
            font=dict(color="#D4AF37", size=13, family="Inter, sans-serif"),
            line_color="#4A3E22",
            height=38
        ),
        cells=dict(
            values=[years_col, months_col, capital_col, interest_col, total_col],
            fill_color=[row_colors * 5],
            align=["center", "center", "right", "right", "right"],
            font=dict(
                color=["#A0A0A0", "#A0A0A0", "#C0C0C0", "#9A7B38", "#F3E5AB"], 
                size=12,
                family="Inter, sans-serif"
            ),
            line_color="#2A2A2A",
            height=30
        )
    )])

    fig.update_layout(
        paper_bgcolor="#1A1A1A",
        margin=dict(l=0, r=0, t=10, b=10)
    )

    return fig


# ==============================================================================
# CALLBACKS FÜR LOKALE ZUSTANDSSYNCHRONISATION
# ==============================================================================

def sync_duration_input():
    st.session_state.duration_val = st.session_state.duration_input

def sync_duration_slider():
    st.session_state.duration_val = st.session_state.duration_slider


# ==============================================================================
# STREAMLIT APP MAIN
# ==============================================================================

def main():
    # Header & Bereitgestellt-Unterzeile
    st.title("👑 Edelmetall- & Ertragsrechner")
    
    # Exklusiver Kontakt-Header unter dem Titel
    st.markdown(
        f"""
        <div class="provider-container">
            <span class="provider-text">
                Bereitgestellt von <strong>{PROVIDE_BY_NAME}</strong>. Buchen Sie gerne ein persönliches Beratungsgespräch:
            </span>
            <a class="provider-link" href="mailto:{CONTACT_EMAIL}">✉ {CONTACT_EMAIL}</a>
            <span class="provider-text">|</span>
            <a class="provider-link" href="tel:{CONTACT_PHONE.replace(' ', '')}">📞 {CONTACT_PHONE}</a>
        </div>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------------------------
    # EINGABEPARAMETER
    # --------------------------------------------------------------------------
    st.subheader("⚙️ Investitionsparameter")
    
    with st.container():
        # Zeile 1: Startkapital, Monatlicher Beitrag, Rendite
        row1_col1, row1_col2, row1_col3 = st.columns([2, 2, 1])
        
        with row1_col1:
            initial_capital = st.number_input(
                "Anfangsinvestition (€)",
                min_value=0.0,
                max_value=5_000_000.0,
                value=0.0,
                step=500.0,
                format="%.2f"
            )

        with row1_col2:
            monthly_savings = st.number_input(
                "Monatlicher Ansparbetrag (€)",
                min_value=0.0,
                max_value=100_000.0,
                value=250.0,
                step=50.0,
                format="%.2f"
            )

        with row1_col3:
            annual_rate = st.number_input(
                "Erwartete Rendite p.a. (%)",
                min_value=0.0,
                max_value=100.0,
                value=11.0,
                step=0.1,
                format="%.2f"
            )

        # Zeile 2: Zeiteinheit & Laufzeit
        row2_col1, row2_col2, row2_col3 = st.columns([1, 2, 1])
        
        with row2_col1:
            time_unit = st.radio(
                "Zeiteinheit",
                options=["Monate", "Jahre"],
                index=1,
                horizontal=True
            )

        max_duration = 50 if time_unit == "Jahre" else 120
        default_duration = 10 if time_unit == "Jahre" else 120

        if "duration_val" not in st.session_state:
            st.session_state.duration_val = default_duration

        if st.session_state.duration_val > max_duration:
            st.session_state.duration_val = max_duration

        with row2_col2:
            st.slider(
                f"Laufzeit-Regler ({time_unit})",
                min_value=1,
                max_value=max_duration,
                value=st.session_state.duration_val,
                key="duration_slider",
                on_change=sync_duration_slider
            )

        with row2_col3:
            st.number_input(
                f"Laufzeit-Fixwert ({time_unit})",
                min_value=1,
                max_value=1000,
                value=st.session_state.duration_val,
                step=1,
                key="duration_input",
                on_change=sync_duration_input
            )

        duration_val = st.session_state.duration_val
        total_months = duration_val * 12 if time_unit == "Jahre" else duration_val

    st.divider()

    # --------------------------------------------------------------------------
    # BERECHNUNG & KPI METRICS
    # --------------------------------------------------------------------------
    df = calculate_compound_interest(
        initial=initial_capital,
        monthly=monthly_savings,
        rate_annual=annual_rate,
        total_months=total_months,
        time_unit=time_unit
    )

    final_row = df.iloc[-1]
    total_deposited = final_row["Eingezahltes Kapital"]
    final_balance = final_row["Gesamtguthaben"]
    interest_earned = final_row["Zinsgewinn"]

    st.subheader("📊 Wertentwicklungs-Übersicht")

    kpi1, kpi2, kpi3 = st.columns(3)

    with kpi1:
        st.metric(
            label="Gesamtes Investiertes Kapital (Silber)",
            value=f"{total_deposited:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    with kpi2:
        st.metric(
            label="Zuwachs / Wertsteigerung (Gold)",
            value=f"{interest_earned:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."),
            delta=f"+{interest_earned:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".") if interest_earned > 0 else None
        )

    with kpi3:
        st.metric(
            label=f"Prognostizierter Endwert ({duration_val} {time_unit})",
            value=f"{final_balance:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    st.write("")

    # --------------------------------------------------------------------------
    # DIAGRAMM
    # --------------------------------------------------------------------------
    fig_chart = create_financial_chart(df, time_unit)
    st.plotly_chart(fig_chart, use_container_width=True)

    # --------------------------------------------------------------------------
    # DETAILLIERTER VERLAUFSPLAN (Plotly-Tabelle mit Excel-Download)
    # --------------------------------------------------------------------------
    with st.expander("📄 Detaillierten Verlaufsplan anzeigen"):
        # Excel-kompatible CSV-Datei erzeugen
        csv_data = convert_df_to_excel_csv(df)

        col_tbl_title, col_dl = st.columns([3, 1])
        with col_dl:
            st.download_button(
                label="📥 Tabelle für Excel herunterladen",
                data=csv_data,
                file_name=f"verlaufsplan_{duration_val}_{time_unit.lower()}.csv",
                mime="text/csv",
                use_container_width=True
            )

        # Visualisierung der dunklen Gold-Tabelle
        fig_table = create_gold_table(df)
        st.plotly_chart(fig_table, use_container_width=True)

    st.caption("*Hinweis: Die Modellrechnung dient zu Anschauungszwecken. Alle Werte ohne Gewähr.")


if __name__ == "__main__":
    main()


# --------------------------------------------------------------------------
    # FOOTER: RECHTSHINWEISE (Homogenes & einheitliches Design)
    # --------------------------------------------------------------------------
    st.write("<br><br><br><br><br>", unsafe_allow_html=True)
    st.divider()

    # CSS-Styling für den dezente Optik des Expanders
    st.markdown("""
        <style>
            /* Expander-Titel unaufdringlich gestalten */
            div[data-testid="stExpander"] details summary span p {
                font-size: 0.8rem !important;
                color: #666666 !important;
                font-weight: 400 !important;
            }
            /* Sehr dunkler Rahmen für den Hintergrund */
            div[data-testid="stExpander"] {
                border: 1px solid #222222 !important;
                border-radius: 8px !important;
                background-color: #121212 !important;
            }
        </style>
    """, unsafe_allow_html=True)

    with st.expander("⚖️ Impressum & Datenschutzerklärung"):
        st.markdown("""
        #### 📜 Impressum
        
        **Angaben gemäß § 5 DDG:**  
        Alexander Paul  
        Vischerstr. 2   
        70563 Stuttgart 

        **Kontakt:**  
        E-Mail: [alexander.paul.edelmetalle@gmail.com](mailto:alexander.paul.edelmetalle@gmail.com)  
        Telefonnummer: +49 1573 3729040

        **Umsatzsteuer-Hinweis:**  
        Als Kleinunternehmer im Sinne von § 19 Abs. 1 UStG wird keine Umsatzsteuer berechnet und nicht ausgewiesen.  

        ---

        #### Datenschutzerklärung

        ##### Präambel
        Mit der folgenden Datenschutzerklärung möchten wir Sie darüber aufklären, welche Arten Ihrer personenbezogenen Daten (nachfolgend auch kurz als "Daten" bezeichnet) wir zu welchen Zwecken und in welchem Umfang verarbeiten. Die Datenschutzerklärung gilt für alle von uns durchgeführten Verarbeitungen personenbezogener Daten, sowohl im Rahmen der Erbringung unserer Leistungen als auch insbesondere auf unseren Webseiten, in mobilen Applikationen sowie innerhalb externer Onlinepräsenzen, wie z. B. unserer Social-Media-Profile (nachfolgend zusammenfassend bezeichnet als "Onlineangebot").

        Die verwendeten Begriffe sind nicht geschlechtsspezifisch.

        *Stand: 3. Oktober 2026*

        ---

        ##### Inhaltsübersicht
        * Präambel
        * Verantwortlicher
        * Übersicht der Verarbeitungen
        * Maßgebliche Rechtsgrundlagen
        * Bereitstellung des Onlineangebots und Webhosting
        * Kontakt- und Anfrageverwaltung
        * Rechte der betroffenen Personen

        ---

        ##### Verantwortlicher
        Alexander Paul  
        Vischerstr. 2   
        70563 Stuttgart  
        E-Mail: [alexander.paul.edelmetalle@gmail.com](mailto:alexander.paul.edelmetalle@gmail.com)  
        Telefonnummer: +49 1573 3729040

        ---

        ##### Übersicht der Verarbeitungen
        Die nachfolgende Übersicht fasst die Arten der verarbeiteten Daten und die Zwecke ihrer Verarbeitung zusammen und verweist auf die betroffenen Personen.

        ###### Arten der verarbeiteten Daten
        * Bestandsdaten
        * Kontaktdaten
        * Inhaltsdaten
        * Nutzungsdaten
        * Meta-, Kommunikations- und Verfahrensdaten
        * Protokolldaten

        ###### Kategorien betroffener Personen
        * Leistungsempfänger und Auftraggeber
        * Kommunikationspartner
        * Nutzer
        * Dritte Personen

        ###### Zwecke der Verarbeitung
        * Kommunikation
        * Sicherheitsmaßnahmen
        * Organisations- und Verwaltungsverfahren
        * Feedback
        * Bereitstellung unseres Onlineangebotes und Nutzerfreundlichkeit
        * Informationstechnische Infrastruktur

        ---

        ##### Maßgebliche Rechtsgrundlagen
        **Maßgebliche Rechtsgrundlagen nach der DSGVO:** Im Folgenden erhalten Sie eine Übersicht der Rechtsgrundlagen der DSGVO, auf deren Basis wir personenbezogene Daten verarbeiten. Bitte nehmen Sie zur Kenntnis, dass neben den Regelungen der DSGVO nationale Datenschutzvorgaben in Ihrem bzw. unserem Wohn- oder Sitzland gelten können. Sollten ferner im Einzelfall speziellere Rechtsgrundlagen maßgeblich sein, teilen wir Ihnen diese in der Datenschutzerklärung mit.

        * **Einwilligung (Art. 6 Abs. 1 S. 1 lit. a) DSGVO):** Die betroffene Person hat ihre Einwilligung in die Verarbeitung der sie betreffenden personenbezogenen Daten für einen spezifischen Zweck oder mehrere bestimmte Zwecke gegeben.
        * **Vertragserfüllung und vorvertragliche Anfragen (Art. 6 Abs. 1 S. 1 lit. b) DSGVO):** Die Verarbeitung ist für die Erfüllung eines Vertrags, dessen Vertragspartei die betroffene Person ist, oder zur Durchführung vorvertraglicher Maßnahmen erforderlich, die auf Anfrage der betroffenen Person erfolgen.
        * **Rechtliche Verpflichtung (Art. 6 Abs. 1 S. 1 lit. c) DSGVO):** Die Verarbeitung ist zur Erfüllung einer rechtlichen Verpflichtung erforderlich, der der Verantwortliche unterliegt.
        * **Berechtigte Interessen (Art. 6 Abs. 1 S. 1 lit. f) DSGVO):** Die Verarbeitung ist zur Wahrung der berechtigten Interessen des Verantwortlichen oder eines Dritten notwendig, vorausgesetzt, dass die Interessen, Grundrechte und Grundfreiheiten der betroffenen Person, die den Schutz personenbezogener Daten verlangen, nicht überwiegen.

        **Nationale Datenschutzregelungen in Deutschland:** Zusätzlich zu den Datenschutzregelungen der DSGVO gelten nationale Regelungen zum Datenschutz in Deutschland. Hierzu gehört insbesondere das Bundesdatenschutzgesetz (BDSG). Das BDSG enthält insbesondere Spezialregelungen zum Recht auf Auskunft, zum Recht auf Löschung, zum Widerspruchsrecht, zur Verarbeitung besonderer Kategorien personenbezogener Daten, zur Verarbeitung für andere Zwecke und zur Übermittlung sowie automatisierten Entscheidungsfindung im Einzelfall einschließlich Profiling.

        ---

        ##### Bereitstellung des Onlineangebots und Webhosting
        Wir verarbeiten die Daten der Nutzer, um ihnen unsere Online-Dienste zur Verfügung stellen zu können. Zu diesem Zweck verarbeiten wir die IP-Adresse des Nutzers, die notwendig ist, um die Inhalte und Funktionen unserer Online-Dienste an den Browser oder das Endgerät der Nutzer zu übermitteln.

        * **Verarbeitete Datenarten:** Nutzungsdaten (z. B. Seitenaufrufe und Verweildauer, Klickpfade, Nutzungsintensität und -frequenz, verwendete Gerätetypen und Betriebssysteme, Interaktionen mit Inhalten und Funktionen); Meta-, Kommunikations- und Verfahrensdaten (z. B. IP-Adressen, Zeitangaben, Identifikationsnummern, beteiligte Personen); Protokolldaten (z. B. Logfiles betreffend Logins oder den Abruf von Daten oder Zugriffszeiten).
        * **Betroffene Personen:** Nutzer (z. B. Webseitenbesucher, Nutzer von Onlinediensten).
        * **Zwecke der Verarbeitung und berechtigte Interessen:** Bereitstellung unseres Onlineangebotes und Nutzerfreundlichkeit; Informationstechnische Infrastruktur (Betrieb und Bereitstellung von Informationssystemen und technischen Geräten); Sicherheitsmaßnahmen.
        * **Aufbewahrung und Löschung:** Logfile-Informationen werden für die Dauer von maximal 30 Tagen gespeichert und danach gelöscht oder anonymisiert.
        * **Rechtsgrundlagen:** Berechtigte Interessen (Art. 6 Abs. 1 S. 1 lit. f) DSGVO).

        **Weitere Hinweise zu Verarbeitungsprozessen, Verfahren und Diensten:**
        * **Erhebung von Zugriffsdaten und Logfiles:** Der Zugriff auf unser Onlineangebot wird in Form von sogenannten "Server-Logfiles" protokolliert. Zu den Serverlogfiles können die Adresse und der Name der abgerufenen Webseiten und Dateien, Datum und Uhrzeit des Abrufs, übertragene Datenmengen, Meldung über erfolgreichen Abruf, Browsertyp nebst Version, das Betriebssystem des Nutzers, Referrer URL (die zuvor besuchte Seite) und im Regelfall IP-Adressen und der anfragende Provider gehören. Die Serverlogfiles können zum einen zu Sicherheitszwecken eingesetzt werden, z. B. um eine Überlastung der Server zu vermeiden (insbesondere im Fall von missbräuchlichen Angriffen, sogenannten DDoS-Attacken), und zum anderen, um die Auslastung der Server und ihre Stabilität sicherzustellen; **Rechtsgrundlagen:** Berechtigte Interessen (Art. 6 Abs. 1 S. 1 lit. f) DSGVO). **Löschung von Daten:** Logfile-Informationen werden für die Dauer von maximal 30 Tagen gespeichert und danach gelöscht oder anonymisiert. Daten, deren weitere Aufbewahrung zu Beweiszwecken erforderlich ist, sind bis zur endgültigen Klärung des jeweiligen Vorfalls von der Löschung ausgenommen.

        ---

        ##### Kontakt- und Anfrageverwaltung
        Bei der Kontaktaufnahme mit uns (z. B. per Post, Kontaktformular, E-Mail, Telefon oder via soziale Medien) sowie im Rahmen bestehender Nutzer- und Geschäftsbeziehungen werden die Angaben der anfragenden Personen verarbeitet, soweit dies zur Beantwortung der Kontaktanfragen und etwaiger angefragter Maßnahmen erforderlich ist.

        * **Verarbeitete Datenarten:** Kontaktdaten (z. B. Post- und E-Mail-Adressen oder Telefonnummern); Inhaltsdaten (z. B. textliche oder bildliche Nachrichten und Beiträge sowie die sie betreffenden Informationen, wie z. B. Angaben zur Autorenschaft oder Zeitpunkt der Erstellung); Meta-, Kommunikations- und Verfahrensdaten (z. B. IP-Adressen, Zeitangaben, Identifikationsnummern, beteiligte Personen).
        * **Betroffene Personen:** Kommunikationspartner.
        * **Zwecke der Verarbeitung und berechtigte Interessen:** Kommunikation; Organisations- und Verwaltungsverfahren; Feedback; Bereitstellung unseres Onlineangebotes und Nutzerfreundlichkeit.
        * **Aufbewahrung und Löschung:** Die Daten werden gelöscht, sobald sie für die Erreichung des Zweckes ihrer Erhebung nicht mehr erforderlich sind und keine gesetzlichen Aufbewahrungspflichten entgegenstehen.
        * **Rechtsgrundlagen:** Berechtigte Interessen (Art. 6 Abs. 1 S. 1 lit. f) DSGVO), Vertragserfüllung und vorvertragliche Anfragen (Art. 6 Abs. 1 S. 1 lit. b) DSGVO).

        **Weitere Hinweise zu Verarbeitungsprozessen, Verfahren und Diensten:**
        * **Kontaktformular & E-Mail-Anfragen:** Bei Kontaktaufnahme über unser Kontaktformular, per E-Mail oder anderen Kommunikationswegen verarbeiten wir die uns übermittelten personenbezogenen Daten zur Beantwortung und Bearbeitung des jeweiligen Anliegens. Dies umfasst in der Regel Angaben wie Name, Kontaktinformationen und gegebenenfalls weitere Informationen, die uns mitgeteilt werden und zur angemessenen Bearbeitung erforderlich sind. Wir nutzen diese Daten ausschließlich für den angegebenen Zweck der Kontaktaufnahme und Kommunikation; **Rechtsgrundlagen:** Vertragserfüllung und vorvertragliche Anfragen (Art. 6 Abs. 1 S. 1 lit. b) DSGVO), Berechtigte Interessen (Art. 6 Abs. 1 S. 1 lit. f) DSGVO).

        ---

        ##### Rechte der betroffenen Personen
        Ihnen stehen als Betroffene nach der DSGVO verschiedene Rechte zu, die sich insbesondere aus Art. 15 bis 21 DSGVO ergeben:

        * **Widerspruchsrecht:** Sie haben das Recht, aus Gründen, die sich aus Ihrer besonderen Situation ergeben, jederzeit gegen die Verarbeitung der Sie betreffenden personenbezogenen Daten, die aufgrund von Art. 6 Abs. 1 lit. e oder f DSGVO erfolgt, Widerspruch einzulegen.
        * **Widerrufsrecht bei Einwilligungen:** Sie haben das Recht, erteilte Einwilligungen jederzeit zu widerrufen.
        * **Auskunftsrecht:** Sie haben das Recht, eine Bestätigung darüber zu verlangen, ob betreffende Daten verarbeitet werden und auf Auskunft über diese Daten sowie auf weitere Informationen und Kopie der Daten gemäß den gesetzlichen Vorgaben.
        * **Recht auf Berichtigung:** Sie haben das Recht, die Vervollständigung der Sie betreffenden Daten oder die Berichtigung der Sie betreffenden unrichtigen Daten zu verlangen.
        * **Recht auf Löschung und Einschränkung der Verarbeitung:** Sie haben das Recht, zu verlangen, dass Sie betreffende Daten unverzüglich gelöscht werden, bzw. alternativ eine Einschränkung der Verarbeitung der Daten zu verlangen.
        * **Recht auf Datenübertragbarkeit:** Sie haben das Recht, Sie betreffende Daten, die Sie uns bereitgestellt haben, in einem strukturierten, gängigen und maschinenlesbaren Format zu erhalten oder deren Übermittlung an einen anderen Verantwortlichen zu verlangen.
        * **Beschwerde bei Aufsichtsbehörde:** Sie haben unbeschadet eines anderweitigen verwaltungsrechtlichen oder gerichtlichen Rechtsbehelfs das Recht auf Beschwerde bei einer Aufsichtsbehörde, insbesondere in dem Mitgliedstaat Ihres gewöhnlichen Aufenthaltsorts, Ihres Arbeitsplatzes oder des Orts des mutmaßlichen Verstoßes.

        ---

        [Erstellt mit kostenlosem Datenschutz-Generator.de von Dr. Thomas Schwenke](https://datenschutz-generator.de/)
        """)