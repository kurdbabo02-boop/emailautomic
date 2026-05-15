import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import smtplib
import time
import uuid
import base64
import html
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formatdate, make_msgid
from email.utils import formataddr


def build_email_bodies(
    recipient_name,
    message_body,
    message_body_after_logo="",
    logo_src=None,
    logo_width=160,
    logo_position="Tussen tekstblokken",
    logo_align="Links",
):
    text_parts = [part.strip() for part in [message_body, message_body_after_logo] if part.strip()]
    plain_text = f"Geachte {recipient_name},\n\n" + "\n\n".join(text_parts)
    safe_name = html.escape(str(recipient_name))
    safe_body = html.escape(message_body.strip()).replace("\n", "<br>\n")
    safe_body_after_logo = html.escape(message_body_after_logo.strip()).replace("\n", "<br>\n")
    normalized_position = logo_position.lower()
    align_map = {"Links": "left", "Midden": "center", "Rechts": "right"}
    logo_td_align = align_map.get(logo_align, "left")
    top_logo_block = ""
    greeting_logo_block = ""
    between_text_logo_block = ""
    bottom_logo_block = ""

    if logo_src and normalized_position != "niet tonen":
        logo_block = f"""
            <tr>
                <td align="{logo_td_align}" style="padding:0 0 20px 0;">
                    <img src="{logo_src}" width="{int(logo_width)}" alt="Logo" style="display:inline-block;max-width:100%;height:auto;border:0;">
                </td>
            </tr>
        """
        if normalized_position == "onder aanhef":
            greeting_logo_block = logo_block
        elif normalized_position == "tussen tekstblokken":
            between_text_logo_block = logo_block.replace("padding:0 0 20px 0;", "padding:20px 0;")
        elif normalized_position == "onder bericht":
            bottom_logo_block = logo_block.replace("padding:0 0 20px 0;", "padding:20px 0 0 0;")
        else:
            top_logo_block = logo_block

    body_before_logo_block = ""
    if safe_body:
        body_before_logo_block = f"""
                    <tr>
                        <td style="color:#202124;">
                            <p style="margin:0;color:#202124;">{safe_body}</p>
                        </td>
                    </tr>
        """

    body_after_logo_block = ""
    if safe_body_after_logo:
        body_after_logo_block = f"""
                    <tr>
                        <td style="color:#202124;">
                            <p style="margin:0;color:#202124;">{safe_body_after_logo}</p>
                        </td>
                    </tr>
        """

    html_body = f"""<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title></title>
</head>
<body style="margin:0;padding:0;background-color:#ffffff;color:#202124;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#ffffff;color:#202124;">
        <tr>
            <td align="left" style="padding:24px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;background-color:#ffffff;color:#202124;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.6;">
                    {top_logo_block}
                    <tr>
                        <td style="color:#202124;">
                            <p style="margin:0 0 16px 0;color:#202124;">Geachte {safe_name},</p>
                        </td>
                    </tr>
                    {greeting_logo_block}
                    {body_before_logo_block}
                    {between_text_logo_block}
                    {body_after_logo_block}
                    {bottom_logo_block}
                </table>
            </td>
        </tr>
    </table>
</body>
</html>"""

    return plain_text, html_body


def logo_data_uri(uploaded_logo):
    if not uploaded_logo:
        return None
    encoded = base64.b64encode(uploaded_logo.getvalue()).decode("ascii")
    mime_type = uploaded_logo.type or "image/png"
    return f"data:{mime_type};base64,{encoded}"

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="Mail Verzender",
    page_icon="✉️",
    layout="wide",
)

# ── Custom CSS ───────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #1a73e8;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        color: #5f6368;
        margin-bottom: 1.5rem;
        font-size: 0.95rem;
    }
    .stButton > button[kind="primary"] {
        background-color: #1a73e8;
        font-weight: 600;
        padding: 0.6rem 1.2rem;
    }
    .preview-box {
        background-color: #ffffff;
        color: #202124;
        border-left: 4px solid #1a73e8;
        border: 1px solid #dadce0;
        padding: 1rem 1.2rem;
        border-radius: 4px;
        font-size: 0.88rem;
    }
    .preview-box p {
        color: #202124;
    }
    .count-badge {
        background-color: #e8f0fe;
        color: #1a73e8;
        font-weight: 600;
        padding: 0.15rem 0.6rem;
        border-radius: 12px;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# ── Session state ────────────────────────────────────────────
if "recipients" not in st.session_state:
    st.session_state.recipients = []
if "send_results" not in st.session_state:
    st.session_state.send_results = []

# ── Header ───────────────────────────────────────────────────
st.markdown('<div class="main-header">✉️ Geautomatiseerde Mail Verzender</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Voeg ontvangers toe, stel je bericht in en verstuur in één klik.</div>', unsafe_allow_html=True)

# ── Sidebar: SMTP instellingen ───────────────────────────────
with st.sidebar:
    st.header("⚙️ E-mailinstellingen")

    sender_name = st.text_input("Jouw naam (afzender)", placeholder="Jan de Vries")
    sender_email = st.text_input("Jouw e-mailadres", placeholder="jij@jouwdomein.nl")
    sender_password = st.text_input(
        "Wachtwoord",
        type="password",
        help="Het wachtwoord dat je voor dit e-mailadres hebt ingesteld bij TransIP.",
    )

    with st.expander("🔧 Geavanceerde SMTP-instellingen", expanded=False):
        smtp_server = st.text_input("SMTP server", value="smtp.transip.email")
        smtp_port = st.number_input("SMTP poort", value=465, min_value=1, max_value=65535, step=1)
        use_ssl = st.toggle("SSL gebruiken (poort 465)", value=True)

    st.markdown("---")
    st.markdown("""
**📌 TransIP SMTP-instellingen:**
- Server: `smtp.transip.email`
- Poort: `465` (SSL)
- Gebruikersnaam: je volledige e-mailadres
- Wachtwoord: het wachtwoord van je mailbox

Wachtwoord vergeten of wijzigen? Ga naar je [TransIP controlepaneel](https://www.transip.nl/cp/) → E-mail.
""")

    with st.expander("📬 Minder vaak in spam belanden", expanded=False):
        st.markdown("""
Spamfilters kijken vooral naar je domeininstellingen en reputatie. Controleer in je DNS:
- **SPF** staat goed ingesteld voor TransIP.
- **DKIM** is actief voor je domein/mailbox.
- **DMARC** bestaat en past bij je verzenddomein.
- Verstuur vanaf hetzelfde domein als je SMTP-account, dus geen afwijkend `From`-adres.

De app verstuurt nu nette plain-text én HTML, gebruikt stabiele headers en voegt geen misleidende mailclient-header toe.
""")

# ── Main layout ──────────────────────────────────────────────
left, right = st.columns([1, 1], gap="large")

# ════════════════════════════════════════════════════════════
# LINKS: Ontvangers
# ════════════════════════════════════════════════════════════
with left:
    st.subheader("👥 Ontvangers")

    tab_manual, tab_excel = st.tabs(["✏️ Handmatig invoeren", "📊 Excel uploaden"])

    # ── Tab 1: Handmatig ─────────────────────────────────────
    with tab_manual:
        with st.form("manual_entry", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                m_name = st.text_input("Naam", placeholder="Sara Janssen")
            with c2:
                m_email = st.text_input("E-mailadres", placeholder="sara@voorbeeld.nl")
            submitted = st.form_submit_button("➕ Toevoegen", use_container_width=True)

        if submitted:
            if m_name.strip() and m_email.strip():
                if "@" not in m_email:
                    st.error("Voer een geldig e-mailadres in.")
                else:
                    st.session_state.recipients.append(
                        {"Naam": m_name.strip(), "E-mail": m_email.strip()}
                    )
                    st.success(f"✅ **{m_name}** toegevoegd!")
            else:
                st.error("Vul zowel naam als e-mailadres in.")

    # ── Tab 2: Excel ─────────────────────────────────────────
    with tab_excel:
        st.caption("Upload een .xlsx bestand met minimaal een kolom voor naam en een voor e-mailadres.")
        uploaded = st.file_uploader("Kies een Excel-bestand", type=["xlsx", "xls"], label_visibility="collapsed")

        if uploaded:
            try:
                df_upload = pd.read_excel(uploaded)
                st.markdown("**Voorbeeld van je bestand:**")
                st.dataframe(df_upload.head(5), use_container_width=True, hide_index=True)

                cols = df_upload.columns.tolist()
                c1, c2 = st.columns(2)
                with c1:
                    naam_col = st.selectbox("Naam-kolom", cols, key="naam_col")
                with c2:
                    email_col = st.selectbox("E-mail-kolom", cols, key="email_col")

                if st.button("📥 Importeer ontvangers", use_container_width=True, type="primary"):
                    new_rows = [
                        {"Naam": str(row[naam_col]).strip(), "E-mail": str(row[email_col]).strip()}
                        for _, row in df_upload.iterrows()
                        if pd.notna(row[naam_col]) and pd.notna(row[email_col])
                        and str(row[email_col]).strip() != ""
                    ]
                    st.session_state.recipients.extend(new_rows)
                    st.success(f"✅ **{len(new_rows)} ontvangers** geïmporteerd!")
            except Exception as e:
                st.error(f"Kon bestand niet lezen: {e}")

    # ── Ontvangerslijst ──────────────────────────────────────
    st.markdown("---")
    count = len(st.session_state.recipients)
    st.markdown(
        f"**📋 Ontvangerslijst** &nbsp;<span class='count-badge'>{count} persoon{'en' if count != 1 else ''}</span>",
        unsafe_allow_html=True,
    )

    if st.session_state.recipients:
        df_r = pd.DataFrame(st.session_state.recipients)
        st.dataframe(df_r, use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🗑️ Lijst wissen", use_container_width=True):
                st.session_state.recipients = []
                st.session_state.send_results = []
                st.rerun()
        with c2:
            csv = df_r.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Download lijst",
                data=csv,
                file_name="ontvangers.csv",
                mime="text/csv",
                use_container_width=True,
            )
    else:
        st.info("Nog geen ontvangers toegevoegd.")

# ════════════════════════════════════════════════════════════
# RECHTS: Bericht
# ════════════════════════════════════════════════════════════
with right:
    st.subheader("📝 Bericht instellen")

    subject = st.text_input("📌 Onderwerp", placeholder="Bijv. Uitnodiging voor ons evenement")

    st.markdown("**Berichttekst**")
    st.caption(
        "Schrijf de tekst die boven het logo moet staan. De aanhef **Geachte [naam],** "
        "wordt automatisch toegevoegd op basis van de naam van de ontvanger."
    )

    message_body = st.text_area(
        "Tekst boven logo",
        height=170,
        placeholder=(
            "bedankt voor je interesse in ons aanbod.\n\n"
            "We nodigen je graag uit voor..."
        ),
    )

    st.markdown("**Logo toevoegen**")
    logo_file = st.file_uploader(
        "Upload je logo",
        type=["png", "jpg", "jpeg", "gif"],
        help="Het logo wordt als inline afbeelding meegestuurd, niet als losse link.",
        label_visibility="collapsed",
    )
    logo_width = st.slider("Logo breedte", min_value=80, max_value=320, value=160, step=10)
    c_logo_position, c_logo_align = st.columns(2)
    with c_logo_position:
        logo_position = st.selectbox(
            "Logo positie",
            ["Tussen tekstblokken", "Bovenaan", "Onder aanhef", "Onder bericht", "Niet tonen"],
        )
    with c_logo_align:
        logo_align = st.selectbox("Logo uitlijning", ["Links", "Midden", "Rechts"])

    message_body_after_logo = st.text_area(
        "Tekst onder logo",
        height=150,
        placeholder=(
            "Met vriendelijke groeten,\n"
            "Jouw naam"
        ),
    )

    # ── Live preview ─────────────────────────────────────────
    st.markdown("**👁️ Voorbeeld e-mail**")
    if message_body.strip() or message_body_after_logo.strip():
        preview_name = (
            st.session_state.recipients[0]["Naam"]
            if st.session_state.recipients
            else "Sara Janssen"
        )
        _, preview_html = build_email_bodies(
            preview_name,
            message_body,
            message_body_after_logo=message_body_after_logo,
            logo_src=logo_data_uri(logo_file),
            logo_width=logo_width,
            logo_position=logo_position,
            logo_align=logo_align,
        )
        components.html(preview_html, height=340, scrolling=True)
    else:
        st.markdown(
            '<div class="preview-box" style="color:#9aa0a6;">Typ je bericht om een voorbeeld te zien...</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Verstuur ─────────────────────────────────────────────
    send_clicked = st.button(
        f"🚀 Verstuur {count} e-mail{'s' if count != 1 else ''}",
        type="primary",
        use_container_width=True,
        disabled=(count == 0),
    )

    if send_clicked:
        errors = []
        if not sender_email.strip():
            errors.append("Vul je e-mailadres in via de zijbalk.")
        if not sender_password.strip():
            errors.append("Vul je app-wachtwoord in via de zijbalk.")
        if not subject.strip():
            errors.append("Onderwerp mag niet leeg zijn.")
        if not message_body.strip() and not message_body_after_logo.strip():
            errors.append("Berichttekst mag niet leeg zijn.")

        if errors:
            for err in errors:
                st.error(f"❌ {err}")
        else:
            results = []
            progress = st.progress(0, text="Bezig met versturen...")
            status = st.empty()
            total = len(st.session_state.recipients)

            try:
                if use_ssl:
                    server = smtplib.SMTP_SSL(smtp_server, int(smtp_port))
                else:
                    server = smtplib.SMTP(smtp_server, int(smtp_port))
                    server.ehlo()
                    server.starttls()
                server.login(sender_email.strip(), sender_password.strip())

                for i, rec in enumerate(st.session_state.recipients):
                    name = rec["Naam"]
                    email = rec["E-mail"]
                    try:
                        msg = MIMEMultipart("related")
                        alternative = MIMEMultipart("alternative")
                        msg.attach(alternative)

                        sender_email_clean = sender_email.strip()
                        display_from = (
                            formataddr((sender_name.strip(), sender_email_clean))
                            if sender_name.strip()
                            else sender_email_clean
                        )
                        sender_domain = sender_email_clean.split("@")[-1]
                        msg["From"] = display_from
                        msg["To"] = email
                        msg["Reply-To"] = display_from
                        msg["Subject"] = subject.strip()
                        msg["Date"] = formatdate(localtime=True)
                        msg["Message-ID"] = make_msgid(domain=sender_domain)
                        msg["MIME-Version"] = "1.0"
                        msg["Content-Language"] = "nl"
                        msg["Auto-Submitted"] = "no"

                        logo_cid = f"logo-{uuid.uuid4().hex}" if logo_file else None
                        full_text, full_html = build_email_bodies(
                            name,
                            message_body,
                            message_body_after_logo=message_body_after_logo,
                            logo_src=f"cid:{logo_cid}" if logo_cid else None,
                            logo_width=logo_width,
                            logo_position=logo_position,
                            logo_align=logo_align,
                        )

                        alternative.attach(MIMEText(full_text, "plain", "utf-8"))
                        alternative.attach(MIMEText(full_html, "html", "utf-8"))

                        if logo_file and logo_cid:
                            logo_subtype = (logo_file.type or "image/png").split("/")[-1].lower()
                            if logo_subtype == "jpg":
                                logo_subtype = "jpeg"
                            logo_image = MIMEImage(logo_file.getvalue(), _subtype=logo_subtype)
                            logo_image.add_header("Content-ID", f"<{logo_cid}>")
                            logo_image.add_header("Content-Disposition", "inline", filename=logo_file.name)
                            msg.attach(logo_image)

                        server.sendmail(sender_email_clean, email, msg.as_string())
                        results.append({"Naam": name, "E-mail": email, "Status": "✅ Verzonden"})
                        status.markdown(f"📤 Verzonden naar **{name}** ({i+1}/{total})")
                    except Exception as e:
                        results.append({"Naam": name, "E-mail": email, "Status": f"❌ {str(e)}"})

                    progress.progress((i + 1) / total, text=f"Bezig... {i+1}/{total}")
                    time.sleep(1.0)

                server.quit()

            except smtplib.SMTPAuthenticationError:
                st.error("❌ Inloggen mislukt. Controleer je e-mailadres en app-wachtwoord.")
                results = []
            except Exception as e:
                st.error(f"❌ Verbindingsfout: {e}")
                results = []

            if results:
                status.empty()
                progress.empty()
                st.session_state.send_results = results

                success_n = sum(1 for r in results if "✅" in r["Status"])
                fail_n = total - success_n

                if success_n == total:
                    st.success(f"🎉 Alle **{total}** e-mails succesvol verzonden!")
                else:
                    st.warning(f"✅ {success_n} verzonden - ❌ {fail_n} mislukt")

                st.markdown("**📊 Verzendresultaten**")
                st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)

# ── Previous results (persistent across reruns) ──────────────
if st.session_state.send_results and not send_clicked:
    with st.expander("📊 Laatste verzendresultaten bekijken", expanded=False):
        st.dataframe(
            pd.DataFrame(st.session_state.send_results),
            use_container_width=True,
            hide_index=True,
        )
