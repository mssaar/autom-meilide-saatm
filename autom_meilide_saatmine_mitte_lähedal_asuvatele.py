import os
import smtplib
import time
from email.message import EmailMessage
from dotenv import load_dotenv
import pandas as pd

# Määra täpne keskkonnafaili nimi soovi korral, hetkel on fail .env
load_dotenv()

# --- SEADISTUSED ---
EXCEL_FAIL = "firmad.xlsx"
SAADETUD_LOG = "saadetud_meilid.txt"
PILDI_FAIL = "suur_saal.jpg"  # Vähendatud mahuga pilt samas kaustas

# ZeptoMaili SMTP seaded
PORT = 587
SMTP_SERVER = "smtp.zeptomail.eu"
USERNAME = "emailapikey"
PASSWORD = os.getenv("ZEPTO_PASSWORD")

# Saatja ja lingid
FROM_EMAIL = "info@therand.ee"
BOOKING_URL = "https://therand.ee"
UNSUBSCRIBE_URL = "mailto:info@therand.ee?subject=Loobu%20uudiskirjast&body=Palun%20eemaldada%20meie%20e-post%20listist."


def lae_saadetud_meilid():
    if os.path.exists(SAADETUD_LOG):
        with open(SAADETUD_LOG, "r", encoding="utf-8") as f:
            return set(line.strip().lower() for line in f if line.strip())
    return set()


def salvesta_saadetud_meil(meil):
    with open(SAADETUD_LOG, "a", encoding="utf-8") as f:
        f.write(f"{meil.strip().lower()}\n")


def saada_turundusmeilid():
    if not PASSWORD:
        print("VIGA: ZEPTO_PASSWORD puudub prod-credentials.env failist!")
        return

    if not os.path.exists(EXCEL_FAIL):
        print(f"VIGA: Faili '{EXCEL_FAIL}' ei leitud!")
        return

    if not os.path.exists(PILDI_FAIL):
        print(
            f"VIGA: Pildifaili '{PILDI_FAIL}' ei leitud samast kaustast skriptiga!"
        )
        return

    with open(PILDI_FAIL, "rb") as f:
        pildi_andmed = f.read()

    df = pd.read_excel(EXCEL_FAIL, header=None)
    df = df.dropna(subset=[0])
    saadetud_aadressid = lae_saadetud_meilid()

    print(
        f"Excelis leitud {len(df)} rida. Varem saadetud: {len(saadetud_aadressid)}."
    )
    print("Alustan saatmist...")

    saadetud_loendur = 0

    for index, row in df.iterrows():
        saaja_meil = str(row[0]).strip().lower()

        if "@" not in saaja_meil or saaja_meil in saadetud_aadressid:
            continue

        raw_firma = row[1] if len(row) > 1 else None
        firma_nimi = (
            str(raw_firma).strip()
            if pd.notna(raw_firma) and str(raw_firma).strip()
            else "meeskond"
        )

        msg = EmailMessage()
        msg["Subject"] = "Tiimiüritus The Rand Rannahallis! 🏖️"
        msg["From"] = f"The Rand Rannahall <{FROM_EMAIL}>"
        msg["To"] = saaja_meil
        msg["List-Unsubscribe"] = f"<{UNSUBSCRIBE_URL}>"

        msg.set_content(
            f"Tere {firma_nimi}!\n\n"
            f"🌴 Suvine meeleolu aastaringselt! Tiimiüritus liival!\n\n"
            f"Otsite võimalust värskendada tiimivaimu? The Rand Rannahallis saate nautida rannaspordialasid aastaringselt!\n\n"
            f"Tutvu võimalustega ja broneeri: {BOOKING_URL}\n\n"
            f"Loobumiseks vasta sellele kirjale märkega 'Loobu'."
        )

        # HTML sisu koos rohelise kasti ja pealkirjaga
        html_sisu = f"""
        <!DOCTYPE html>
        <html>
        <head><meta charset="utf-8"></head>
        <body style="margin: 0; padding: 20px 0; background-color: #FBF8F3; font-family: sans-serif;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td align="center">
                        <table role="presentation" width="100%" style="max-width: 600px; background-color: #FFFFFF; border-radius: 12px; border: 1px solid #F5EDDD; overflow: hidden;">
                            <tr>
                                <td style="padding: 32px; text-align: center; background-color: #FAF5EB;">
                                    <!-- Õrn roheline kast ja palm -->
                                    <div style="display: inline-block; background-color: #E8F5E9; border: 1px solid #C8E6C9; color: #2E7D32; font-size: 13px; font-weight: bold; padding: 6px 16px; border-radius: 20px; margin-bottom: 12px;">
                                        🌴 Suvine meeleolu aastaringselt
                                    </div>
                                    <h1 style="margin: 0; font-size: 24px; color: #1A1614;">Tiimiüritus liival!</h1>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding: 32px; font-size: 15px; color: #1A1614; line-height: 1.6;">
                                    <p>Tere, <strong>{firma_nimi}</strong>!</p>
                                    <p>Otsite võimalust värskendada tiimivaimu või tähistada ettevõtte olulist sündmust? 6 väljakuga <strong>The Rand Rannahall</strong> ootab teid külla!</p>
                                    <p>Tulge mängige rannavõrkpalli, rannatennist või korraldage meeleolukas turniir.</p>
                                    
                                    <div style="background-color: #FFFDF9; border-left: 4px solid #C85F0E; padding: 12px; margin: 20px 0;">
                                        🔥 <strong>Saun hinnas!</strong>
                                    </div>

                                    <div style="text-align: center; margin: 25px 0;">
                                        <a href="{BOOKING_URL}" target="_blank" style="background-color: #C85F0E; color: #FFFFFF; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-weight: bold; display: inline-block;">
                                            Tutvu võimalustega ja broneeri &rarr;
                                        </a>
                                    </div>
                                    
                                    <div style="text-align: center; margin-top: 20px;">
                                        <img src="cid:pilt_suur_saal" alt="The Rand rannahall" style="width: 100%; max-width: 500px; border-radius: 8px;">
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding: 15px; background-color: #FAF5EB; text-align: center; font-size: 12px; color: #6B615A;">
                                    The Rand Rannahall | info@therand.ee | +372 5806 5253<br><br>
                                    <a href="{UNSUBSCRIBE_URL}" style="color: #999999; text-decoration: underline;">Loobu teavitustest / Unsubscribe</a>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </body>
        </html>
        """
        msg.add_alternative(html_sisu, subtype="html")

        # Seome pildi otse HTML-i sisse CID tunnusega
        msg.get_payload(1).add_related(
            pildi_andmed,
            maintype="image",
            subtype="jpeg",
            cid="<pilt_suur_saal>",
        )

        try:
            with smtplib.SMTP(SMTP_SERVER, PORT) as server:
                server.starttls()
                server.login(USERNAME, PASSWORD)
                server.send_message(msg)

            saadetud_loendur += 1
            print(f"{saadetud_loendur}. Saadetud: {firma_nimi} ({saaja_meil})")
            salvesta_saadetud_meil(saaja_meil)
            saadetud_aadressid.add(saaja_meil)

        except Exception as e:
            print(f"Viga saatmisel aadressile {saaja_meil}: {e}")

        time.sleep(2)

    print(f"\nValmis! Saadeti kokku {saadetud_loendur} kirja.")


if __name__ == "__main__":
    saada_turundusmeilid()