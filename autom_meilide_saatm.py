import pandas as pd
import smtplib
import time
import os
from email.message import EmailMessage
from email.utils import make_msgid

# --- SEADISTUSED ---
EXCEL_FAIL = 'firmad.xlsx'
MINU_MEIL = 'info@therand.ee'
PAROOL = 'J0nnipunn1976.'
SMTP_SERVER = 'mail.therand.ee'
SMTP_PORT = 465
PILDI_FAIL = 'suur_saal.jpg'

def saada_meilid():
    try:
        # Kontrollime, kas pilt on üldse olemas
        if not os.path.exists(PILDI_FAIL):
            print(f"VIGA: Pilti '{PILDI_FAIL}' ei leitud samast kaustast!")
            return

        df = pd.read_excel(EXCEL_FAIL, header=None)
        df = df.dropna(subset=[0])

        print(f"Leitud {len(df)} rida. Alustan saatmist...")

        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(MINU_MEIL, PAROOL)

        for index, row in df.iterrows():
            saaja_meil = str(row[0]).strip()
            
            raw_firma = row[1]
            if pd.isna(raw_firma) or str(raw_firma).strip() == "":
                firma_nimi = "meeskond"
            else:
                firma_nimi = str(raw_firma).strip()

            if "@" not in saaja_meil:
                print(f"Jätan vahele vigase aadressi: {saaja_meil}")
                continue

            # Koostame kirja
            msg = EmailMessage()
            msg['Subject'] = "Teie naabrusesse kolis suvi! ☀️"
            msg['From'] = f"The Rand Rannahall <{MINU_MEIL}>"
            msg['To'] = saaja_meil

            # Loome unikaalse ID pildi jaoks, et seda HTML-is kasutada
            image_cid = make_msgid()

            # Kirja sisu HTML-vormingus (et pilti näidata)
            html_sisu = f"""
            <html>
                <body>
                    <p>Hea {firma_nimi},</p>
                    <p>Mis oleks, kui järgmine meeskonnaüritus toimuks varbad liivas?</p>
                    <p>Oleme teie vahetus läheduses (https://maps.app.goo.gl/Q3eYYykHVpcFUE3w6) avanud uue rannahalli, mis on loodud liikumiseks ja ühisteks hetkedeks. Rannasport on suurepärane viis meeskonnavaimu turgutamiseks ja tervise edendamiseks – ja seda sõltumata välistemperatuurist.</p>
                    <p>Pakume teile mugavat asukohta, kvaliteetset liiva ja parimat atmosfääri.</p>
                    <p>Rohkem infot: <a href="https://therand.ee">therand.ee</a></p>
                    <p>Tervitades,<br>The Rand tiim</p>
                    <img src="cid:{image_cid[1:-1]}" alt="The Rand saal" style="max-width: 600px; height: auto;">
                </body>
            </html>
            """
            
            msg.set_content("Oleme teie vahetus läheduses avanud uue rannahalli") # Varutekst
            msg.add_alternative(html_sisu, subtype='html')

            # Lisame pildi faili manusesse ja seostame CID-ga
            with open(PILDI_FAIL, 'rb') as img:
                msg.get_payload()[1].add_related(
                    img.read(), 
                    maintype='image', 
                    subtype='jpeg', 
                    cid=image_cid
                )

            try:
                server.send_message(msg)
                print(f"{index + 1}. Saadetud: {firma_nimi} ({saaja_meil})")
            except Exception as e:
                print(f"Viga saatmisel aadressile {saaja_meil}: {e}")
            
            time.sleep(2)

        server.quit()
        print("\nKõik töödeldud!")

    except Exception as e:
        print(f"Kriitiline viga: {e}")

if __name__ == "__main__":
    saada_meilid()