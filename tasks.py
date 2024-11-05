from robocorp.tasks import task
import tkinter as tk
from tkinter import ttk, messagebox
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Browser import Browser
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os

load_dotenv()
lahettaja_email = os.getenv("GMAIL_USERNAME")
salasana = os.getenv("GMAIL_PASSWORD")


pdf = FPDF()
http = HTTP()

#globaalit muuttujat
kentta_var = None
root = None
sulje = False

@task
def HaeGolfkenttatiedot():
    """Hakee golfkentän tiedot ja lähettää ne sähköpostitse PDF-tiedostona."""
    luo_kayttoliittyma() 
    kentta = valitse_kentta()
    tallenna_kaikki_tiedot_pdf(kentta)

def sulje_ikkuna():
    """Kysyy käyttäjältä haluaako hän sulkea ikkunan"""
    global sulje
    if messagebox.askokcancel("Sulje", "Haluatko sulkea ohjelman?"):
        sulje = True 
        root.destroy() # sulkee ikkunan

def luo_kayttoliittyma():
    global kentta_var  # globaali muuttuja
    global root # globaali muuttuja
    root = tk.Tk()  # Tkinterin pääikkuna
    root.title("Golfkenttävalinta")  # Otsikko
    root.geometry("600x400")  # Ikkunan koko
    root.configure(bg="#f0f0f0")  # Taustaväri

    # Otsikkokehys
    header_frame = tk.Frame(root, bg="#5cb85c")  # Taustaväri
    header_frame.pack(fill=tk.X)  # Täyttää koko leveyden

    # Otsikko
    otsikko = tk.Label(header_frame, text="Valitse Golfkenttä", font=("Poppins", 32, "bold"), bg="#5cb85c", fg="white", padx=10, pady=10)
    otsikko.pack()  # Sijoittaa otsikon kehykseen

    # Kenttävalinta
    kentta_var = tk.StringVar()  # Muuttuja johon tallennetaan käyttäjän valitsema golfkenttä
    kentta_valinta = ttk.Combobox(root, textvariable=kentta_var, font=("Poppins", 14), state="readonly")  # Pudotusvalikko
    kentta_valinta['values'] = ("Kaikki kentät",  # Kenttävaihtoehdot
                                 "Peuramaa Golf",
                                 "Tuusulan Golfklubi",
                                 "Nevas Golf",
                                 "Keimola Golf",
                                 "Helsingin Golfklubi",
                                 "Espoo Ringside Golf",
                                 "Hyvigolf",
                                 "Golf Talma")  # Kenttävaihtoehdot
    kentta_valinta.set("Valitse golfkenttä")  # Oletusvalinta
    kentta_valinta.pack(pady=(20, 10), padx=20, fill=tk.X)  # padding

    # Näytä tiedot -nappi
    nayta_nappi = tk.Button(root, text="Hae Tiedot",
                            command=lambda: tallenna_kaikki_tiedot_pdf(kentta_var.get()), # Suorittaa tallenna_kaikki_tiedot funktion 
                            bg="#5cb85c", fg="white", font=("Poppins", 16), width=15,
                            relief=tk.RAISED, bd=3, activebackground="#4cae4f")
    nayta_nappi.pack(pady=20)  # Grid-tyylinen sijoittelu

    # Sulje -nappi
    sulje_nappi = tk.Button(root, text="Sulje", command=sulje_ikkuna,  # Suorittaa sulje_ikkuna()-funktion
                            bg="#d9534f", fg="white", font=("Poppins", 16), width=15,
                            relief=tk.RAISED, bd=3, activebackground="#c9302c")
    sulje_nappi.pack(pady=10) # padding 


    # Footer
    footer = tk.Label(root, text="© 2024 Golfkenttävalinta", font=("Poppins", 10), bg="#f0f0f0", fg="#555")
    footer.pack(side=tk.BOTTOM, pady=10)

    root.protocol("WM_DELETE_WINDOW", sulje_ikkuna) # Asettaa sulje_ikkuna funktion ikkunaan
    root.mainloop()  # Käynnistää Tkinterin pääsilmukan




# Golfkentän valinta
def valitse_kentta():
    """Palauttaa käyttäjän valitseman kentän nimen."""
    kentta = kentta_var.get() # Hakee kenttä_var muuttujan arvon
    return kentta # Palauttaa valitun kentän

def tallenna_kaikki_tiedot_pdf(kentta):
    """Hakee kentän säätiedot ja layout-tiedot, ja tallentaa ne samaan PDF-tiedostoon."""
    global sulje
    if sulje:
        return # Lopettaa toiminnon, jos ohjelma suljetaan
    
    if kentta == "Valitse golfkenttä":
        messagebox.showwarning("Virhe", "Valitse kenttä!") # Näyttää varoituksen jos kenttää ei valittu
        return
    
    pdf = FPDF() # Luo uuden PDF
    pdf.set_auto_page_break(auto=True, margin=15) 



# Tarkistaa, valittiinko "Kaikki kentät" ja käytä silloin kaikkia kenttiä
    kentat = [kentta] if kentta != "Kaikki kentät" else [
        "Peuramaa Golf", "Paloheinä Golf", "Tuusulan Golfklubi", 
        "Nevas Golf", "Keimola Golf", "Helsingin Golfklubi", 
        "Espoo Ringside Golf", "Hyvigolf", "Golf Talma"
    ]
    
    for yksittainen_kentta in kentat:

        if sulje:
            break # Lopettaa silmukan, jos ohjelma suljetaan

        # Hakee ja tallentaa säätiedot yksittäiselle kentälle
        saatiedot_kuvat = hae_ja_tallenna_saatiedot(yksittainen_kentta)

        # Lisää säätietojen kuvat PDF:ään
        for kuva in saatiedot_kuvat:
            pdf.add_page()
            pdf.set_font("Arial", "B", size=16)
            pdf.multi_cell(0, 10, f"Säätiedot kentälle: {yksittainen_kentta}")
            pdf.image(kuva, x=10, y=30, w=180)

        # Hakee ja tallentaa layout-tiedot yksittäiselle kentälle
        layout_kuvat = hae_ja_tallenna_layout_tiedot(yksittainen_kentta)

        # Lisää layout-tietojen kuvat PDF:ään
        if layout_kuvat:
            for kuva in layout_kuvat:
                pdf.add_page()
                pdf.set_font("Arial", "B", size=16)
                pdf.multi_cell(0, 10, f"Layout-tiedot kentälle: {yksittainen_kentta}")
                pdf.image(kuva, x=10, y=30, w=180)

        # Ilmoittaa jos layout tietoja ei löydy        
        else:
            pdf.add_page()
            pdf.set_font("Arial", "B", size=16)
            pdf.multi_cell(0, 10, f"Layout-tietoja ei löytynyt kentälle: {yksittainen_kentta}")

    if not sulje:
        # Lopullinen tallennus, tiedoston nimi "Kaikki kentät" -valinnan mukaan
        pdf_output_path = f"C:\\Users\\miika_bz20f79\\Desktop\\golf\\tiedot\\tiedot_{kentta if kentta != 'Kaikki kentät' else 'kaikki_kentat'}.pdf"
        pdf.output(pdf_output_path)  # Tallennetaan PDF tiedostoon 
        send_email(pdf_output_path) # Kutsuu send_email funktion

def hyvaksy_evasteet():
    """Hyväksyy evästäät"""
    page = browser.page()

    evasteet_layout = page.locator("#CybotCookiebotDialogBodyButtonAccept")
    evasteet_saa = page.locator("[aria-label='Salli kaikki evästeet']")
    evasteet_layout2 = page.locator("#wt-cli-accept-all-btn")

    # Hyväksyy evästeet jos näkyvillä
    if evasteet_saa.is_visible():
        evasteet_saa.click()

    elif evasteet_layout.is_visible():
        evasteet_layout.click()

    elif evasteet_layout2.is_visible():
        evasteet_layout2.click()




# Golfkentän valinta
def hae_ja_tallenna_saatiedot(kentta):
    """Hakee säätiedot valitulta kentältä ja palauttaa kuvatiedostot."""
    kentta_urls = {
        "Peuramaa Golf": "https://www.ilmatieteenlaitos.fi/saa/kirkkonummi/peuramaa%20golf",
        "Paloheinä Golf": "https://www.ilmatieteenlaitos.fi/saa/helsinki/palohein%C3%A4",
        "Tuusulan Golfklubi": "https://www.ilmatieteenlaitos.fi/saa/tuusula/tuusula%20golf",
        "Nevas Golf": "https://www.ilmatieteenlaitos.fi/saa/sipoo/nevas%20golf",
        "Keimola Golf": "https://www.ilmatieteenlaitos.fi/saa/vantaa/keimola%20golf",
        "Helsingin Golfklubi": "https://www.ilmatieteenlaitos.fi/saa/helsinki/helsingin%20golfklubi",
        "Espoo Ringside Golf": "https://www.ilmatieteenlaitos.fi/saa/espoo/espoo%20ringside%20golf",
        "Hyvigolf": "https://www.ilmatieteenlaitos.fi/saa/hyvink%C3%A4%C3%A4/hyvink%C3%A4%C3%A4%20golf",
        "Golf Talma": "https://www.ilmatieteenlaitos.fi/saa/sipoo/talma%20golf"
    }

    # Käy läpi kaikki kentät tai käyttää vain valittua kenttää
    kentat = [kentta] if kentta != "Kaikki kentät" else kentta_urls.keys()
    
    kuvat = []  # Lista kuvatiedostoille

    for kentta_nimi in kentat:
        url = kentta_urls.get(kentta_nimi)
        if not url:
            print(f"Säätietoja ei löydy kentälle: {kentta_nimi}")
            continue

        # Siirry URL-osoitteeseen ja hyväksy evästeet
        page = browser.page()
        browser.goto(url)
        hyvaksy_evasteet()
        page.wait_for_load_state("networkidle")

        # Ottaa näyttökuvan kentän säätiedoista
        locator = page.locator("#skip-to-main-content > main > section.row.forecast-container.mx-0") 
        screenshot_path = f"C:\\Users\\miika_bz20f79\\Desktop\\golf\\tiedot\\saatiedot_{kentta_nimi}.png"
        locator.screenshot(path=screenshot_path)
        kuvat.append(screenshot_path)  # Lisää kuvatiedoston listalle

        # Sulkee sivun
        page.close()
    
    return kuvat  # Palauttaa listan kuvatiedostoista


def hae_ja_tallenna_layout_tiedot(kentta):
    """Hakee kentän layout tiedot ja palauttaa kuvatiedostot."""
    layout_urls = {
        "Peuramaa Golf": "https://peuramaagolf.fi/kentat",
        "Paloheinä Golf": "https://paloheinagolf.fi/iso-ysi/",
        "Tuusulan Golfklubi": "https://tgk.fi/kenttaesittely/",
        "Nevas Golf": "https://nevasgolf.fi/fi-fi/kentat/136/",
        "Keimola Golf": "https://keimolagolf.com/fi-fi/saras/vaylaesittely/442/",
        "Hyvigolf": "https://hyvigolf.fi/fi-fi/kentta/kenttaesittely/209/",
        "Golf Talma": "https://www.golftalma.fi/kentta/laakso",
    }

    # Käy läpi kaikki kentät tai käyttää vain valittua kenttää
    kentat = [kentta] if kentta != "Kaikki kentät" else layout_urls.keys()
    
    kuvat = []  # Lista kuvatiedostoille

    for kentta_nimi in kentat:
        url = layout_urls.get(kentta_nimi)
        if not url:
            print(f"Layout-elementtia ei loydy kentalle: {kentta_nimi}")
            continue

        # Siirtyy URL-osoitteeseen ja hyväksyy evästeet
        page = browser.page()
        browser.goto(url)
        hyvaksy_evasteet()
        page.wait_for_load_state("networkidle")

        # Määrittää layout-locatorin kentän mukaan
        if kentta_nimi == "Peuramaa Golf":
            layout_locator = page.locator("#article > div.col-xs-12.element-cdn-image > a > picture > img")
            layout_locator.wait_for(state="visible", timeout=10000)
        elif kentta_nimi == "Paloheinä Golf":
            layout_locator = page.locator("#post-394 > figure > img")
            layout_locator.wait_for(state="visible", timeout=10000)
        elif kentta_nimi == "Tuusulan Golfklubi":
            layout_locator = page.locator("#article > div.col-xs-12.element-cdn-image > a > picture > img")
            layout_locator.wait_for(state="visible", timeout=10000)
        elif kentta_nimi == "Nevas Golf":
            layout_locator = page.locator("#article > div:nth-child(2) > a > img")
            layout_locator.wait_for(state="visible", timeout=10000)
        elif kentta_nimi == "Keimola Golf":
            layout_locator = page.locator("#article > div.col-xs-12.element-cdn-image > a > picture > img")
            layout_locator.wait_for(state="visible", timeout=10000)
        elif kentta_nimi == "Hyvigolf":
            layout_locator = page.locator("#article > div > div.col-xs-12.element-cdn-image.col-sm-6 > a > picture > img")
            layout_locator.wait_for(state="visible", timeout=10000)
        elif kentta_nimi == "Golf Talma":
            layout_locator = page.locator("body > div.root > div > div.module.single-course > div > div.content.text-styles > p > img")
            layout_locator.wait_for(state="visible", timeout=10000)


        # Otta näyttökuvan
        kuvatiedosto = f"C:\\Users\\miika_bz20f79\\Desktop\\golf\\tiedot\\layout_{kentta_nimi}.png"
        layout_locator.screenshot(path=kuvatiedosto)
        kuvat.append(kuvatiedosto)  # Lisää kuvatiedoston listalle

        # Sulkee sivun
        page.close()
    
    return kuvat  # Palautaa listan kuvatiedostoista


def send_email(pdf_file_path):
    """Lähettää PDF tiedoston sähköpostitse"""

    vastaanottaja_email = "miika.hanninen@student.laurea.fi"

    # Valmistelee sähköpostin
    viesti = MIMEMultipart()
    viesti['FROM'] = lahettaja_email
    viesti['TO'] = vastaanottaja_email
    viesti['Subject'] = "Golfkenttätiedot PDF"

    # Liitetään PDF tiedosto sähköpostiin
    with open(pdf_file_path, "rb") as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_file_path)}')
        viesti.attach(part)

        # Luo yhteyden ja lähettää sähköpostin
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server: # Gmailin SMTP-palvelin
            server.starttls() # Salaus
            server.login(lahettaja_email, salasana) # Kirjautuu sisään
            server.send_message(viesti) # Lähettää sähköpostin
            print("Sähköposti lähetetty")
    except Exception as e:
        print(f"Sähköpostin lähettämisessä tapahtui virhe: {e}") # Virheilmoitus