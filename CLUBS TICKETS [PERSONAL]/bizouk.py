import datetime
import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

myDatetime = datetime.datetime.now()
today = datetime.date.today()
today_format = today.strftime('%Y%m%d')
# print(today_format)

un_vendredi = datetime.datetime(2022, 6, 24)
un_vendredi_format = un_vendredi.strftime('%A')
# print(un_vendredi_format)


un_samedi = datetime.datetime(2022, 6, 25)
un_samedi_format = un_samedi.strftime('%A')
# print(un_samedi_format)

#CREDENTALS WEBSITE
print("Veuillez entrer votre email de connexion ")
mail_user= input()
print("Veuillez entrer votre password de connexion ")
password_user= input()

if "Friday" != un_vendredi_format:
    fin = "It's not friday !!! STOP ALL"
    print(fin)
    sys.exit()


if "Saturday" != un_samedi_format:
    fin = "It's not saturday !!! STOP ALL"
    print(fin)
    sys.exit()

td = myDatetime - un_samedi
sortie = td.days % 7
# print(sortie)

"""
Séparation en 2 partie : vendredi ou samedi 
"""

# Vendredi
date_party = myDatetime
compteur_days = 1
while sortie != 0:
    compteur_days = compteur_days + 1
    one_day = datetime.timedelta(days=1)
    date_party = date_party + one_day
    td = date_party - un_vendredi
    sortie = td.days % 7

# print(str(compteur_days) + ' days')

day = datetime.timedelta(days=compteur_days, weeks=0)
date_party = myDatetime + day

print(date_party.strftime('DATE SEARCHED = %Y%m%d : %A'))
date = date_party.strftime('%Y%m%d')
# print(date)

# create variable  for lunch driver
driver = webdriver.Chrome()

# CACHER CREDENTIALS

driver.get("https://www.bizouk.com/login/log-in")

email = driver.find_element(By.ID, "email")
email.send_keys(email_user)
email.submit()

email = driver.find_element(By.ID, "password")
email.send_keys(password_user)
email.submit()

#  CHOIX DE SOIREE

    # L'EMPIRE CLUB
    # 911 PARIS
    #

#

driver.get('https://www.bizouk.com/soirees/agenda/region/paris/' + date)
lieu_soiree = driver.find_elements(By.CSS_SELECTOR, "p.soiree_lieu")
lieu_choisi = "L'EMPIRE CLUB"

for element in lieu_soiree:
    if element.text == lieu_choisi:
        place_party = element.text

        print(element.text)
        element.click()

        # Quantite de tickets

        theme_party = driver.find_element(By.CLASS_NAME, 'party_entete')
        select = driver.find_elements(By.CLASS_NAME,)
        print('Vous souhaitez prendre combient de ticket pour  ' + theme_party.text + ' au ' + place_party + ' ?')
        quantity_tikets = '1'

        gratuite = driver.find_element(By.XPATH,
                                       "/html/body/main/div[1]/div[2]/div[1]/div/div[3]/div/form/div/table[1]/tbody/tr[1]")
        price = gratuite.text
        # Vérifier si la sous-chaine se trouve dans la chaine principale
        if "/n0.00€" not in price:
            # print('Sous-chaîne non trouvée')
            print("\n Il n'y a plus de places gratuites!\n")
            sys.exit()
        else:
            # print('Sous-chaîne trouvée')
            nbr_place = Select(driver.find_element(By.XPATH,
                                                   "/html/body/main/div[1]/div[2]/div[1]/div/div[3]/div/form/div/table[1]/tbody/tr[1]/td[1]/select"))
            nbr_place.select_by_value(quantity_tikets)
            confirm = driver.find_element(By.ID, "btn-resa-etape-contenu")
            confirm.submit()

            # Remplir informations de ticket
            nbr_person = driver.find_elements(By.CLASS_NAME, 'panel.panel-primary.bzk-panel-primary')
            for person in nbr_person:
                nom_du_porteur = driver.find_element(By.ID, "last_name_6002214")
                nom_du_porteur.send_keys("NAME")

                prenom_du_porteur = driver.find_element(By.ID, "first_name_6002214")
                prenom_du_porteur.send_keys("SURNAME")

                nom_prenom = driver.find_element(By.ID, "answers_13666_1")
                nom_prenom.send_keys(mail_user)

                email_ticket = driver.find_element(By.ID, "answers_13667_1")
                email_ticket.send_keys(mail_user)

                telephone = driver.find_element(By.ID, "answers_13668_1")
                telephone.send_keys("0702050607")
"""
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[1])
driver.get('https://www.bizouk.com/soirees/agenda/region/paris/' + today_format)
"""