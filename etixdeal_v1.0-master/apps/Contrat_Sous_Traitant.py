import streamlit as st
from docxtpl import DocxTemplate

import docx
from email import header
from email import message
from nbformat import write
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from PIL import Image
import base64
import pyautogui
from dateutil.relativedelta import relativedelta
import os

import keyboard
import threading


def app():
 



 
 def threadFunc():
   keyboard.press_and_release('ctrl+w')
 image = Image.open('etixdeal.png')
 st.text("""ETIXDEAL""")
 st.image(image)
 





 col1,col2= st.columns(2)
 def local_css(file_name):
    with open(file_name) as f:
        st.markdown('<style>{}</style>'.format(f.read()), unsafe_allow_html=True)

 local_css("style.css")



 wdFormatPDF = 17

 NOM_SOCIETE=""
 ADRESSE_SIEGE=""
 NOM_RCS=""
 NUMERO_RCS=""
 REPRESENTANT_ENTREPRISE_SOUS_TRAITANT=""
 with col1:
    use_column_width=500
    st.header("ETIXWAY INFORMATION")
    DATE_REDACTION =input_feature = col1.date_input("DATE_REDACTION")
    DATE_DEBUT_CONTRAT =input_feature = col1.date_input('DATE_DEBUT_CONTRAT')
    DATE_DEBUT_PRESTATION =input_feature = col1.date_input('DATE_DEBUT_PRESTATION')
    DUREE_PRESTATION=col1.selectbox('DUREE_PRESTATION_EN_MOIS',('1 mois', '3 mois','6 mois','1 an'))
    CLIENT_FINAL =input_feature = col1.text_input('CLIENT_FINAL')
    INTERLOCUTEUR_ETIXWAY =input_feature = col1.selectbox('INTERLOCUTEUR_ETIXWAY',('OUMAIMA BENYOUSSEF', 'ANOUAR MEHJOUB','ALI CHEKHAB'))
    LIEU_EXECUTION_PRESTATION =input_feature = col1.text_input('LIEU_EXECUTION_PRESTATION')
    TARIF_ACHAT_JOURNALIER =input_feature = col1.number_input('TARIF_ACHAT_JOURNALIER',step=1)

 if INTERLOCUTEUR_ETIXWAY=='OUMAIMA BENYOUSSEF':
      INTERLOCUTEUR_TELEPHONE_ETIXWAY= "07 65 82 55 37"
      INTERLOCUTEUR_MAIL_ETIXWAY="oumaima.benyoussef@etixway.com"
 elif INTERLOCUTEUR_ETIXWAY=='ANOUAR MEHJOUB':
      INTERLOCUTEUR_TELEPHONE_ETIXWAY= "06 99 12 42 41"
      INTERLOCUTEUR_MAIL_ETIXWAY="anouar.mehjoub@etixway.com"
 else:
      INTERLOCUTEUR_TELEPHONE_ETIXWAY= "06 66 14 35 94 "
      INTERLOCUTEUR_MAIL_ETIXWAY="ali.chekhab@etixway.com"
 local_css("style.css")
 with col2:
    st.header("INFORMATION SOUS-TRAITANT")
    genre = col2.radio("",('Mr','Mme'))
    NOM_PRENOM_SOUS_TRAITANT =input_feature = col2.text_input('NOM_PRENOM_SOUS_TRAITANT')
    NOM_SOCIETE =input_feature = col2.text_input('NOM_SOCIETE',NOM_SOCIETE)
    NOM_RCS =input_feature = col2.text_input('NOM_RCS',NOM_RCS)
    NUMERO_RCS =input_feature = col2.text_input('NUMERO_RCS',NUMERO_RCS)
    ADRESSE_SIEGE =input_feature = col2.text_input('ADRESSE_SIEGE',ADRESSE_SIEGE)
    REPRESENTANT_ENTREPRISE_SOUS_TRAITANT =input_feature = col2.text_input('REPRESENTANT_ENTREPRISE_SOUS_TRAITANT',REPRESENTANT_ENTREPRISE_SOUS_TRAITANT)
    QUALITE =input_feature = col2.selectbox('QUALITE',('Président', 'Gérant'))
    TYPE_DE_PROFIL =input_feature = st.selectbox('TYPE_DE_PROFIL',('Le chef de projet technique', 'L\'architecte logiciel', 'L\'ingenieur systeme et reseau', 'Le developpeur fullstack', 'L\'ingenieur cypersecurité ', 'L\'ingenieur DevOps', 'L\'ingenieur fullstack '))
    PRESTATION_DEMANDE =input_feature = col2.text_input('PRESTATION_DEMANDE')
    TELEPHONE_SOUS_TRAITANT =input_feature = col2.text_input('TELEPHONE_SOUS_TRAITANT',max_chars=10)
    MAIL_SOUS_TRAITANT =input_feature = col2.text_input('MAIL_SOUS_TRAITANT')



    
 TARIF_ACHAT_JOURNALIER=float(TARIF_ACHAT_JOURNALIER)
 Titre_Doc="Faux.docx"
 str1=contents= open('Original.txt', 'r').read()

 if genre == 'Mr':
     SEXE="Monsieur"

 else:
     SEXE="Madame"
     if QUALITE=="Président":
        QUALITE ="Présidente"
     else:
          QUALITE="Gérante"
# Titre_Doc="Contrat Consultant ETIXWAY pour "+NOM_SOUS_TRAITANT+".docx"

 DATE_DEBUT_CONTRAT = DATE_DEBUT_CONTRAT.strftime("%d/%m/%Y")
 DATE_REDACTION= DATE_REDACTION.strftime("%d/%m/%Y")
 DATE_DEBUT_PRESTATION= DATE_DEBUT_PRESTATION.strftime("%d/%m/%Y")
 date_format = '%d/%m/%Y'
 dtObj = datetime.strptime(DATE_DEBUT_CONTRAT, date_format)
 if DUREE_PRESTATION=='1 mois':
  n = 1
 elif DUREE_PRESTATION=='3 mois':
  n= 3
 elif DUREE_PRESTATION=='6 mois':
  n=6
 else:
  n=12
 future_date = dtObj + relativedelta(months=n)
 DATE_FIN_PRESTATION = future_date.strftime(date_format)
 TELEPHONE_SOUS_TRAITANT=TELEPHONE_SOUS_TRAITANT[0:2]+" "+TELEPHONE_SOUS_TRAITANT[2:4]+" "+TELEPHONE_SOUS_TRAITANT[4:6]+" "+TELEPHONE_SOUS_TRAITANT[6:8]+" "+TELEPHONE_SOUS_TRAITANT[8:10]
# from dateutil.relativedelta import relativedelta 
# DATE_FIN_PRESTATION = DATE_DEBUT_PRESTATION+ relativedelta(months=3) 
 with st.container():
# cheminDoc="C:\Users\massi\Desktop\project"+Titre_Doc
  if st.button('Finaliser le contrat'):
          document = DocxTemplate("TEMPLATE Contrat Consultant ETIXWAY pour.docx")
          context = {"SEXE":SEXE,"NOM_SOCIETE":NOM_SOCIETE,"NOM_PRENOM_SOUS_TRAITANT":NOM_PRENOM_SOUS_TRAITANT,"NOM_RCS":NOM_RCS,"NUMERO_RCS":NUMERO_RCS,"ADRESSE_SIEGE":ADRESSE_SIEGE,"REPRESENTANT_ENTREPRISE_SOUS_TRAITANT":REPRESENTANT_ENTREPRISE_SOUS_TRAITANT,"QUALITE":QUALITE,"TYPE_DE_PROFIL":TYPE_DE_PROFIL,"DATE_REDACTION":DATE_REDACTION,"CLIENT_FINAL":CLIENT_FINAL,"PRESTATION_DEMANDE":PRESTATION_DEMANDE,"DATE_DEBUT_CONTRAT":DATE_DEBUT_CONTRAT,"DATE_DEBUT_PRESTATION":DATE_DEBUT_PRESTATION,"DATE_FIN_PRESTATION":DATE_FIN_PRESTATION,"TELEPHONE_SOUS_TRAITANT":TELEPHONE_SOUS_TRAITANT,"MAIL_SOUS_TRAITANT":MAIL_SOUS_TRAITANT,"INTERLOCUTEUR_ETIXWAY":INTERLOCUTEUR_ETIXWAY,"INTERLOCUTEUR_TELEPHONE_ETIXWAY":INTERLOCUTEUR_TELEPHONE_ETIXWAY,"INTERLOCUTEUR_MAIL_ETIXWAY":INTERLOCUTEUR_MAIL_ETIXWAY,"LIEU_EXECUTION_PRESTATION":LIEU_EXECUTION_PRESTATION,"TARIF_ACHAT_JOURNALIER":TARIF_ACHAT_JOURNALIER,}
          document.render(context)
          Titre_Doc="Contrat Consultant ETIXWAY pour "+NOM_PRENOM_SOUS_TRAITANT+".docx"
          document.save(Titre_Doc)
          
          inputFile = "Contrat Consultant ETIXWAY pour "+NOM_PRENOM_SOUS_TRAITANT+".docx"
          
         
          st.success('Le contrat a bien été crée ! Veuillez maintenant rentrer l\'adresse mail de reception du contrat ')
          dict1 = inputFile
          file1 = open("Original.txt", "w") 
          str1 = repr(dict1)
          file1.write(str1)
          file1.close()
 
          f = open('Original.txt', 'r')
          if f.mode=='r':
            contents= f.read()

 str1=contents= open('Original.txt', 'r').read()
 str1=str1[1:-1]
 with st.expander("Envoyer le contrat par mail"):
     MAIL_A_ENVOYER=st.text_input('Mail de receptionneur')
     if st.button('Envoyer le contrat par Mail'):
                  Fromadd = "contact@etixway.com"
                  Toadd = MAIL_A_ENVOYER   ##  Spécification des destinataires
                  message = MIMEMultipart()    ## Création de l'objet "message"
                  message['From'] = Fromadd    ## Spécification de l'expéditeur
                  message['To'] = Toadd    ## Attache du destinataire à l'objet "message"
                  message['Subject'] = "Contrat"    ## Spécification de l'objet de votre mail
                  msg = """
Bonjour, 
Je vous prie de trouver ci-joint votre contrat de sous traitance pour signature.
Bonne réception

"""    ## Message à envoyer
                  message.attach(MIMEText(msg.encode('utf-8'), 'plain', 'utf-8'))    ## Attache du message à l'objet "message", et encodage en UTF-8


                  NOM_fichier = str1    ## Spécification du NOM de la pièce jointe
                  piece = open(str1,"rb")    ## Ouverture du fichier
                  part = MIMEBase('application','octet-stream')    ## Encodage de la pièce jointe en Base64
                  part.set_payload((piece).read())
                  encoders.encode_base64(part)
                  part.add_header('Content-Disposition', "piece; filename= %s" % NOM_fichier)
                  message.attach(part)    ## Attache de la pièce jointe à l'objet "message" 

                  
                  
                  serveur = smtplib.SMTP('ns0.ovh.net', 5025)    ## Connexion au serveur sortant (en précisant son NOM et son port)
                  serveur.starttls()    ## Spécification de la sécurisation
                  serveur.login(Fromadd, "De@l!2022!")    ## Authentification
                  texte= message.as_string().encode('utf-8')    ## Conversion de l'objet "message" en chaine de caractère et encodage en UTF-8
                  serveur.sendmail(Fromadd, Toadd, texte)    ## Envoi du mail
                  serveur.quit()    ## Déconnexion du serveur
                  st.success('Le contrat a bien été envoyer à: \n'+  MAIL_A_ENVOYER )





 if st.button("Remplir un nouveau contrat"):
    pyautogui.hotkey("ctrl","F5")

