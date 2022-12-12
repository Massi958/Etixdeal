import streamlit as st
from docxtpl import DocxTemplate

import docx
from email import header
from email import message
from nbformat import write
import streamlit as st
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
 

 import keyboard
 import threading
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



 with col1:
    use_column_width=500
    st.header("ETIXWAY INFORMATION")
    DATE_REDACTION =input_feature = col1.date_input("DATE_REDACTION")
    DATE_DEBUT_CONTRAT =input_feature = col1.date_input('DATE_DEBUT_CONTRAT')
    DATE_DEBUT_ANCIEN_CONTRAT =input_feature = col1.date_input('DATE_DEBUT_ANCIEN_CONTRAT')
    
    
    

 local_css("style.css")
 with col2:
    st.header("INFORMATION COLLABORATEUR")
    genre = col2.radio("",('Mr','Mme'))
    NOM_PRENOM =input_feature = col2.text_input('NOM_PRENOM')
    NUMERO_SECURITE_SOCIAL =input_feature = col2.text_input('NUMERO_SECURITE_SOCIAL')
    ADRESSE =input_feature = col2.text_input('ADRESSE')
    QUALITE =input_feature = col2.selectbox('QUALITE',('Ingénieur développement', 'Ingénieur DevOps','Ingérnieur testeur','Ingénieur Data analyste','Chef de Projet','Chargé des ressources humaines','Chargé Administration des ventes','Consultant Recrutement','Chargé de Recrutement ','Ingénieur Reseau et Securité','Ingénieur Data scientist'))
    POSITION =input_feature = col2.selectbox('POSITION',('1.1','1.2','2.1','2.2','2.3','3.1','3.2','3.3'))
    STATUT =input_feature = col2.selectbox('STATUT',('Cadre','Non Cadre'))
    REMUNERATION =input_feature = col1.number_input('REMUNERATION',step=1)
    REMUNERATION_EN_LETTRE =input_feature = col1.text_input('REMUNERATION_EN_LETTRE')
    NOMBRE_JOUR_TELETRAVAIL_PAR_SEMAINE= col2.text_input('NOMBRE_JOUR_TELETRAVAIL_PAR_SEMAINE')
    
 Titre_Doc="Faux.docx"
 str1=contents= open('Original.txt', 'r').read()
 if POSITION=='1.1':
    COEFFICIENT='95'
 elif POSITION=='1.2':
    COEFFICIENT='100'
 elif POSITION=='2.1':
    COEFFICIENT='115'
 elif POSITION=='2.2':
    COEFFICIENT='130'
 elif POSITION=='2.3':
    COEFFICIENT='150'
 elif POSITION=='3.1':
    COEFFICIENT='170'
 elif POSITION=='3.2':
    COEFFICIENT='210'
 else:
    COEFICIENT='270'
 if genre == 'Mr':
     SEXE="Monsieur"
     PRONOM="le salarié"
 else:
     SEXE="Madame"
     PRONOM="la salariée"
     if QUALITE=="Président":
        QUALITE ="Présidente"
        
     else:
          QUALITE="Gérante"
# Titre_Doc="Contrat Consultant ETIXWAY pour "+NOM_SOUS_TRAITANT+".docx"
 date_format = '%d/%m/%Y'
 DATE_DEBUT_CONTRAT = DATE_DEBUT_CONTRAT.strftime("%d/%m/%Y")
 DATE_REDACTION= DATE_REDACTION.strftime("%d/%m/%Y")
 DATE_DEBUT_ANCIEN_CONTRAT = DATE_DEBUT_ANCIEN_CONTRAT.strftime("%d/%m/%Y")


 dtObj = datetime.strptime(DATE_DEBUT_CONTRAT, date_format)
# from dateutil.relativedelta import relativedelta 

 with st.container():
# cheminDoc="C:\Users\massi\Desktop\project"+Titre_Doc
  if st.button('Finaliser le contrat'):
          document = DocxTemplate("TEMPLATE AVENANT_CDI.docx")
          context = {"NUMERO_SECURITE_SOCIAL":NUMERO_SECURITE_SOCIAL,"DATE_DEBUT_ANCIEN_CONTRAT":DATE_DEBUT_ANCIEN_CONTRAT,"REMUNERATION":REMUNERATION,"NOMBRE_JOUR_TELETRAVAIL_PAR_SEMAINE":NOMBRE_JOUR_TELETRAVAIL_PAR_SEMAINE,"REMUNERATION_EN_LETTRE":REMUNERATION_EN_LETTRE,"PRONOM":PRONOM,"STATUT":STATUT,"POSITION":POSITION,"COEFFICIENT":COEFFICIENT,"SEXE":SEXE,"NOM_PRENOM":NOM_PRENOM,"ADRESSE":ADRESSE,"QUALITE":QUALITE,"DATE_REDACTION":DATE_REDACTION,"DATE_DEBUT_CONTRAT":DATE_DEBUT_CONTRAT}
          document.render(context)
          Titre_Doc="Contrat AVENANT CDI "+NOM_PRENOM+".docx"
          document.save(Titre_Doc)
          
          inputFile = "Contrat AVENANT CDI "+NOM_PRENOM+".docx"
          

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

