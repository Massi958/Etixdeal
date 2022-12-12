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

 NOM_SOCIETEe="" 
 ADRESSE_SIEGE=""
 NOM_RCS=""
 NUMERO_RCS=""
 REPRESENTANT_ENTREPRISE_SOUS_TRAITANT=""
 with col1:
    use_column_width=500
    st.header("ETIXWAY INFORMATION")
    DATE_DEBUT_ANCIEN_CONTRAT =input_feature = col1.date_input("DATE_DEBUT_ANCIEN_CONTRAT")
    DATE_DEBUT_NOUVEAU_CONTRAT =input_feature = col1.date_input('DATE_DEBUT_NOUVEAU_CONTRAT')
    DATE_FIN_NOUVEAU_CONTRAT =input_feature = col1.date_input('DATE_FIN_NOUVEAU_CONTRAT')
    DATE_DEBUT_CONTRAT =input_feature = col1.date_input('DATE_DEBUT_CONTRAT')
    

 local_css("style.css")
 with col2:
    st.header("INFORMATION SOUS-TRAITANT/COLLABORATEUR")
    genre = col2.radio("",('Mr','Mme'))
    NOM_PRENOM_SOUS_TRAITANT =input_feature = col2.text_input('NOM_PRENOM_SOUS_TRAITANT',"")
    NOM_SOCIETE =input_feature = col2.text_input('NOM_SOCIETE',NOM_SOCIETEe)
    NOM_RCS =input_feature = col2.text_input('NOM_RCS',NOM_RCS)
    NUMERO_RCS =input_feature = col2.text_input('NUMERO_RCS',NUMERO_RCS)
    ADRESSE_SIEGE =input_feature = col2.text_input('ADRESSE_SIEGE',ADRESSE_SIEGE)
    REPRESENTANT_ENTREPRISE_SOUS_TRAITANT =input_feature = col2.text_input('REPRESENTANT_ENTREPRISE_SOUS_TRAITANT',REPRESENTANT_ENTREPRISE_SOUS_TRAITANT)
    QUALITE =input_feature = col2.selectbox('QUALITE',('Président', 'Gérant'))
    # PRESTATION_DEMANDE =input_feature = col2.text_input('PRESTATION_DEMANDE')


 
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
 date_format = '%d/%m/%Y'
 DATE_DEBUT_NOUVEAU_CONTRAT = DATE_DEBUT_NOUVEAU_CONTRAT.strftime("%d/%m/%Y")
 DATE_DEBUT_ANCIEN_CONTRAT= DATE_DEBUT_ANCIEN_CONTRAT.strftime("%d/%m/%Y")
 DATE_FIN_NOUVEAU_CONTRAT = DATE_FIN_NOUVEAU_CONTRAT.strftime("%d/%m/%Y")
 DATE_DEBUT_CONTRAT = DATE_DEBUT_CONTRAT.strftime("%d/%m/%Y")

 dtObj = datetime.strptime(DATE_DEBUT_NOUVEAU_CONTRAT, date_format)
 # from dateutil.relativedelta import relativedelta 

 with st.container():
 # cheminDoc="C:\Users\massi\Desktop\project"+Titre_Doc
  if st.button('Finaliser le contrat'):
          document = DocxTemplate("TEMPLATE Avenant.docx")
          context = {"SEXE":SEXE,"DATE_DEBUT_CONTRAT":DATE_DEBUT_CONTRAT,"NOM_SOCIETE":NOM_SOCIETE,"NOM_PRENOM_SOUS_TRAITANT":NOM_PRENOM_SOUS_TRAITANT,"NOM_RCS":NOM_RCS,"NUMERO_RCS":NUMERO_RCS,"ADRESSE_SIEGE":ADRESSE_SIEGE,"REPRESENTANT_ENTREPRISE_SOUS_TRAITANT":REPRESENTANT_ENTREPRISE_SOUS_TRAITANT,"QUALITE":QUALITE,"DATE_DEBUT_ANCIEN_CONTRAT":DATE_DEBUT_ANCIEN_CONTRAT,"DATE_DEBUT_NOUVEAU_CONTRAT":DATE_DEBUT_NOUVEAU_CONTRAT,"DATE_FIN_NOUVEAU_CONTRAT":DATE_FIN_NOUVEAU_CONTRAT}
          document.render(context)
          Titre_Doc="Contrat Avenant "+NOM_SOCIETE+".docx"
          document.save(Titre_Doc)
          
          inputFile = "Contrat Avenant "+NOM_SOCIETE+".docx"
        
          st.success('Le contrat a bien été crée ! Veuillez maintenant rentrer l\'adresse mail de reception du contrat ')
          dict1 = inputFile
          file1 = open("Original.txt", "w") 
          str1 = repr(dict1)
          file1.write(str1)
          file1.close()
          
          f = open('Original.txt', 'r')
        
                     

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


