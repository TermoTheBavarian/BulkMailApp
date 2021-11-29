import pandas as pd
import pathlib
from time import sleep
import win32com.client as client
import PySimpleGUI as sg
import random
from _codecs import *
import mammoth
def show_app():
    image_path= "image.png"
    layout= [
                    [sg.Text("Welcome to the Bulk Maling App. Make sure you are logged in Outlook")],
                    [sg.Text("Introduce Subject Text (Recipient´s Name will be added at the end)"), sg.Input("", key="subject")],
                    [sg.Text("Introduce attached PDF file (Optional)"),sg.Input("", key= "pdf"), sg.FileBrowse()],
                    [sg.Text("Introduce Exel Agenda"), sg.Input("", key="excel"), sg.FileBrowse()],
                    [sg.Text("Introduce text message in Word format"), sg.Input("", key="text"), sg.FileBrowse()],
                    [sg.Button("Mail Preview", key="draft")],
                    [sg.Image(image_path)]

            ],

    window = sg.Window("Mailing App", layout)
    return window
def main():
    run_app= True
    while run_app:
        window= show_app()
        event, values= window.read()
        if event == sg.WIN_CLOSED:
            window.close()
            run_app = False
            break
        input_subject = (values["subject"])
        input_pdf = (values["pdf"])
        input_excel = (values["excel"])
        path_text = (values["text"])

        if input_pdf == "" and input_excel == "":
            sg.popup("!You have not introduced anything!")

        elif input_excel == "":
            sg.popup("!You have not introduced an Excel Agenda")

        elif event == ("draft" and  path_text != "" and input_excel != "" and input_pdf != "") or "draft" and  path_text != "" and input_excel != "":
            window.close()
            if input_pdf == "":
                sg.popup("!Remember that you have not attach any PDF!")
            #with open(path_text) as f:
                    #input = f.readlines()
                    #input_text = "".join(input)

            with open(path_text, "rb") as docx_file:
                input = mammoth.convert_to_html(docx_file)
                input_text= input.value

            draft_mail(input_excel,input_text, input_pdf, input_subject)
            confirm_window(input_excel,input_text, input_pdf,input_subject)


def confirm_window(input_excel,input_text, input_pdf= "", input_subject=""):
    path = input_excel
    reader = pd.read_excel(path)
    names = reader["Name"].values
    emails = reader["mail"].values
    lista = ("")
    for x in range(len(names)):
        lista += str(names[x])+"-"+ str(emails[x])+"\n"

    new_layout = [[sg.Text("Close the draft mail. You are about to send the message to the following people:", key="new")],[sg.Button("Send", key="submit")],
                  [sg.Text(lista, key="confirm",size=(80, 20))]

                  ]
    new_window = sg.Window("Second Window", new_layout)
    second_window = True
    while second_window:
        event, values = new_window.read()
        if event == sg.WIN_CLOSED:
            break
        if event == ("submit"):

            send_mail(input_excel,input_text,input_pdf,input_subject)
            sg.Popup("!Campaign Finished Successfully!")
            new_window.close()
            second_window = False


def draft_mail(input_excel,input_text, input_pdf= "", input_subject= ""):
    path = input_excel
    reader = pd.read_excel(path)
    names = reader["Name"].values
    emails = reader["mail"].values
    """chunks= [distro[x:x+500]for x in range(0,len(distro),500)]"""
    if input_pdf != "":
        dossier_ecosen = pathlib.Path(input_pdf)
        dossier_absolute = str(dossier_ecosen.absolute())
    outlook = client.Dispatch("Outlook.Application")


    for x in range(1):
        message = outlook.CreateItem(0)
        message.To = (emails[x])
        message.Subject = input_subject+ " {}".format(names[x])
        message.HTMLBody = input_text
        if input_pdf != "":
            message.Attachments.Add(dossier_absolute)


        message.Display()
        sleep(5)
        outlook.Quit()

    return

def send_mail(input_excel,input_text, input_pdf = "", input_subject= ""):

        path= input_excel
        reader= pd.read_excel(path)
        names = reader["Name"].values
        emails= reader["mail"].values

        """chunks= [distro[x:x+500]for x in range(0,len(distro),500)]"""
        if input_pdf != "":
            dossier_ecosen= pathlib.Path(input_pdf)
            dossier_absolute= str(dossier_ecosen.absolute())
        outlook= client.Dispatch("Outlook.Application")


        for x in range(len(names)):
                sg.PopupTimed("Sending mail to {}...".format(names[x]), auto_close_duration=50 )
                message = outlook.CreateItem(0)
                message.To = (emails[x])
                message.Subject = input_subject+ " {}".format(names[x])
                message.HTMLBody = input_text
                if input_pdf != "":
                    message.Attachments.Add(dossier_absolute)

                message.Send()
                sleep(60)



if __name__ == "__main__":
    main()
