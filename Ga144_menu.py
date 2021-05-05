# -*- coding: utf-8 -*-
# 03/05/2021
# programme adapte pour Windows
# notepad++  lancer un terminal Run : cmd en mode administrateur
# configurer notepad++ pour pouvoir etre appelle de n'importe quel repertoire
# set path=%path%;c:\Program Files (x86)\Notepad++
# il faut faire attention d avoir le keyboard langage sur US pour arrayforth
# la transformation de colorforth vers forth et forth vers colorforth ,
# le source est http://unternet.net/src/clusterFix/
# les programmes python2 ont ete transforme en :
# cf2f_f2cf.exe cf2f okadback.cf okad.f        ou       cf2f_f2cf.exe f2cf okad.f okadback.cf
# donc nous pouvons utiliser ces outils sans devoir le reecrires en python3
# pour la gestion du clavier on utilise pyautogui
# pour changer le type de keyboard language on utilise detectkeyboardlayout.py
# qui repose sur user32.GetKeyboardLayout pour le get
# et sur win32api.SendMessage(handle, WM_INPUTLANGCHANGEREQUEST, 0, language_id) pour le set
# arrayforth sauvegarde le fichier decompresse en okadback.cf
# depuis un editeur ici notepad++ , la conversion sera alors le fichier okad.f
# le repertoire de travail est "C:/GreenArrays/EVB002"


import os
import time

import pyautogui as keyboard
import win32gui
from win32com import client

from detectkeyboardlayout import *

shell = client.Dispatch("WScript.Shell")
cheminarrayforth = "C:/GreenArrays/EVB002"
programme_arrayforth = "Okad.bat"

'''
C:\GreenArrays\EVB002>cf2f_f2cf.exe cf2f okadback.cf okad.f
(' colorforth  to  forth ', ['okadback.cf', 'okad.f'])

C:\GreenArrays\EVB002>cf2f_f2cf.exe f2cf okad.f okadback.cf
(' forth to   colorforth  ', ['okad.f', 'okadback.cf'])
'''
programme_cf2f = 'cf2f_f2cf.exe cf2f okadback.cf okad.f'  # conversion cf to f
programme_f2cf = 'cf2f_f2cf.exe f2cf okad.f okadback.cf'  # conversion f to cf
programmeEditeur = "notepad++.exe okad.f"
arrayforth = "colorForth"  # nom de la fenetre active du prg ArrayForth

notepad = "C:\GreenArrays\EVB002\okad.f - Notepad++"
tsmall = 0.1  # 100ms
tlong = 1
os.chdir(cheminarrayforth)  # on va travailler dans le repertoire de Colorforth


# -------------------------------------------------------------------
# ---------------------- fonctions ----------------------------------
# -------------------------------------------------------------------

def reactivewindow(hdnle):
    shell.SendKeys('%')  # ALT Key
    win32gui.SetForegroundWindow(hdnle)  # windows activation
    time.sleep(tlong)


def sendMesg(message):
    keyboard.write(message, tsmall)

def idColorForth():
    h = win32gui.FindWindow(None, arrayforth)  # recupere le handle de la fenetre arrayforth
    # print("id arrayforth", h)
    return h

def idnotepad():
    h = win32gui.FindWindow(None, notepad)  # recupere le handle de la fenetre notepad
    return h


def idMenu():
    h = win32gui.FindWindow(None, "Ga144_menu")

    return h


def ArrayForth():
    get_set_keyboard(keyboardUS)  # on passe la clavier en US
    print("\n Lancer ArrayForth")
    shell.Run(programme_arrayforth)
    time.sleep(tlong)
    id = idColorForth()  # recuperation id de ColorForth
    if id == 0:
        print("\n ArrayForth introuvable...")
    else:
        # print("\n La fenetre colorForth a %s " % id, " suivant ")
        reactivewindow(id)  # rend actif la fenetre ColorForth
    print("Ok\n")


def Editeur():
    get_set_keyboard(keyboard_init)  # on passe la clavier different de US s il le faut
    print("\n Lancer editeur Notepad++ ")
    shell.Run(programmeEditeur)
    time.sleep(tlong)
    print("Ok\n")


def Bye():
    if idColorForth():
        chaine = " bye "  # bye
        CommandeArrayForth(chaine)

    print("\n bye")
    quit()


def CommandeArrayForth(chaine):
    id = idColorForth()  # recuperation id de ColorForth
    if id == 0:
        print("\n ArrayForth introuvable...")
    else:
        get_set_keyboard(keyboardUS)  # on passe la clavier en US
        # print("\n La fenetre colorForth a %s " % id, " suivant ")
        reactivewindow(id)  # rend actif la fenetre ColorForth
        time.sleep(tlong)
        sendMesg(chaine)


def Commandenotepad():
    id = idnotepad()  # recuperation id de notepad
    if id == 0:
        print("\n notepad introuvable...")
    else:
        get_set_keyboard(keyboard_init)  # on passe la clavier different de US s il le faut
        # print("\n La fenetre notepad a %s " % id, " suivant ")
        reactivewindow(id)  # rend actif la fenetre notepad
        time.sleep(tlong)


def InitArrayForth():
    get_set_keyboard(keyboardUS)  # on passe la clavier en US
    time.sleep(tlong)

    print("\n Initialisation ArrayForth")
    chaine = " c 0 "  #
    CommandeArrayForth(chaine)
    time.sleep(tlong)
    chaine = "!back "  # 0 !back sauvegarde 1440 blocs vers okadback.cf
    CommandeArrayForth(chaine)
    print("\n  Sauvegarde okadback.cf: ")


def MajArrayForth():
    get_set_keyboard(keyboardUS)  # on passe la clavier en US
    chaine = " c 0 "  # 0 @back recharge  1440 blocs
    CommandeArrayForth(chaine)
    time.sleep(tlong)
    chaine = "@back "  # 0 @back recharge  1440 blocs de okadback.cf
    CommandeArrayForth(chaine)
    time.sleep(tlong)

    print("\n  mise a jour okadback.cf ")


def Commande():
    get_set_keyboard(keyboard_init)  # on passe la clavier different de US s il le faut
    chaine = input("\n Commande a envoyer :")
    get_set_keyboard(keyboardUS)  # on passe la clavier en US
    CommandeArrayForth(chaine)


def ConversionCF_toForth():
    os.chdir(cheminarrayforth)
    print("\n Conversion ColorForth vers Forth")
    InitArrayForth()
    os.system(programme_cf2f)
    print("\n fin conversion")
    Commandenotepad()



def ConversionForth_toCF():
    os.chdir(cheminarrayforth)
    print("\n Conversion  Forth vers ColorForth")
    os.system(programme_f2cf)
    print("\n fin conversion")




def erreur():
    print("\n erreur saisie menu")
    print("\n" * 100)


def get_set_keyboard(keyboard):
    if get_keyboard_language() == keyboard:
        print("clavier : ", keyboard)
    else:
        key = (int(keyboard, 16))
        set_keyboard_language(key)
        print("configuration clavier   ", hex(key))


def clear_terminal():
    os.system("cls")



# -------------------------------------------------------------------
# ----------------------- Menu --------------------------------------
# -------------------------------------------------------------------
def menu():


    print("-------------------------------------------------------------")
    print("|                     ArrayForth GA144                      |")
    print("|  0) Commande a envoyer notepad                            |")
    print("|  1) Lancer ArrayForth                                     |")
    print("|  2) Lancer editeur Notepad++                              |")
    print("|  3) okadback.cf ---> okadback.f  vers editeur Notepad++   |")
    print("|  4) okadback.f  ---> okadback.cf vers ArrayForth          |")
    print("|  5) mise a jour donnees  ArrayForth                       |")
    print("|  6) Commande a envoyer                                    |")
    print("|  8) Initialiser ArrayForth (sauvegarde okadback.cf)       |")
    print("|  9) Quitter                                               |")
    print("-------------------------------------------------------------")
    choix = input("    commande : ")
    option = {'0': Commandenotepad,
              '1': ArrayForth,
              '2': Editeur,
              '3': ConversionCF_toForth,
              '4': ConversionForth_toCF,
              '5': MajArrayForth,
              '6': Commande,
              '8': InitArrayForth,
              '9': Bye
              }
    option.get(choix, erreur)()
    clear_terminal()

# on vient sauvegarder le type de clavier , et on le passe en US
keyboardUS = '0x409' # keyboard US
keyboard_init = get_keyboard_language()
print("clavier  initial ", keyboard_init)
get_set_keyboard(keyboardUS)

while True:
    menu()
