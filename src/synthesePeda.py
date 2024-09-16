#!/usr/bin/env python
# -*- coding: utf-8 -*-

##This file is part of pySequence
#############################################################################
#############################################################################
##                                                                         ##
##                               SynthesePeda                               ##
##                                                                         ##
#############################################################################
#############################################################################

## Copyright (C) 2014 Cédrick FAURY - Jean-Claude FRICOU
##
## pySéquence : aide à la construction
## de Séquences et Progressions pédagogiques
## et à la validation de Projets

#    pySequence is free software; you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation; either version 3 of the License, or
#    (at your option) any later version.
    
#    pySequence is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with pySequence; if not, write to the Free Software
#    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA


u"""
Module SynthesePeda
*******************

Génération des tableaux de synthèse pédagogique des Progressions : 
Compétences, Centres d'intérêts , ...


"""

import wx
import sys


import xlwt
from xlwt import Workbook, Font, XFStyle, Borders, Alignment, Formula, Pattern

import constantes
import util_path

from Referentiel import REFERENTIELS
import Referentiel

import images

import os, glob

from widgets import messageErreur, getHoraireTxt

import objects_wx
import pysequence

# Pour enregistrer en xml
import xml.etree.ElementTree as ET

import  wx.lib.mixins.listctrl  as  listmix

#################################################################################################################
#
# Synthèse pédagogique (sur pluieures séquences)
#
#################################################################################################################
class FenetreBilan(wx.Frame):
    def __init__(self, parent, dossierCourant = util_path.PATH, 
                 referentiel = None):
        wx.Frame.__init__(self, parent, -1, u"Import de séquence", style = wx.DEFAULT_FRAME_STYLE|wx.STAY_ON_TOP)
        
        self.pourProjet = False
        self.lastPath = None
        if referentiel == None:
            referentiel = REFERENTIELS[constantes.TYPE_ENSEIGNEMENT_DEFAUT]
        self.sizer = wx.GridBagSizer()
        panel = wx.Panel(self, -1)
        sizer = wx.BoxSizer()
        self.SetIcon(images.getlogoIcon())
        
        #############################################################################################
        # Cr�ation de la barre d'outils
        #############################################################################################
        self.ConstruireTb()
        
        self.lstClasse = []
        self.lstSeq = []
            
        #
        # Type d'enseignement
        #
        titre = wx.StaticBox(panel, -1, u"Type d'enseignement")
        sb = wx.StaticBoxSizer(titre, wx.VERTICAL)
        te = objects_wx.ArbreTypeEnseignement(panel, self)
        self.Bind(wx.EVT_RADIOBUTTON, self.EvtRadioBox, te)
        sb.Add(te, flag = wx.EXPAND)
        self.referentiel = referentiel
#        l = []
#        for i, e in enumerate(REFERENTIELS.keys()):
#            l.append(REFERENTIELS[e].Enseignement[0])
#        rb = wx.RadioBox(
#                panel, -1, u"Type d'enseignement", wx.DefaultPosition, (130,-1),
#                l,
#                1, wx.RA_SPECIFY_COLS
#                )
#        rb.SetToolTip(wx.ToolTip(u"Choisir le type d'enseignement"))
#        for i, e in enumerate(REFERENTIELS.keys()):
#            rb.SetItemToolTip(i, REFERENTIELS[e].Enseignement[1])
#        self.Bind(wx.EVT_RADIOBOX, self.EvtRadioBox, rb)
#            
#        # Provosoirement uniquement pour SSI
#        for i in [1,2,3,4]:
#            rb.EnableItem(i, False)    
        
        self.sizer.Add(sb, (0,0), (2,1), flag = wx.EXPAND|wx.ALL)
        self.cb_type = te
        te.SetStringSelection(referentiel.Enseignement[0])

        #
        # Dossiers de recherche
        #
        sb = wx.StaticBox(panel, -1, u"Dossiers où chercher les fichiers de séquence")
        sbs = wx.StaticBoxSizer(sb, wx.VERTICAL)
        hs = wx.BoxSizer(wx.HORIZONTAL)
        self.dossiers = [os.path.abspath(dossierCourant)]
        self.dossiersOk = True
        self.txtDoss = wx.TextCtrl(panel, -1, os.path.abspath(dossierCourant))
        self.txtDoss.SetToolTipString(u"Saisir les dossiers de recherche, séparés par \";\"")
#        self.txtDoss.Bind(wx.EVT_KILL_FOCUS, self.EvtTextDoss)
        self.Bind(wx.EVT_TEXT, self.EvtTextDoss, self.txtDoss)
        hs.Add(self.txtDoss, 1, flag = wx.EXPAND|wx.ALL)
        
        self.boutonDoss = wx.Button(panel, -1, "...", size = (30, -1))
        self.boutonDoss.SetToolTipString(u"Selectionner un dossier de recherche")
        self.Bind(wx.EVT_BUTTON, self.OnDossier, self.boutonDoss)
        hs.Add(self.boutonDoss, flag = wx.EXPAND|wx.ALL)
        
        sbs.Add(hs, flag = wx.EXPAND|wx.ALL)
        self.cbr = wx.CheckBox(panel, -1, u"Inclure les sous-dossiers")
        self.Bind(wx.EVT_CHECKBOX, self.EvtCheckBox, self.cbr)
        self.recurs = True
        self.cbr.SetValue(self.recurs)
        sbs.Add(self.cbr, flag = wx.EXPAND|wx.ALL)
        
        self.sizer.Add(sbs, (0,1), (1,1) , flag = wx.EXPAND|wx.ALL)
        self.sizer.AddGrowableCol(1)
        
        #
        #    Liste des fichiers trouv�s
        #
        self.listeSeq = PanelListe(panel, self)

        self.sizer.Add(self.listeSeq, (1,1), (1,1) , flag = wx.EXPAND|wx.ALL)
        self.MiseAJourListe()
        
        panel.SetSizer(self.sizer)
        sizer.Add(panel, 1, flag = wx.EXPAND)
        
        self.SetSizerAndFit(sizer)
        
        self.Bind(wx.EVT_SIZE, self.OnResize)
#        wx.CallAfter(self.cb_type.CollapseAll)
#        wx.CallAfter(self.Thaw)
        
        
    ######################################################################################              
    def EvtCheckBox(self, event):
        self.recurs = event.IsChecked()
        self.MiseAJourListe()
        
    ######################################################################################              
    def OnResize(self, evt = None):
#        print "Resize ArbreTypeEnseignement", self.GetClientSize()
        self.cb_type.SetMinSize((-1, self.GetClientSize()[1]-30))
        self.cb_type.Layout()
        self.sizer.Layout()
        if evt:
            evt.Skip()
        
    ###############################################################################################
    def ConstruireTb(self):
        """ Construction de la ToolBar
        """
#        print "ConstruireTb"

        #############################################################################################
        # Création de la barre d'outils
        #############################################################################################
        self.tb = self.CreateToolBar(wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_FLAT)
        
        tsize = (24,24)
        new_bmp = images.Icone_excel.GetBitmap()
        
        self.tb.SetToolBitmapSize(tsize)
        
        self.tb.AddLabelTool(10, u"Exporter", new_bmp, 
                             shortHelp=u"Exporter la synthèse dans un fichier Excel", 
                             longHelp=u"Exporter la synthèse dans un fichier Excel")
        self.Bind(wx.EVT_TOOL, self.commandeExporter, id=10)    
        
        self.tb.AddLabelTool(30, u"Actualiser", images.Bouton_Actualiser.GetBitmap(), 
                             shortHelp=u"Actualiser la liste des séquences", 
                             longHelp=u"Actualiser la liste des séquences")
        self.Bind(wx.EVT_TOOL, self.MiseAJourListe, id=30)
        
        open_bmp = wx.ArtProvider.GetBitmap(wx.ART_FILE_OPEN, wx.ART_TOOLBAR, tsize)
        self.tool_open = self.tb.AddLabelTool(20, u"Ouvrir", open_bmp, 
                             shortHelp=u"Ouvrir la synthèse", 
                             longHelp=u"Ouvrir le fichier Excel contenant la synthèse")
        self.Bind(wx.EVT_TOOL, self.OuvrirSynthese, id=20)     
        
#        print "fini"

        
        #################################################################################################################
        #
        # Mise en place
        #
        #################################################################################################################
        self.tb.Realize()
        
        self.tb.RemoveTool(20)
#        print "fini 2"
    
    
    ######################################################################################  
    def commandeExporter(self, event = None):
        mesFormats = r"Fichier Excel|*.xls"
        dlg = wx.FileDialog(self, 
                            message = u"Enregistrement de la synthèse", 
                            defaultFile="", wildcard=mesFormats, 
                            style=wx.SAVE|wx.OVERWRITE_PROMPT|wx.CHANGE_DIR
                            )
        dlg.SetFilterIndex(0)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            dlg.Destroy()
#            ok = self.enregistrer(path)
            ok = self.genererBilan(path)
            if ok:
                dlg = wx.MessageDialog(self, u"Export de la synthèse réussi !", u"Export de la synthèse réussi",
                           wx.OK | wx.ICON_INFORMATION)
                dlg.ShowModal()
                dlg.Destroy()
                self.lastPath = path
                self.tb.InsertToolItem(2,self.tool_open)
                self.tb.Realize()
                
            else:
                dlg = wx.MessageDialog(self, u"L'export de la synthèse a échoué !", u"Echec de l'export de la synthèse",
                           wx.OK | wx.ICON_WARNING)
                dlg.ShowModal()
                dlg.Destroy()
        else:
            dlg.Destroy()

    ######################################################################################  
    def OuvrirSynthese(self, event = None):
        if self.lastPath != None and os.path.isfile(self.lastPath):
            try:
                os.startfile(self.lastPath)
            except:
                messageErreur(None, u"Ouverture impossible",
                              u"Impossible d'ouvrir la synthèse\n\n%s\n" %util_path.toSystemEncoding(self.lastPath))

    ######################################################################################  
    def definirStyles(self, book):
        #
        # Couleurs
        #
        xlwt.add_palette_colour("vert1", 8)
        book.set_colour_RGB(8, 120, 200, 120)
        xlwt.add_palette_colour("vert2", 9)
        book.set_colour_RGB(9, 160, 220, 160)
        xlwt.add_palette_colour("vert3", 10)
        book.set_colour_RGB(10, 200, 255, 200)
        
        xlwt.add_palette_colour("rouge1", 18)
        book.set_colour_RGB(18, 200, 120, 120)
        xlwt.add_palette_colour("rouge2", 19)
        book.set_colour_RGB(19, 220, 160, 160)
        xlwt.add_palette_colour("rouge3", 20)
        book.set_colour_RGB(20, 255, 200, 200)
        
        #
        # Styles
        #    
        
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 37
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 36*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER

        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 42
        
        styleT = XFStyle()
        styleT.font = fnt
        styleT.borders = borders
        styleT.alignment = al
        styleT.pattern = pattern
        self.styleT = styleT
        
        ######################################################
        
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 16*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER

        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 1

        styleE = XFStyle()
        styleE.font = fnt
        styleE.borders = borders
        styleE.alignment = al
        styleE.pattern = pattern
        self.styleE = styleE
        
        ######################################################

        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 14*20

        borders = Borders()
        borders.left = 1
        borders.right = 0
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 43

        style0 = XFStyle()
        style0.font = fnt
        style0.borders = borders
        style0.alignment = al
        style0.pattern = pattern
        self.style0 = style0
        
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 14*20

        borders = Borders()
        borders.left = 0
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        al.wrap = True
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 43
        
        style01 = XFStyle()
        style01.font = fnt
        style01.borders = borders
        style01.alignment = al
        style01.pattern = pattern
        self.style01 = style01
        
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 14*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        al.wrap = True
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 43
        
        style02 = XFStyle()
        style02.font = fnt
        style02.borders = borders
        style02.alignment = al
        style02.pattern = pattern
        self.style02 = style02
    
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 12*20

        borders = Borders()
        borders.left = 1
        borders.right = 0
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 41

        style1 = XFStyle()
        style1.font = fnt
        style1.borders = borders
        style1.alignment = al
        style1.pattern = pattern
        self.style1 = style1
        
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = False
        fnt.height = 12*20

        borders = Borders()
        borders.left = 0
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 41
        al.wrap = True

        style11 = XFStyle()
        style11.font = fnt
        style11.borders = borders
        style11.alignment = al
        style11.pattern = pattern
        self.style11 = style11
        
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = False
        fnt.height = 12*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 41
        al.wrap = True

        style12 = XFStyle()
        style12.font = fnt
        style12.borders = borders
        style12.alignment = al
        style12.pattern = pattern
        self.style12 = style12
        
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 10*20

        borders = Borders()
        borders.left = 1
        borders.right = 0
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 1

        style2 = XFStyle()
        style2.font = fnt
        style2.borders = borders
        style2.alignment = al
        style2.pattern = pattern
        self.style2 = style2
        
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = False
        fnt.height = 10*20
        

        borders = Borders()
        borders.left = 0
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        al.wrap = True
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 1

        style21 = XFStyle()
        style21.font = fnt
        style21.borders = borders
        style21.alignment = al
        style21.pattern = pattern
        self.style21 = style21
        
        ######################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = False
        fnt.height = 10*20
        

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_LEFT
        al.vert = Alignment.VERT_CENTER
        al.wrap = True
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 1

        style22 = XFStyle()
        style22.font = fnt
        style22.borders = borders
        style22.alignment = al
        style22.pattern = pattern
        self.style22 = style22
        
        # Croix #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 12*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 47

        styleX = XFStyle()
        styleX.font = fnt
        styleX.borders = borders
        styleX.alignment = al
        styleX.pattern = pattern
        self.styleX = styleX
        
        # rien #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.height = 12*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 1

        stylenX = XFStyle()
        stylenX.font = fnt
        stylenX.borders = borders
        stylenX.alignment = al
        stylenX.pattern = pattern
        self.stylenX = stylenX
        
        # Dur�e de s�quence #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.italic = True
        fnt.height = 9*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 29

        styleD = XFStyle()
        styleD.num_format_str='hh:mm'
        styleD.font = fnt
        styleD.borders = borders
        styleD.alignment = al
        styleD.pattern = pattern
        self.styleD = styleD
        
        
        # P�riode de s�quence #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.italic = False
        fnt.height = 18*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 29

        stylePer = XFStyle()
        stylePer.font = fnt
        stylePer.borders = borders
        stylePer.alignment = al
        stylePer.pattern = pattern
        self.stylePer = stylePer
        
        
        # Num�ro de s�quence #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.italic = False
        fnt.height = 16*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 19

        styleN = XFStyle()
        styleN.font = fnt
        styleN.borders = borders
        styleN.alignment = al
        styleN.pattern = pattern
        self.styleN = styleN
        
        # Intitul� de s�quence #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.italic = False
        fnt.height = 9*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        al.rotation = Alignment.ORIENTATION_90_CC
        
        al.wrap = True
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 20

        styleS = XFStyle()
        styleS.num_format_str='hh:mm'
        styleS.font = fnt
        styleS.borders = borders
        styleS.alignment = al
        styleS.pattern = pattern
        self.styleS = styleS
        
        ######################################################
        
        f = Font()
        f.height = 20*6
        f.name = 'Verdana'
        f.bold = False
        f.underline = Font.UNDERLINE_SINGLE
        f.colour_index = 4
        
        al = Alignment()
        al.horz = Alignment.HORZ_RIGHT
        
        h_style = XFStyle()
        h_style.font = f
        h_style.alignment = al
        self.h_style = h_style
        
        # Position cible CI #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.italic = True
        fnt.height = 12*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 8

        stylePosCI = XFStyle()
        stylePosCI.num_format_str='hh:mm'
        stylePosCI.font = fnt
        stylePosCI.borders = borders
        stylePosCI.alignment = al
        stylePosCI.pattern = pattern
        self.stylePosCI = stylePosCI
        
        # Num�ro de CI #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = True
        fnt.italic = False
        fnt.height = 14*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        al.wrap = False
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 9

        styleNumCI = XFStyle()
        styleNumCI.font = fnt
        styleNumCI.borders = borders
        styleNumCI.alignment = al
        styleNumCI.pattern = pattern
        self.styleNumCI = styleNumCI
        
        # Intitul� de CI #####################################################
        fnt = Font()
        fnt.name = 'Arial'
        fnt.colour_index = 0
        fnt.outline = True
        fnt.struck_out = False
        fnt.bold = False
        fnt.italic = False
        fnt.height = 9*20

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        al = Alignment()
        al.horz = Alignment.HORZ_CENTER
        al.vert = Alignment.VERT_CENTER
        al.rotation = Alignment.ORIENTATION_90_CC
        
        al.wrap = True
        
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 10

        styleCI = XFStyle()
        styleCI.num_format_str='hh:mm'
        styleCI.font = fnt
        styleCI.borders = borders
        styleCI.alignment = al
        styleCI.pattern = pattern
        self.styleCI = styleCI
    
    #######################################################################################  
    def getCIcommuns(self, listeSeq):
        """ Dresse un bilan des CI intégrés dans les s�quences
            >> renvoie : 
            [[[liste de séquences], [liste des CI de ces séquences]], [[autre liste de séquences], [liste des CI de ces séquences]], ...]
        """
        
        def memeCI(lst1, lst2):
            if len(lst1) <> len(lst2):
                return False
            for i, ci in enumerate(lst2):
                if ci != lst1[i]:
                    return False
            return True
        
        
        
        new = True
        listeCI = []
        for seq in listeSeq:
#            print "***", seq
            ref = seq.GetReferentiel()
#            print "   ", ref.CentresInterets
            if len(listeCI) == 0: # Initialisation de la listeCI
                listeCI.append([[seq], list(ref.CentresInterets)])
            else:
                for s, l in listeCI:
                    if memeCI(l, ref.CentresInterets):
                        s.append(seq)
                        new = False
                        break
                if new:
#                    print "   Nouveau >>", seq, "<>", s[-1]
                    listeCI.append([[seq], list(ref.CentresInterets)])
                new = True
                    
        
#        print "listeCI", listeCI
                            
        return listeCI
    
    #######################################################################################  
    def traiterDic(self, dic, ws0, l = 6, c = 0):
        """ Fonction de traitement des dictionnaires (compétences, savoir)
            >>> positionnement sur la partie gauche
        """
        dicLigne = {}
        for c0, v0 in sorted(dic):
            ws0.write(l, c, c0, self.style0)                                     # Code
            ws0.write_merge(l, l, c+1, c+3, v0[0], self.style01)                   # Intitul�
            dicLigne[c0] = [l, [], 0]
            l += 1
            if type(v0[1]) == dict:
                for c1, v1 in sorted(v0[1].items()):
                    ws0.write(l, c+1, c1, self.style1)                           # Code
                    dicLigne[c1] = [l, [], 1]
                    dicLigne[c0][1].append(c1)
                    ws0.write_merge(l, l, c+2, c+3, v1[0], self.style11)           # Intitul�
                    if type(v1[1]) == dict:
                        for c2, v2 in sorted(v1[1].items()):
                            l += 1
                            ws0.write(l, c+2, c2, self.style2)                   # Code
                            dicLigne[c2] = [l, [], 2]
                            dicLigne[c1][1].append(c2)
                            ws0.write_merge(l, l, c+3, c+3, v2[0], self.style21)   # Intitul�
#                                l += 1
                    l += 1
        return dicLigne, l
        
    
    ######################################################################################  
    def reglerLargeursGauche(self, ws0, c, c0 = 0):
        """ Règle la largeur des colonnes de la partie gauche
        """
        #
        # Largeur des colonnes
        #
        ws0.col(c0+0).width = 100*20
        ws0.col(c0+1).width = 100*20
        ws0.col(c0+2).width = 100*20
        ws0.col(c0+3).width = 900*20
        ws0.col(c0+4).width = 15*20
#        for cc in range(c0+5, c0+c):
#            ws0.col(cc).width = 80*20
    
    ######################################################################################  
    def getStyleComp(self, niv):
        if niv == 0:
            return self.style02
        elif niv == 1:
            return self.style12
        elif niv == 2:
            return self.style22
        
                               
    ######################################################################################  
    def genererBilanComplexe(self, wb, comp_sav = 'C'):
        """ Génère une feuille affichant un croisement
                CI/Compétence (comp_sav = 'C')
                CI/Savoirs (comp_sav = 'S')
        """
        
        def traiterSeq(nom, dic):
            ws0 = wb.add_sheet(nom)
            ws0.write_merge(1, 3, 1, 3, self.referentiel.Enseignement[0], self.styleT)
            ws0.write_merge(5, 5, 0, 3, nom, self.styleE)
            dicLigne, last = self.traiterDic(dic.items(), ws0)
          
            #
            # On met les croix
            #
            c = 5
            l = 6
            n = "HYPERLINK"
            for p, lst in enumerate(self.seqTriees):
                for i, seq in enumerate(lst):
                    ws0.write(l-4, c, (i*(p+1))+1, self.styleN)                  # Num�ro de s�quence
                    ws0.write(l-3, c, seq.intitule, self.styleS)                 # Intitul�
                    ws0.write(l-2, c, str(seq.GetDuree()), self.styleD)          # Dur�e
                    ws0.write(l-1, c, Formula(n + '("%s";"%s")' %(seq.nomFichier, os.path.split(seq.nomFichier)[1])), self.h_style)           # Fichier
                    c += 1
                if len(lst) > 0:
                    ws0.write_merge(l-5, l-5, c-i-1, c-1, seq.position+1, self.styleN)
                ws0.col(c).width = 10*20
                c += 1
            return  ws0, dicLigne, last 
        
        col_deb = 6
        
        
        
        if comp_sav == 'C':
            nom = self.referentiel.nomCompetences + u" - CI"
        else:
            nom = self.referentiel.nomSavoirs + u" - CI"
        ws0 = wb.add_sheet(nom)
        ws0.write_merge(1, 3, 1, col_deb-2, self.referentiel.Enseignement[0], self.styleT)
        ws0.write_merge(5, 5, 0, col_deb-2, nom, self.styleE)
        if comp_sav == 'C':
            dicLigne, lastComp = self.traiterDic(self.referentiel.dicCompetences.items(), ws0, c = 1)
        else:
            dicLigne, lastComp = self.traiterDic(self.referentiel.dicSavoirs.items(), ws0, c = 1)
        
        #
        # S�quences en colonne
        #  
        c = col_deb
        l = lastComp+1
        
        ws0.write_merge(l, l, 0, col_deb-2, u"Séquences", self.styleE)
        
        l += 1
        ws0.write(l, c-6, u"Année", self.styleS)             # Ann�e
        ws0.write(l, c-5, u"Période", self.styleS)           # P�riode
        ws0.write(l, c-4, u"Numéro", self.styleS)            # Num�ro de s�quence
        ws0.write(l, c-3, u"Fichier", self.styleS)           # Fichier
        ws0.write(l, c-2, u"Intitulé", self.styleS)          # Intitul�
        
        ws0.write(l, c, u"Durée", self.styleS)               # Dur�e
        
        l += 2
        pt = 0
        n = "HYPERLINK"
        for p, lst in enumerate(self.seqTriees):
            for i, seq in enumerate(lst):
                ws0.write(l, c-4, (i*(p+1))+1, self.styleN)                  # Num�ro de s�quence
                ws0.write(l, c-2, seq.intitule, self.styleS)                 # Intitul�
                ws0.write(l, c, getHoraireTxt(seq.GetDuree()), self.styleD)            # Dur�e
                ws0.write(l, c-3, Formula(n + '("%s";"%s")' %(seq.nomFichier, os.path.split(seq.nomFichier)[1])), self.h_style)           # Fichier
                l += 1
            if len(lst) > 0:
                ws0.write_merge(l-i-1, l-1, c-5, c-5, seq.position+1, self.stylePer)  # P�riode
                if pt == 0 and seq.position == 4:
                    pt = l
            ws0.row(c).height = 10*20
            l += 1
        lastSeq = l-1
        if pt != 0:
            ws0.write_merge(lastComp+4, pt-1, c-6, c-6, u"1re", self.styleN)  # P�riode
            if pt < lastSeq:
                ws0.write_merge(pt+1, lastSeq, c-6, c-6, u"Tale", self.styleN)  # P�riode
                
                
        #
        # CI en ligne
        #        
        c = col_deb+2
        l = 4
        ws0.write_merge(l-3, l-3, c, c+len(self.referentiel.CentresInterets)-1, 
                        constantes.getSingulierPluriel(self.referentiel.nomCI, True), self.styleT)
        
        for i, ci in enumerate(self.referentiel.CentresInterets):
            if len(self.referentiel.positions_CI) > i:
                pos = self.referentiel.positions_CI[i]
                ws0.write(l-2, c, pos, self.stylePosCI)                 # position cible
            ws0.write(l-1, c, self.referentiel.abrevCI+str(i+1), self.styleNumCI)          # num�ro CI
            ws0.write(l, c, ci, self.styleCI)                           # ci
            ws0.col(c).width = 120*20
            c += 1
        lastCI = c
        
        #
        # On additionne les poids horaires par s�quence
        #
        lstSomPoidsSeq = []
        for p, lst in enumerate(self.seqTriees):
            for i, seq in enumerate(lst):
#                if seq.CI.poids != []:
                lstSomPoidsSeq.append(0)
                for poids in seq.CI.poids:
                    lstSomPoidsSeq[-1] += poids
#        print  lstSomPoidsSeq

        #
        # On met les dur�es au croisement seq/CI
        #
        c = col_deb+2
        l = lastComp+4
        j = 0
        sumCI = [0]*len(self.referentiel.CentresInterets)
        for p, lst in enumerate(self.seqTriees):
            for i, seq in enumerate(lst):
                # Poids horaires
                for ci, poids in zip(seq.CI.numCI, seq.CI.poids):
                    p = poids*seq.GetDuree()/lstSomPoidsSeq[j]
                    sumCI[ci] += p
                    ws0.write(l, c+ci, getHoraireTxt(p), self.styleX)
                # Case vides
                for co in range(c, lastCI):
                    try:
                        ws0.write(l, co, "", self.stylenX)
                    except:
                        pass
                l += 1
                j += 1
            l += 1
        
        
        def CodeDansListe(code, dic):
#            print code
#            print dic
            if comp_sav == 'C':
                if code in dic.keys():
                    return True, dic[code][0]
                else:
                    return False, None
            else:
                if code[1:] in dic.keys() and (code[0] == 'S' or code[0] == 'B'):
                    return True , dic[code[1:]][0]
                else:
                    return False, None
        
        
        #
        # On met les croix Comp�tences/CI ou Savoirs/CO
        #
        c = col_deb+2
        l = 6
        for p, lst in enumerate(self.seqTriees):
            for i, seq in enumerate(lst):
                if comp_sav == 'C':
                    lstc = seq.obj["C"].competences
                else:
                    lstc = seq.obj["S"].savoirs
                for cod in lstc:
                    x, li = CodeDansListe(cod, dicLigne)
                    if x:
                        for ci in seq.CI.numCI:
                            try:
                                ws0.write(li, c+ci, "X", self.styleX)
                            except:
                                pass
                                
        for co in range(c, lastCI):
#            for li in range(l, lastComp):
            for li, ch, niv  in dicLigne.values():
                try:
                    ws0.write(li, co, "", self.getStyleComp(niv))
                except:
                    pass
                
        #
        # On met les formules de somme
        #
        l = lastComp+2
        c = col_deb+2
        for ci in range(c, lastCI):
#            cc = constantes.lettreCol(ci)
#            ws0.write(l, ci, Formula("SUM("+cc+str(l+3)+":"+cc+str(lastSeq+1)+")"), self.styleN)
            ws0.write(l, ci, getHoraireTxt(sumCI[ci-c]), self.styleN)
        
        ws0.col(0).width = 80*20
        ws0.col(1).width = 70*20
        self.reglerLargeursGauche(ws0, c, c0 = 1)
        ws0.col(col_deb).width = 80*20
        ws0.col(col_deb+1).width = 15*20
        
        
        return
        
    
    ######################################################################################  
    def genererBilan(self, nomFichier):
#        print "genererBilan"
        wb = Workbook()
        self.definirStyles(wb)
        
        #
        # On trie les s�quences par p�riode
        #
        self.seqTriees = [[], [], [], [], [], [], [], [], [], []]
        for i in range(self.listeSeq.listeSeq.GetItemCount()):
            seq = self.listeSeq.GetSequence(i)
            self.seqTriees[seq.position].append(seq)
        
        
        def traiterSeq(nom, dic):
            ws0 = wb.add_sheet(nom)
            ws0.write_merge(1, 3, 1, 3, self.referentiel.Enseignement[0], self.styleT)
            ws0.write_merge(5, 5, 0, 3, nom, self.styleE)
            dicLigne, last = self.traiterDic(dic.items(), ws0)
          
            #
            # S�quences en ligne
            #
            c = 5
            l = 6
            pt = 0
            n = "HYPERLINK"
            for p, lst in enumerate(self.seqTriees):
                for i, seq in enumerate(lst):
                    ws0.write(l-4, c, (i*(p+1))+1, self.styleN)                  # Num�ro de s�quence
                    ws0.write(l-3, c, seq.intitule, self.styleS)                 # Intitul�
                    ws0.write(l-2, c, getHoraireTxt(seq.GetDuree()), self.styleD)          # Dur�e
                    ws0.write(l-1, c, Formula(n + '("%s";"%s")' %(seq.nomFichier, os.path.split(seq.nomFichier)[1])), self.h_style)           # Fichier
                    ws0.col(c).width = 100*20
                    c += 1
                if len(lst) > 0:
                    ws0.write_merge(l-5, l-5, c-i-1, c-1, seq.position+1, self.styleN) # P�riode
                    if pt == 0 and seq.position == 4:
                        pt = c
                ws0.col(c).width = 15*20
                c += 1
            
            if pt != 0:
                ws0.write_merge(l-6, l-6, 5, pt-1, u"1re", self.styleN)  # Ann�e
                if pt < c:
                    ws0.write_merge(l-6, l-6, pt+1, c, u"Tale", self.styleN)  # Ann�e
                
                
            return  ws0, dicLigne, last 
        
        
        #
        # Feuille Comp�tences
        #
        ws0, dicLigne, last = traiterSeq(self.referentiel.nomCompetences, 
                                         self.referentiel.dicCompetences)
        
        c = 5
        l = 6
        for p, lst in enumerate(self.seqTriees):
            for i, seq in enumerate(lst):         
                for sav in seq.obj["C"].competences:
                    if sav in dicLigne.keys():
                        li = dicLigne[sav][0]
#                        for li in dicLigne[sav]:
                        ws0.write(li, c, "X", self.styleX)
#                for li in range(l+1, last):
                for li, ch, niv in dicLigne.values():
                    try:
                        ws0.write(li, c, "", self.getStyleComp(niv))
                    except:
                        pass
                c += 1
            c += 1
                
        self.reglerLargeursGauche(ws0, c)
        
        
        
        #
        # Feuille Savoirs
        #
        if self.referentiel.tr_com != []: # C'est un enseignement de sp�cialit� avec un tronc commun
            ws0, dicLigne, last = traiterSeq(self.referentiel.nomSavoirs + " " + REFERENTIELS[self.referentiel.tr_com[0]].Enseignement[0], 
                                              REFERENTIELS[self.referentiel.tr_com[0]].dicSavoirs)
            c = 5
            for p, lst in enumerate(self.seqTriees):
                for i, seq in enumerate(lst):
                    for sav in seq.obj["S"].savoirs:
                        if sav[1:] in dicLigne.keys() and sav[0] == 'B':
                            li = dicLigne[sav[1:]][0]
                            ws0.write(li, c, "X", self.styleX)       # X
                    
#                    for li in range(l+1, last):
                    for li, ch, niv in dicLigne.values():
                        try:
                            ws0.write(li, c, "", self.getStyleComp(niv))
                        except:
                            pass
                    c += 1
                c += 1
            self.reglerLargeursGauche(ws0, c)
            
            
        
        
        #
        # Feuille Savoirs
        # 
        ws0, dicLigne, last = traiterSeq(self.referentiel.nomSavoirs + " " + self.referentiel.Enseignement[0], 
                                         self.referentiel.dicSavoirs)
#        print "dicLigne", dicLigne
        c = 5
#        print dicLigne.keys()
        for p, lst in enumerate(self.seqTriees):        # Périodes
            for i, seq in enumerate(lst):               # Séquences
                for sav in seq.obj["S"].savoirs:        
#                    print "  ", sav
                    if sav[1:] in dicLigne.keys() and (sav[0] == 'S' or (sav[0] == 'B' and self.referentiel.tr_com == [])):
                        li = dicLigne[sav[1:]][0]
#                        for li in dicLigne[sav[1:]]:
                        ws0.write(li, c, "X", self.styleX)       # X
                
#                for li in range(l+1, last):
                for li, ch, niv in dicLigne.values():
                    try:
                        ws0.write(li, c, "", self.getStyleComp(niv))
                    except:
                        pass
                c += 1
            c += 1
        self.reglerLargeursGauche(ws0, c)
            
        self.genererBilanComplexe(wb, 'C')
        self.genererBilanComplexe(wb, 'S')
        
#        ws0 = wb.add_sheet(self.referentiel.nomSavoirs)
#        
#        ws0.write_merge(1, 3, 1, 3, self.referentiel.Enseignement[0], styleT)
#        ws0.write_merge(5, 5, 0, 3, self.referentiel.nomSavoirs, styleE)
#        dicLigne, last = traiter(self.referentiel.dicSavoirs.items())
#      
#        #
#        # On met les croix
#        #
#        c = 5
#        l = 5
#        n = "HYPERLINK"
#        for i, seq in enumerate(listePrem + listeTerm):
#            ws0.write(l-4, c, i+1, styleN)                          # Num�ro de s�quence
#            ws0.write(l-3, c, seq.intitule, styleS)                 # Intitul�
#            ws0.write(l-2, c, str(seq.GetDuree()), styleD)          # Dur�e
#            ws0.write(l-1, c, Formula(n + '("%s";"%s")' %(seq.nomFichier,seq.nomFichier)), h_style)           # Fichier
#            
#            for sav in seq.obj["S"].savoirs:
#                if sav[1:] in dicLigne.keys() and sav[0] == 'B':
#                    for li in dicLigne[sav[1:]]:
#                        ws0.write(li, c, "X", styleX)       # X
#            
#            for li in range(l+1, last):
#                try:
#                    ws0.write(li, c, "", stylenX)
#                except:
#                    pass
#            c += 1
#            
#        #
#        # Largeur des colonnes
#        #
#        ws0.col(0).width = 100*20
#        ws0.col(1).width = 100*20
#        ws0.col(2).width = 100*20
#        ws0.col(3).width = 1000*20
#        ws0.col(4).width = 50*20
#        for cc in range(5, c):
#            ws0.col(cc).width = 80*20
#        
#        #
#        # Feuille Comp�tences
#        #        
#        ws0 = wb.add_sheet(self.referentiel.nomCompetences)
#        
#        ws0.write_merge(1, 3, 1, 3, self.referentiel.Enseignement[0], styleT)
#        ws0.write_merge(5, 5, 0, 3, self.referentiel.nomCompetences, styleE)
#        dicLigne, last = traiter(self.referentiel.dicCompetences.items())
#      
#        #
#        # On met les croix
#        #
#        c = 5
#        l = 5
#        for i, seq in enumerate(listePrem + listeTerm):
#            ws0.write(l-4, c, i+1, styleN)                          # Num�ro de s�quence
#            ws0.write(l-3, c, seq.intitule, styleS)                 # Intitul�
#            ws0.write(l-2, c, str(seq.GetDuree()), styleD)          # Dur�e
#            ws0.write(l-1, c, Formula(n + '("%s";"%s")' %(seq.nomFichier,seq.nomFichier)), h_style)           # Fichier
#            
#            for sav in seq.obj["C"].competences:
#                if sav in dicLigne.keys():
#                    for li in dicLigne[sav]:
#                        ws0.write(li, c, "X", styleX)
#            for li in range(l+1, last):
#                try:
#                    ws0.write(li, c, "", stylenX)
#                except:
#                    pass
#            c += 1
#            
#        #
#        # Largeur des colonnes
#        #
#        ws0.col(0).width = 100*20
#        ws0.col(1).width = 100*20
#        ws0.col(2).width = 100*20
#        ws0.col(3).width = 1000*20
#        ws0.col(4).width = 50*20
#        for cc in range(5, c):
#            ws0.col(cc).width = 80*20
        
        #
        # Sauvegarde
        #
        try:
            wb.save(nomFichier)
        except IOError:
            messageErreur(self, u'Accés refusé',
                          u"Le fichier n'a pas pu être enregistré !\n\n" \
                          u"Il est peut-être déja ouvert ...")
            return False
        
        return True
        
#    ######################################################################################  
#    def enregistrer(self, nomFichier):
#        fichierPP = constantes.fichierProgressionProgramme[self.typeEnseignement]
#        try:
#            tableau = grilles.PyExcel(os.path.join(TABLE_PATH, fichierPP))
#        except:
#            print fichierPP, "est d�ja ouvert !"
#            return False
#
##        def ecrire(feuille, l, c):
##            v = tableau.getCell(feuille, l, c)
##            if v == "A":
##                v = "B"
##            elif v == "B":
##                v = "C"
##            elif v == "C":
##                v = "C"
##            else:
##                v = "A"
##            tableau.setCell(feuille, l, cc+i, v)
#            
#        feuilleP = u"Progression - Programme"
#        feuilleS = u"Progression - Systèmes"
#        
#        # Première cellule "séquence"
#        lcp, ccp = (4, 11) # K4
#        lct, cct = (4, 15) # O4
#        
#        # Première cellule "durée"
#        ldp, cdp = (5, 11) # K5
#        
#        # Première cellule "systèmes"
#        lsp, csp = (8, 4) # D8
#        lst, cst = (8, 8) # H8
#        
#        listePrem = []
#        listeTerm = []
#
#        
#        for i in range(self.listeSeq.listeSeq.GetItemCount()):
#            seq = self.listeSeq.GetSequence(i)
#            if seq.position >= 4:
#                listeTerm.append(seq)
#            else:
#                listePrem.append(seq)
#        
##        listePrem = sorted(listePrem, key=lambda s: s.intitule)
##        listeTerm = sorted(listeTerm, key=lambda s: s.intitule)
#        
#        listeSystemes = []
#        for i, seq in enumerate(listePrem + listeTerm):
#            
#            if seq in listePrem:
#                c = ccp + i
#                cs = csp + i
#            else:
#                c = cct + i - len(listePrem)
#                cs = cst + i - len(listePrem)
#                
#            tableau.setCell(feuilleP, lcp, c, seq.intitule)
#            tableau.setCell(feuilleP, ldp, c, str(seq.GetDuree()))
#            tableau.setLink(feuilleP, lcp, c, seq.nomFichier)
#            tableau.setCell(feuilleP, lcp-1, c, i+1)
#            
#            dicCel = constantes.dicCellSavoirs[self.typeEnseignement]
#            for sav in seq.obj["S"].savoirs:
#                if sav in dicCel.keys():
#                    lig0, lig1 = dicCel[sav]
#                    for l in range(lig0, lig1+1):
#                        tableau.setCell(feuilleP, l, c, "X")
#                else:
#                    continuer = True
#                    s=1
#                    while continuer:
#                        sav1 = sav+'.'+str(s)
#                        if sav1 in dicCel.keys():
#                            lig0, lig1 = dicCel[sav1]
#                            for l in range(lig0, lig1+1):
#                                tableau.setCell(feuilleP, l, c, "X")
#                            s += 1
#                        else:
#                            continuer = False
#            
#            #
#            # Tableau "Systèmes"
#            #
#            nbrSystemes = seq.GetNbrSystemes()
#            for syst in seq.systemes:
#                # ligne du tableau correspondant au syst�me
#                if not syst.nom in listeSystemes:
#                    listeSystemes.append(syst.nom)
#                    l = lsp+len(listeSystemes)-1
#                    tableau.setCell(feuilleS, l, 1, syst.nom)
#                else:
#                    l = lsp+listeSystemes.index(syst.nom)-1
#                
#                # nombre d'exemplaires du système utilisés dans la séquence
#                if syst.nom in nbrSystemes.keys():
#                    tableau.setCell(feuilleS, l, cs, nbrSystemes[syst.nom])
#                
#            #
#            # Ajout éventuel de colonnes
#            #
#            if seq in listePrem:
#                if len(listePrem) > 3 and i < len(listePrem) - 3:
#                    tableau.insertPasteCol(feuilleP, c+1)
#                    tableau.insertPasteCol(feuilleS, cs+1)
#                    cct += 1
#            else:
#                if len(listeTerm) > 3 and i - len(listePrem) < len(listeTerm) - 3:
#                    tableau.insertPasteCol(feuilleP, c+1)
#                    tableau.insertPasteCol(feuilleS, c+1)
#                
#        try:                   
#            tableau.save(nomFichier)
#        except :
#            print nomFichier, "est d�ja ouvert !"
#            return False
#            
#        tableau.close()
#        return True
        
    ######################################################################################  
    def EvtRadioBox(self, event):
#        print "EvtRadioBox"
        radio_selected = event.GetEventObject()
        typeEnseignement = Referentiel.getEnseignementLabel(radio_selected.GetLabel())[0]
        self.referentiel = REFERENTIELS[typeEnseignement]
        
#        self.referentiel = REFERENTIELS[Referentiel.getEnseignementLabel(self.cb_type.GetItemLabel(event.GetInt()))[0]]
        
#        self.typeEnseignement = Referentiel.getEnseignementLabel(self.cb_type.GetItemLabel(event.GetInt()))[0]
        self.MiseAJourListe()
#        for c, e in constantes.Enseignement.items():
#            if e[0] == self.cb_type.GetItemLabel(event.GetInt()):
#                self.typeEnseignement = c
#                self.MiseAJourListe()
#                break
        
        
    #############################################################################            
    def EvtTextDoss(self, event):
        self.dossiers = self.txtDoss.GetValue().split(";")
        self.VerifierDossiers()
        self.MiseAJourListe()
        
        
    ########################################################################################################
    def OnDossier(self, event):
        dlg = wx.DirDialog(self, u"Choisir un dossier",
                           style = wx.DD_DEFAULT_STYLE|wx.DD_DIR_MUST_EXIST
                           #| wx.DD_CHANGE_DIR
                           )

        if dlg.ShowModal() == wx.ID_OK:
            if len(self.dossiers) == 0:
                self.txtDoss.ChangeValue(dlg.GetPath())
            else:
                self.txtDoss.ChangeValue(self.txtDoss.GetValue()+";"+dlg.GetPath())
            self.dossiers.append(dlg.GetPath())
            self.MiseAJourListe()
            
        dlg.Destroy()
        
        
    ########################################################################################################
    def VerifierDossiers(self):
        col = wx.SystemSettings_GetColour(wx.SYS_COLOUR_WINDOW)
        self.txtDoss.SetBackgroundColour(col)
        self.dossiersOk = True
        for dossier in self.dossiers:
            if not os.path.isdir(dossier):
                self.txtDoss.SetBackgroundColour("pink")
                self.dossiersOk = False
                break
        self.txtDoss.Refresh()
    

    

        
    ########################################################################################################
    def MiseAJourListe(self, event = None):       
        if self.dossiersOk:
            wx.BeginBusyCursor()
            
            l = []
            if self.recurs:
                for dossier in self.dossiers:
                    for root, dirs, files in os.walk(dossier):
#                        print files
                        l.extend([os.path.join(root, f) for f in files if os.path.splitext(f)[1] == '.seq'])
            else:
                for dossier in self.dossiers:
                    l.extend(glob.glob(os.path.join(dossier, "*.seq")))
                

            dlg =    wx.ProgressDialog(u"Recherche des séquences",
                                       u"",
                                       maximum = len(l),
                                       parent=self.Parent,
                                       style = 0
                                        | wx.PD_APP_MODAL
                                        #| wx.PD_CAN_ABORT
                                        #| wx.PD_CAN_SKIP
                                        #| wx.PD_ELAPSED_TIME
                                        | wx.PD_ESTIMATED_TIME
                                        | wx.PD_REMAINING_TIME
                                        | wx.PD_AUTO_HIDE
                                        )
#            print "l =", l
            listSequences = []
            count = 0
            for f in l:
                try:
                    dlg.Update(count, f)
                except:
                    print "Erreur", f
                
                classe, sequence = self.OuvrirFichierSeq(f)
#                print classe.typeEnseignement ,  self.referentiel.Code
                if classe != None and classe.typeEnseignement == self.referentiel.Code:
                    sequence.nomFichier = f
                    listSequences.append(sequence)
                count += 1
            
#            print "listSequences AVANT", listSequences
            listeCI = self.getCIcommuns(listSequences)
            if len(listeCI) > 1:
                dlgc = DialogChoixCI(self, listeCI, self.referentiel)
                dlgc.ShowModal()
                choix = dlgc.choix
                listSequences = listeCI[choix][0]
        
                dlgc.Destroy()
            
#            print "listSequences APRES", listSequences
            self.listeSeq.MiseAJourListe(listSequences)
            self.tb.RemoveTool(20)
#            self.tb.Realize()
            
            dlg.Update(count, u"Terminé")
            dlg.Destroy()
            wx.EndBusyCursor()
        
        
        
    ########################################################################################################
    def OuvrirFichierSeq(self, nomFichier):
        fichier = open(nomFichier,'r')

        classe = pysequence.Classe(self.Parent)
        sequence = pysequence.Sequence(self, classe)
        classe.SetDocument(sequence)

        try:
            root = ET.parse(fichier).getroot()
            rsequence = root.find("Sequence")
            rclasse = root.find("Classe")
            if rclasse is not None:
                classe.setBranche(rclasse)
            if rsequence is not None:
                sequence.setBranche(rsequence)
            else:   # Ancienne version , forc�ment STI2D-ETT !!
                classe.typeEnseignement, self.classe.familleEnseignement = ('ET', 'STI')
                classe.referentiel = REFERENTIELS[classe.typeEnseignement]
                sequence.setBranche(root)
            return classe, sequence
        except:
            print u"Le fichier n'a pas pu être ouvert :",nomFichier
#            messageErreur(self,u"Erreur d'ouverture",
#                          u"La s�quence p�dagogique\n    %s\n n'a pas pu �tre ouverte !" %nomFichier)
#            fichier.close()
#            self.Close()
            return None, None



class PanelListe(wx.Panel, listmix.ColumnSorterMixin):
    def __init__(self, parent, fen):
        wx.Panel.__init__(self, parent, -1, style = wx.BORDER_RAISED)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        self.fen = fen
        
        self.il = wx.ImageList(16, 16)
        self.idx1 = self.il.Add(images.Icone_sequence.GetImage().Rescale(16,16).ConvertToBitmap())
        self.sm_up = self.il.Add(images.SmallUpArrow.GetBitmap())
        self.sm_dn = self.il.Add(images.SmallDnArrow.GetBitmap())
        
        
        self.listeSeq = TestListCtrl(self, -1, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
        self.currentItem = None
        self.listeSeq.SetImageList(self.il, wx.IMAGE_LIST_SMALL)
        sizer.Add(self.listeSeq, 1, wx.EXPAND)
        
        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        
#        self.listeSeq.Bind(wx.EVT_SIZE, self.OnResize)
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected, self.listeSeq)
        self.listeSeq.Bind(wx.EVT_LEFT_DCLICK, self.OnDoubleClick)
#        self.listeSeq.Bind(wx.EVT_RIGHT_DOWN, self.OnRightDown)
        
        # for wxMSW
        self.listeSeq.Bind(wx.EVT_COMMAND_RIGHT_CLICK, self.OnRightClick)

        # for wxGTK
        self.listeSeq.Bind(wx.EVT_RIGHT_UP, self.OnRightClick)
        
    def GetListCtrl(self):
        return self.listeSeq

    def GetSortImages(self):
        return (self.sm_dn, self.sm_up)
    
    ########################################################################################################
    def OnResize(self, event=None):
        height = 30
        for indx in xrange(self.listeSeq.GetItemCount()):
            height += self.listeSeq.GetItemRect(indx).height
        self.listeSeq.SetMinSize((-1, height))
        
    ########################################################################################################
    def OnItemSelected(self, event):
        self.currentItem = event.m_itemIndex
        event.Skip()
    
    ########################################################################################################
    def OnDoubleClick(self, event=None):
        nomFichier = self.GetSequence(self.currentItem).nomFichier
        self.fen.Parent.ouvrir(nomFichier)
        event.Skip()
        
    ########################################################################################################
    def OnRightDown(self, event):
        x = event.GetX()
        y = event.GetY()
        item, flags = self.listeSeq.HitTest((x, y))

        if item != wx.NOT_FOUND and flags & wx.LIST_HITTEST_ONITEM:
            self.listeSeq.Select(item)

        event.Skip()
        
    ########################################################################################################
    def OnRightClick(self, event):
        # only do this part the first time so the events are only bound once
        if not hasattr(self, "popupID4"):
            self.popupID4 = wx.NewId()
            self.popupID5 = wx.NewId()

            self.Bind(wx.EVT_MENU, self.OnPopupFour, id=self.popupID4)
            self.Bind(wx.EVT_MENU, self.OnPopupFive, id=self.popupID5)

        # make a menu
        menu = wx.Menu()
        # add some items
        menu.Append(self.popupID4, u"Supprimer")
        menu.Append(self.popupID5, u"Ouvrir")


        # Popup the menu.  If an item is selected then its handler
        # will be called before PopupMenu returns.
        self.PopupMenu(menu)
        menu.Destroy()

    def OnPopupFour(self, event):
#        item = self.listeSeq.GetItem(self.currentItem)
        self.listeSeq.DeleteItem(self.currentItem)

    def OnPopupFive(self, event):
        self.OnDoubleClick(event)
        
        
    def GetSequence(self, num):
        return self.itemDataMap[self.listeSeq.GetItemData(num)][2]
    
    ########################################################################################################
    def MiseAJourListe(self, listSequences):
        self.listeSeq.ClearAll()
        self.listeSeq.InsertColumn(0, u"Fichier séquence")
        self.listeSeq.InsertColumn(1, u"Dossier")
    
        self.itemDataMap = {}
        for i, seq in enumerate(listSequences):
            f = seq.nomFichier
            d, f = os.path.split(f)
            f = os.path.splitext(f)[0]
            try:
                pos = self.listeSeq.InsertStringItem(i, f)
                self.listeSeq.SetStringItem(pos, 1, d)
                self.itemDataMap[i] = (f, d, seq)
                self.listeSeq.SetItemData(i, i)
            except:
                print "Erreur", f
            
#        self.listeSeq.SetColumnWidth(0, wx.LIST_AUTOSIZE)
#        self.listeSeq.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        self.listeSeq.Layout()
#        self.OnResize()
#        self.Fit()
        
        listmix.ColumnSorterMixin.__init__(self, 2)
        
        
#############################################################################################################
#
# Fenetre de bilan d'objectifs
# 
#############################################################################################################

class TestListCtrl(wx.ListCtrl, listmix.ListCtrlAutoWidthMixin):
    def __init__(self, parent, ID, pos=wx.DefaultPosition,
                 size=wx.DefaultSize, style=0):
        wx.ListCtrl.__init__(self, parent, ID, pos, size, style)
        listmix.ListCtrlAutoWidthMixin.__init__(self)
        
        self.Bind(wx.EVT_LIST_BEGIN_DRAG, self.startDrag)
        self.Bind(wx.EVT_LEFT_UP, self.endDrag)
        self.Bind(wx.EVT_MOTION, self.onMove)
        self.Bind(wx.EVT_LEAVE_WINDOW, self.onLeave)

        self.idx = None

    def startDrag(self, e):

        self.idx = e.GetIndex()
        self.SetCursor(wx.CursorFromImage(constantes.images.Curseur_InsererApres.GetImage()))

    def onLeave(self, event):
        if self.idx != None:
            self.Select(self.idx)
            self.idx = None
        self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        event.Skip()
        
    def onMove(self, event):    
        if self.idx != None:
            x = event.GetX()
            y = event.GetY()
            index, flags = self.HitTest((x, y))
            if index != wx.NOT_FOUND and flags & wx.LIST_HITTEST_ONITEM:
                rect = self.GetItemRect(index)
                if y > rect.y + rect.height/2:
                    index += 1
                
                self.Select(index)
            

    def endDrag(self, event):    
        if self.idx != None:
            x = event.GetX()
            y = event.GetY()
            
            
            index, flags = self.HitTest((x, y))
            if index != wx.NOT_FOUND and flags & wx.LIST_HITTEST_ONITEM:
                rect = self.GetItemRect(index)
                if y > rect.y + rect.height/2:
                    index += 1
                index += 1

                
                f = self.GetItemText(self.idx, 0)
                d = self.GetItemText(self.idx, 1)
                i = self.GetItemData(self.idx)

                pos = self.InsertStringItem(index, f)

                self.SetStringItem(pos, 1, d)
                self.SetItemData(index, i)
    #            self.itemDataMap[index] = self.itemDataMap[self.idx]
                
                if index > self.idx:
                    self.DeleteItem(self.idx)
    #                del self.itemDataMap[self.idx]
                else:
                    self.DeleteItem(self.idx+1)
    #                del self.itemDataMap[self.idx+1]
            
            self.idx = None
            self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        
    def GetHeight2(self):
        height = int(2*16)
        for indx in xrange(self.GetItemCount()):
            height += self.GetItemRect(indx).height
        return height
    
    def GetWidth2(self):
        return self.GetViewRect()[2]
    
    
class DialogChoixCI(wx.Dialog):
    def __init__(self, parent, listeCI, referentiel, style=wx.DEFAULT_DIALOG_STYLE):

        # Instead of calling wx.Dialog.__init__ we precreate the dialog
        # so we can set an extra style that must be set before
        # creation, and then we create the GUI object using the Create
        # method.
        pre = wx.PreDialog()
        pre.SetExtraStyle(wx.DIALOG_EX_CONTEXTHELP)
        pre.Create(parent, -1, constantes.getSingulierPluriel(referentiel.nomCI, True) + u" différents")

        # This next step is the most important, it turns this Python
        # object into the real wrapper of the dialog (instead of pre)
        # as far as the wxPython extension is concerned.
        self.PostCreate(pre)

        # Now continue with the normal construction of the dialog
        # contents
        sizer = wx.BoxSizer(wx.VERTICAL)

        label = wx.StaticText(self, -1, u"Les séquences identifiées n'utilisent pas toutes les mêmes centres d'intérêt.\n" \
                              u"Merci de choisir la liste de centres d'intérêt à utiliser pour la synthèse.\n")
        label.SetHelpText("Choisir la liste de centres d'intérêt à utiliser pour la synthèse    ²")
        sizer.Add(label, 0, wx.ALIGN_CENTRE|wx.ALL, 5)

        box = wx.BoxSizer(wx.HORIZONTAL)
        
        maxs = 600/len(listeCI)
        self.radio = []
        for i, (l, s) in enumerate(listeCI):
            lstsz = wx.BoxSizer(wx.VERTICAL)
            if i == 0:
                radio = wx.RadioButton(self, -1, constantes.getSingulierPluriel(referentiel.nomCI, True) + u" #"+str(i+1), style = wx.RB_GROUP )
            else:
                radio = wx.RadioButton(self, -1, constantes.getSingulierPluriel(referentiel.nomCI, True) +u" #"+str(i+1))
            radio.SetHelpText(u"Liste de "+referentiel.abrevCI+" #"+str(i))
            self.Bind(wx.EVT_RADIOBUTTON, self.OnSelect, radio )
            self.radio.append(radio)
            lstsz.Add(radio, 0, wx.ALIGN_CENTRE|wx.ALL|wx.EXPAND, 5)
            
            label = wx.StaticText(self, -1, str(len(l))+u" séquence"+(len(l)>1)*"s")
            lstsz.Add(label, 0, wx.ALIGN_CENTRE|wx.ALL|wx.EXPAND, 5)                      
                                  
            text = wx.TextCtrl(self, -1, u"\n".join(s), size=(maxs,-1), style = wx.TE_READONLY|wx.TE_MULTILINE)
            text.SetHelpText(u"Liste de "+referentiel.abrevCI+" #"+str(i))
            lstsz.Add(text, 1, wx.ALIGN_CENTRE|wx.ALL|wx.EXPAND, 5)
            box.Add(lstsz, 0, wx.ALIGN_CENTRE|wx.ALL|wx.EXPAND, 5 )
            
        sizer.Add(box, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)

        line = wx.StaticLine(self, -1, size=(20,-1), style=wx.LI_HORIZONTAL)
        sizer.Add(line, 0, wx.GROW|wx.ALIGN_CENTER_VERTICAL|wx.RIGHT|wx.TOP, 5)

        btnsizer = wx.StdDialogButtonSizer()
        
        btn = wx.Button(self, wx.ID_OK)
        btn.SetHelpText(u"Valider")
        btn.SetDefault()
        btnsizer.AddButton(btn)

        btnsizer.Realize()

        sizer.Add(btnsizer, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5)

        self.choix = 0
        
        self.SetSizer(sizer)
        sizer.Fit(self)              
    
    
    def OnSelect( self, event ):
        self.choix = self.radio.index(event.GetEventObject())
        
