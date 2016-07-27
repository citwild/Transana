# Copyright (C) 2002-2016 Spurgeon Woods LLC

#This program is free software; you can redistribute it and/or
#modify it under the terms of the GNU General Public License
#as published by the Free Software Foundation; either version 2
#of the License, or (at your option) any later version.

#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.

#You should have received a copy of the GNU General Public License
#along with this program; if not, write to the Free Software
#Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

""" This module implements the Synonym Editor for teh Word Frequency Reports for Transana """

__author__ = "David K. Woods <dwoods@wcer.wisc.edu>"

# Import Python's os module
import os

# Import wxPython
import wx

# Import Transana's Globals
import TransanaGlobal

class SynonymEditor(wx.Dialog):
    """ This is the main form for the Synonym Editor """

    def __init__(self, parent, dataItem):

        # Create the basic Frame structure with a white background
        wx.Dialog.__init__(self, parent, -1, _("Word Group Editor"), size=wx.Size(600, 600), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundColour(wx.WHITE)
        
        # Set the report's icon
        transanaIcon = wx.Icon(os.path.join(TransanaGlobal.programDir, "images", "Transana.ico"), wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)

        # Create main form sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create main Panel for the form
        panel = wx.Panel(self, -1)
        # Define a sizer for the Panel
        pnlSizer = wx.FlexGridSizer(rows=3, cols=2, hgap=0, vgap=0)

        # Word Group
        txt1 = wx.StaticText(panel, -1, "Word Group:", style=wx.ALIGN_RIGHT)
        self.synonymGroup = wx.TextCtrl(panel, -1)

        # Words
        txt2 = wx.StaticText(panel, -1, "Words:", style=wx.ALIGN_RIGHT)
        self.synonyms = wx.ListBox(panel, -1)

        # Delete Button
        self.deleteBtn = wx.Button(panel, -1, _("Delete Selected Word"))
        self.deleteBtn.Bind(wx.EVT_BUTTON, self.OnDelete)

        # Places form items into the FlexGridSizer on the Panel
        pnlSizer.AddMany([(txt1, 0, wx.ALL, 10), (self.synonymGroup, 1, wx.EXPAND | wx.TOP | wx.BOTTOM | wx.RIGHT, 10),
                          (txt2, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10), (self.synonyms, 1, wx.EXPAND | wx.BOTTOM | wx.RIGHT, 10),
                          ((1, 1), 0), (self.deleteBtn, 1, wx.ALIGN_RIGHT | wx.RIGHT, 10)])
        # Specify that Row 1 can grow
        pnlSizer.AddGrowableRow(1, 1)
        # Specify that Column 1 can grow
        pnlSizer.AddGrowableCol(1, 1)
        # Set the Sizer on the Panel
        panel.SetSizer(pnlSizer)
        # Add the Panel to the main Sizer
        mainSizer.Add(panel, 1, wx.EXPAND)

        # Add a Horizontal Sizer for the form buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer on the left to expand, allowing the buttons to be right-justified
        btnSizer.Add((10, 1), 1)
        # Add an OK button, a Cancel button, and a Help button to the Form
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        btnOK.Bind(wx.EVT_BUTTON, self.OnOK)
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        btnCancel.Bind(wx.EVT_BUTTON, self.OnOK)
        btnHelp = wx.Button(self, -1, _("Help"))

        # Put the buttons in the  Button sizer
        btnSizer.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        btnSizer.Add(btnCancel, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        btnSizer.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        # Add the Button Size to the Dialog Sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)

        # Set the form's Main Sizer and enable auto-layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)

        # Set the Word Group and Words list data into the controls
        self.synonymGroup.SetValue(dataItem[0])
        self.synonyms.Set(dataItem[2])
        # If this is the "Do Not Show" group ...
        if dataItem[0] == 'Do Not Show Group':
            # ... disable this field, as we do not allow editing of this group name
            self.synonymGroup.Enable(False)

    def OnOK(self, event):
        """ Handle the OK button """
        # End the modal display of the form
        self.EndModal(event.GetId())

    def OnDelete(self, event):
        """ Handle Delete Button press """
        # If there is a selection ...
        if self.synonyms.GetSelection() != wx.NOT_FOUND:
            # ... delete the Word from the Control
            self.synonyms.Delete(self.synonyms.GetSelection())

    def GetSynonymValues(self):
        """ Return the values left in the control """
        # Create an empty list
        result = []
        # Iterate through the Word List ...
        for item in range(self.synonyms.GetCount()):
            # ... and add the contents to the Results list
            result.append(self.synonyms.GetString(item))
        # Return the Results List
        return result
