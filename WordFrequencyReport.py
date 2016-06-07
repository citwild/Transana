#Copyright (C) 2002-2016  The Board of Regents of the University of Wisconsin System

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

""" This module implements the Word Frequency Reports for Transana """

__author__ = "David K. Woods <dwoods@wcer.wisc.edu>"

# import Python's os and sys modules
import os, sys
# import Python's Regular Expression module
import re

# Import wxPython
import wx
import wx.lib.mixins.listctrl as ListCtrlMixins


# If running stand-alone ...
if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation


# Import Transana's Clip Object
import Clip
# Import Transana's Database Interface
import DBInterface
# Import Transana's Dialog Boxes
import Dialogs
# Import Transana's Document object
import Document
# import Transana's Search module
import ProcessSearch
# Import Transana's Quote object
import Quote
# Import Transana's Synonym Editor
import SynonymEditor
# Import Transana's Text Report infrastructure
import TextReport
# Import Transana's Constants
import TransanaConstants
# Import Transana's Globals
import TransanaGlobal
# Import Transana's Images
import TransanaImages
# Import Transana's Transcript Object
import Transcript


class CheckListCtrl(wx.ListCtrl, ListCtrlMixins.CheckListCtrlMixin):
    """ Create a wxListCtrl class with the CheckListCtrlMixin applied """
    def __init__(self, parent):
        """ Create a wxListCtrl with the CheckListCtrlMixin """
        # Initialize a variable for the parent FORM.  (a Panel is passed in here, but that's not useful!)
        self.parentForm = None
        # Initialize a ListCtrl
        wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_REPORT | wx.LC_SORT_ASCENDING)
        # Add the CheckListCtrlMixin
        ListCtrlMixins.CheckListCtrlMixin.__init__(self)
        # Set EVT_LIST_ITEM_ACTIVATED handler (Double-click checks/unchecks item)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnItemActivated)
        # Set EVT_LIST_COL_CLICK handler to track what column and what direction of sort we have
        self.Bind(wx.EVT_LIST_COL_CLICK, self.OnColumnClick)
        # Initialize the SynonymGroupCtrl (which this control may need to be able to change) to None
        self.synonymGroupCtrl = None

    # This is called by the base class when an item is checked/unchecked
    def OnCheckItem(self, index, flag):
        """ Handle check / uncheck of list items """
        # If a SynonymGroupCtrl is defined and blank ...
        if (self.synonymGroupCtrl != None) and (self.synonymGroupCtrl.GetValue() == ''):
            # ... and the item just clicked is becoming Checked ...
            if self.IsChecked(index):
                # ... add the text of the checked item to the SynonymGroupCtrl
                self.synonymGroupCtrl.SetValue(self.GetItem(itemId=index, col=0).GetText())

        # We need to know if there are ANY checked items in the list
        count = 0
        # Iterate through the list
        for itemId in range(self.GetItemCount()):
            # If we find a checked item ...
            if self.IsChecked(itemId):
                # ... increment our counter
                count += 1
                # ... and stop looking
                break
        # if we found NO checked items ...
        if count == 0:
            # ... reset the SynonymGroupCtrl to blank
            self.synonymGroupCtrl.SetValue('')

    def OnItemActivated(self, evt):
        """ Handle item double-click """
        # Toggle the checked state of the selected item
        self.ToggleItem(evt.m_itemIndex)

    def OnColumnClick(self, event):
        """ Handle setting the sort by clicking the column header """
        # Determine which column has been selected
        col = event.GetColumn()
        # Column image is a proxy for sort direction!!  I don't seem to be able to get that information directly.
        # NOTE that this information is set AFTER the sort has occurred, so the secondary sort need to take that
        #      into account.
        # If the parent Form has been defined ...
        if self.parentForm != None:
            # ... inform the parent form what column is sorted in which direction
            self.parentForm.sortColumn = col
            # GetImage() == 2 is ascending order, GetImage() == 3 is descending order (!)
            self.parentForm.sortAscending = (self.GetColumn(col).GetImage() == 2)

    def SetParentForm(self, form):
        """ Allow the control to know about the parent form, allowing us to track Column and Direction of sort """
        # Register the parent form.  This form must have sortColumn and sortAscending variables to accept data
        self.parentForm = form
        
    def SetSynonymGroupCtrl(self, ctrl):
        """ Allow definition of a Control which should be labelled with the first clicked item in this control """
        # Note the control passed in for naming Synonym Groups
        self.synonymGroupCtrl = ctrl



class AutoWidthListCtrl(wx.ListCtrl, ListCtrlMixins.ListCtrlAutoWidthMixin):
    """ Create a wxListCtrl class with the ListCtrlAutoWidthMixin applied """
    def __init__(self, parent, ID=-1, style=0):
        # Initialize a ListCtrl
        wx.ListCtrl.__init__(self, parent, ID, style=style)
        # Add the ListCtrlAutoWidthMixin
        ListCtrlMixins.ListCtrlAutoWidthMixin.__init__(self)


class WordFrequencyReport(wx.Frame, ListCtrlMixins.ColumnSorterMixin):
    """ This is the main Window for the Word Frequency Reports for Transana """

    def __init__(self, parent, tree, startNode):
        # Remember the parent
        self.parent = parent
        # Note the Control Object from the parent
        self.ControlObject = self.parent.ControlObject
        # Remember the tree that is passed in
        self.tree = tree
        # Remember the start node passed in
        self.startNode = startNode
        # We need a specifically-named dictionary in a particular format for the ColumnSorterMixin.
        # Initialize that here.  (This also tells us if we've gotten the data yet!)
        self.itemDataMap = {}
        # We need a flag that indicates the need to repoputate the itemDataMap because it is out of date.
        self.needsUpdate = True
        # We need a flag that indicates if the report's itemDataMap is CURRENTLY being updated
        self.isUpdating = False
        # We also need a synonyms list.  Initialize it here
        self.synonyms = {}
        # Let's keep track of what column and what direction our current sort is.  That way, we
        # can recreate it as we manipulate the list.  Start with a Descending sort of the Count
        self.sortColumn = 1
        self.sortAscending = False
        # Get the global print data
        self.printData = TransanaGlobal.printData

        # Determine the screen size for setting the initial dialog size
        if __name__ == '__main__':
            rect = wx.Display(0).GetClientArea()  # wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
        else:
            rect = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
        # The width and height of the form should be 80% of the full screen
        width = min(int(rect[2] * .80), 800)
        height = rect[3] * .80

        # Build the Report Title
        self.title = unicode(_('Word Frequency Report'), 'utf8') + u' '
        # Get the Item Name and Item Data from the tree node passed in.
        itemName = tree.GetItemText(startNode)
        itemData = tree.GetPyData(startNode)

        # Build the Title based on the node type passed in
        if itemData.nodetype in ['LibraryRootNode']:
            self.title += unicode(_('for all Libraries'), 'utf8')
        elif itemData.nodetype in ['LibraryNode', 'SearchLibraryNode']:
            self.title += unicode(_('for Library "%s"'), 'utf8') % itemName
        elif itemData.nodetype in ['DocumentNode', 'SearchDocumentNode']:
            self.title +=unicode( _('for Document "%s"'), 'utf8') % itemName
        elif itemData.nodetype in ['EpisodeNode', 'SearchEpisodeNode']:
            self.title += unicode(_('for Episode "%s"'), 'utf8') % itemName
        elif itemData.nodetype in ['TranscriptNode', 'SearchTranscriptNode']:
            self.title += unicode(_('for Transcript "%s"'), 'utf8') % itemName
        elif itemData.nodetype in ['CollectionsRootNode']:
            self.title += unicode(_('for all Collections'), 'utf8')
        elif itemData.nodetype in ['CollectionNode', 'SearchCollectionNode']:
            self.title += unicode(_('for Collection "%s"'), 'utf8') % itemName

        # Create the basic Frame structure with a white background
        wx.Frame.__init__(self, parent, -1, self.title, size=wx.Size(width, height), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundColour(wx.WHITE)
        
        # Set the report's icon
        transanaIcon = wx.Icon(os.path.join(TransanaGlobal.programDir, "images", "Transana.ico"), wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)

        # Create a Toolbar for the Form
        self.toolbar = self.CreateToolBar()
        self.toolbar.SetToolBitmapSize((24, 24))

        # Add a Check All button
        self.checkAll = wx.BitmapButton(self.toolbar, -1, TransanaImages.Check.GetBitmap(), size=(24, 24))
        self.checkAll.SetToolTipString(_("Check All Selected"))
        self.toolbar.AddControl(self.checkAll)
        self.checkAll.Bind(wx.EVT_BUTTON, self.OnCheck)

        # Add a Uncheck All button
        self.uncheckAll = wx.BitmapButton(self.toolbar, -1, TransanaImages.NoCheck.GetBitmap(), size=(24, 24))
        self.uncheckAll.SetToolTipString(_("Uncheck All Selected"))
        self.toolbar.AddControl(self.uncheckAll)
        self.uncheckAll.Bind(wx.EVT_BUTTON, self.OnCheck)

        # Add a separator
        self.toolbar.AddSeparator()

        # Add a Search button
        self.search = wx.BitmapButton(self.toolbar, -1, TransanaImages.Search16.GetBitmap(), size=(24, 24))
        self.search.SetToolTipString(_("Text Search for Checked Items"))
        self.toolbar.AddControl(self.search)
        self.search.Bind(wx.EVT_BUTTON, self.OnSearch)

        # Add a separator
        self.toolbar.AddSeparator()

        # Add a Save / Print button
        self.printReport = wx.BitmapButton(self.toolbar, -1, TransanaImages.SavePrint.GetBitmap(), size=(24, 24))
        self.printReport.SetToolTipString(_("Save / Print"))
        self.toolbar.AddControl(self.printReport)
        self.printReport.Bind(wx.EVT_BUTTON, self.OnPrintReport)

        # Add a separator
        self.toolbar.AddSeparator()

        # Add a Help button
        self.help = wx.BitmapButton(self.toolbar, -1, TransanaImages.ArtProv_HELP.GetBitmap(), size=(24, 24))
        self.help.SetToolTipString(_("Help"))
        self.toolbar.AddControl(self.help)
        self.help.Bind(wx.EVT_BUTTON, self.OnHelp)

        # Create the Close button
        self.closeButton = wx.BitmapButton(self.toolbar, -1, TransanaImages.Exit.GetBitmap(), size=(24, 24))
        self.closeButton.SetToolTipString(_("Close"))
        self.toolbar.AddControl(self.closeButton)
        self.closeButton.Bind(wx.EVT_BUTTON, self.OnOK)

        # Finalize the Toolbar
        self.toolbar.Realize()

        # Define a sizer for the form
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Place a Notebook Control on the Form
        self.notebook = wx.Notebook(self, -1)

        # Create a panel for the Word Frequency Results
        resultsPanel = wx.Panel(self.notebook, -1)
        resultsPanel.SetBackgroundColour(wx.WHITE)
        # Create a Sizer for the Results Panel
        pnl1Sizer = wx.BoxSizer(wx.VERTICAL)
        # Put a CheckListCtrl for the Word Frequency Results on the Results Panel
        self.resultsList = CheckListCtrl(resultsPanel)
        # Let the resultsList Control know about the parent form so it can provide feedback about
        # sort column and direction
        self.resultsList.SetParentForm(self)

        # Create an ImageList for the ResultsList
        self.imageList = wx.ImageList(16, 16)
        # Add images for unchecked and checked for list items
        self.unCheckedImage = self.imageList.Add(TransanaImages.NoCheck.GetBitmap())
        self.checkedImage = self.imageList.Add(TransanaImages.Check.GetBitmap())
        # Add the Column Header Sort images
        self.sortDownImage = self.imageList.Add(TransanaImages.SmallDnArrow.GetBitmap())
        self.sortUpImage = self.imageList.Add(TransanaImages.SmallUpArrow.GetBitmap())
        # Add the Column Header Sort images
        self.sortDownImage2 = self.imageList.Add(TransanaImages.SmallDnArrow.GetBitmap())
        self.sortUpImage2 = self.imageList.Add(TransanaImages.SmallUpArrow.GetBitmap())
        # Add the Column Header Sort images
        self.sortDownImage3 = self.imageList.Add(TransanaImages.SmallDnArrow.GetBitmap())
        self.sortUpImage3 = self.imageList.Add(TransanaImages.SmallUpArrow.GetBitmap())
        # Associate the ImageList with the Control
        self.resultsList.SetImageList(self.imageList, wx.IMAGE_LIST_SMALL)
        
        # Place the Results List on the Result Panel Sizer
        pnl1Sizer.Add(self.resultsList, 1, wx.EXPAND | wx.ALL, 5)
        # Assign the Results Panel Sizer to the Results Panel
        resultsPanel.SetSizer(pnl1Sizer)

        # Create a horizontal Sizer for the Synonyms controls
        addSynonymSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Create a button for adding checked items to a Synonym Group
        addSynonymBtn = wx.Button(resultsPanel, -1, _("Group Checked Items"))
        addSynonymSizer.Add(addSynonymBtn, 0, wx.EXPAND | wx.RIGHT, 10)
        addSynonymBtn.Bind(wx.EVT_BUTTON, self.OnSetSynonyms)
        # Create a Text Control for naming the Synonym Group
        self.synonymGroup = wx.TextCtrl(resultsPanel, -1)
        # Let the ResultsList control know this is the field that needs updating for setting the synonmy group name
        self.resultsList.SetSynonymGroupCtrl(self.synonymGroup)
        addSynonymSizer.Add(self.synonymGroup, 1, wx.EXPAND | wx.RIGHT, 10)
        # Add a button for a hidden group of words, the "Do Not Show" group.
        self.addNoShowBtn = wx.Button(resultsPanel, -1, _('Do Not Show Checked Items'))
        addSynonymSizer.Add(self.addNoShowBtn, 0, wx.EXPAND | wx.RIGHT, 10)
        self.addNoShowBtn.Bind(wx.EVT_BUTTON, self.OnSetSynonyms)
        # Add button for editing the selected Synonym Group
        editSynonymsBtn = wx.Button(resultsPanel, -1, _('Edit Word Group'))
        editSynonymsBtn.Bind(wx.EVT_BUTTON, self.OnEditSynonym)
        addSynonymSizer.Add(editSynonymsBtn, 0, wx.EXPAND)
        # Add the Synonyms Controls' Sizer to the notebook Panel's sizer
        pnl1Sizer.Add(addSynonymSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # Add a page to the Notebook control for the Results tab
        self.notebook.AddPage(resultsPanel, _("Results"))


        # Add a Panel for a Synonym Seeking tab
        synonymPanel = wx.Panel(self.notebook, -1)

        # Create a vertical Sizer for the Synonym Seeking Panel
        pnl3Sizer = wx.BoxSizer(wx.VERTICAL)
        # Create a horizontal Sizer for the top row of controls
        pnl3HSizer1 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the Synonym Ending prompt
        txt = wx.StaticText(synonymPanel, -1, _("Word Ending:"))
        pnl3HSizer1.Add(txt, 0, wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Add the Synonym Ending Text Control
        self.synonymExtension = wx.TextCtrl(synonymPanel, -1)
        pnl3HSizer1.Add(self.synonymExtension, 1, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Add the button to apply the extension to the Word List
        self.btnSynonymExtension = wx.Button(synonymPanel, -1, _("Apply Pattern"))
        # Define the button press handler
        self.btnSynonymExtension.Bind(wx.EVT_BUTTON, self.OnSynonymExtension)
        pnl3HSizer1.Add(self.btnSynonymExtension, 1, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Add the top row to the vertical Sizer for the panel
        pnl3Sizer.Add(pnl3HSizer1, 0, wx.EXPAND | wx.ALL, 5)

        # Create a TextCtrl to display the results of Synonym Seeking
        self.synonymResults = wx.ListCtrl(synonymPanel, -1, style=wx.LC_REPORT | wx.LC_SORT_ASCENDING)
        # Place the Synonym Seeking Results List on the Synonym Panel Sizer
        pnl3Sizer.Add(self.synonymResults, 1, wx.EXPAND | wx.ALL, 5)

        # Add a button for deleting false positive results
        self.btnDelete = wx.Button(synonymPanel, -1, _("Delete Selected Word Group"))
        # Define the button press handler
        self.btnDelete.Bind(wx.EVT_BUTTON, self.OnDeleteSynonym)
        pnl3Sizer.Add(self.btnDelete, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        # Assign the Synonym Seeking Sizer to the synonym Panel
        synonymPanel.SetSizer(pnl3Sizer)
        
        # Add a page to the Notebook control for the Synonym Seeking tab
        self.notebook.AddPage(synonymPanel, _("Group Words by Pattern"))




##        # Add a Panel for a Word Cloud output tab
##        self.wordCloudPanel = wx.Panel(self.notebook, -1)
##        self.wordCloudPanel.SetBackgroundColour(wx.BLUE)
##        
##        # Add a page to the Notebook control for the Word Cloud tab
##        self.notebook.AddPage(self.wordCloudPanel, _("Word Cloud"))



        
        # Add a Panel for an Options tab
        self.optionsPanel = wx.Panel(self.notebook, -1)
        self.optionsPanel.SetBackgroundColour(wx.WHITE)
        # Create a Sizer for the Options Panel
        pnl2Sizer = wx.FlexGridSizer(rows=5, cols=4, hgap=20, vgap=20)
        self.optionsPanel.SetSizer(pnl2Sizer)

        # Minimum Word Frequency
        txt1 = wx.StaticText(self.optionsPanel, -1, "Minimum Frequency:", style=wx.ALIGN_RIGHT)
        self.minFrequency = wx.TextCtrl(self.optionsPanel, -1, "1")

        # Minimum Word Length
        txt2 = wx.StaticText(self.optionsPanel, -1, "Minimum Word Length:", style=wx.ALIGN_RIGHT)
        self.minLength = wx.TextCtrl(self.optionsPanel, -1, "1")

        # Clear Word Groupings
        self.btnClearAll = wx.Button(self.optionsPanel, -1, "Clear all Word Grouping Data", style=wx.ALIGN_CENTER)
        self.btnClearAll.Bind(wx.EVT_BUTTON, self.OnClearAllWordGroupings)

        # Use the GridSizer to center our data entry fields both horizontally and vertically
        pnl2Sizer.AddMany([((1, 1), 12, wx.EXPAND),
                           ((1, 1), 3, wx.EXPAND),
                           ((1, 1), 1, wx.EXPAND),
                           ((1, 1), 14, wx.EXPAND),

                           ((1, 1), 12, wx.EXPAND),
                           (txt1, 3, wx.EXPAND | wx.ALIGN_RIGHT),
                           (self.minFrequency, 1, wx.EXPAND),
                           ((1, 1), 14, wx.EXPAND),

                           ((1, 1), 12, wx.EXPAND),
                           (txt2, 3, wx.EXPAND | wx.ALIGN_RIGHT),
                           (self.minLength, 1, wx.EXPAND),
                           ((1, 1), 14, wx.EXPAND),
                           
                           ((1, 1), 12, wx.EXPAND),
                           ((1, 1), 3, wx.EXPAND),
                           (self.btnClearAll, 1, wx.EXPAND),
                           ((1, 1), 14, wx.EXPAND),

                           ((1, 1), 12, wx.EXPAND),
                           ((1, 1), 3, wx.EXPAND),
                           ((1, 1), 1, wx.EXPAND),
                           ((1, 1), 14, wx.EXPAND)])

        # Make the top and bottom rows growable to center vertically
        pnl2Sizer.AddGrowableRow(0, 12)
        pnl2Sizer.AddGrowableRow(4, 14)
        # Make the left and right columns growable to center horizontally
        pnl2Sizer.AddGrowableCol(0, 12)
        pnl2Sizer.AddGrowableCol(3, 14)
        
        # Assign the Results Panel Sizer to the Results Panel
        self.optionsPanel.SetSizer(pnl2Sizer)
        # Add a page to the Notebook control for the Options tab
        self.notebook.AddPage(self.optionsPanel, _("Options"))
        
        # Add the Notebook control to the form's Main Sizer
        mainSizer.Add(self.notebook, 1, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 10)

        # Add in the Column Sorter mixin to make the results panel sortable
        ListCtrlMixins.ColumnSorterMixin.__init__(self, 3)
        self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.OnNotebookPageChanged)

        # Populate the Word Frequency Results
        self.PopulateWordFrequencies()

        # Set the form's Main Sizer and enable auto-layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)

        # Center the display on the screen
        if __name__ != '__main__':
            TransanaGlobal.CenterOnPrimary(self)
        else:
            self.CenterOnScreen()

        # If a Control Object has been passed in ...
        if self.ControlObject != None:
            # ... register this report with the Control Object (which adds it to the Windows Menu)
            self.ControlObject.AddReportWindow(self)

        # Define the Form Close event, which removes the report from Transana's Window menu
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # Display the form
        self.Show(True)

    def GetSecondarySortValues(self, col, key1, key2):
        """ Set the secondary sort for Column Sorting on the Results tab """
        # Determine whether the sort is ascending or descending
        ascending = self.GetSortState()[1]

        # If the sort is Ascending ...
        if ascending:
            # ... we want key1, then key2 values so the Word sort is alphabetic and ascending
            return(self.itemDataMap[key1][0].lower(), self.itemDataMap[key2][0].lower())
        # If the sort is Descending ...
        else:
            # ... then we want key2, then key1 values so the Word sort is still alphabetic and ascending
            return(self.itemDataMap[key2][0].lower(), self.itemDataMap[key1][0].lower())
        
    def OnNotebookPageChanged(self, event):
        """ Handle Notebook Page Change Events """
        # Allow the underlying Control to process the Page Change
        event.Skip()
        # If we're changing to the Results Tab ...
        if event.GetSelection() == 0:
            # ... enable the check and print buttons
            self.checkAll.Enable(True)
            self.uncheckAll.Enable(True)
            self.printReport.Enable(True)
        
            # If we're moving to the Results page FROM the Options page ...
            # (The Word Groups page updates the Results as needed)
            if event.GetOldSelection() == 2:
                # ... refresh the Results!
                self.PopulateWordFrequencies()
                
        else:
            # ... disable the check and print buttons
            self.checkAll.Enable(False)
            self.uncheckAll.Enable(False)
            self.printReport.Enable(False)

    def OnEditSynonym(self, event):
        """ Edit a Synonym Group """
        # Get the currently selected item, ignoring check status or multiple selections
        sel = self.resultsList.GetFirstSelected()
        # Initialize the mapItem value
        mapItem = -1
        # If there IS a currently focussed item ...
        if sel > -1:
            # ... get the mapItem for the focussed item
            mapItem = self.resultsList.GetItemData(sel)
        # If there is NOT a currently focussed item ...
        else:
            # ... search through the itemDataMap ...
            for key, data in self.itemDataMap.items():
                # ... if we find the Do Not Show Group (assuming one has been defined) ...
                if (data[0] == "Do Not Show Group"):
                    # ... note it's mapItem
                    mapItem = key
                    # ... and stop searching
                    break

        # If a mapItem has been found ... (it won't be if there's no focussed item and no Do Not Show Group entry)
        if mapItem > -1:
            # Create a SynonymEditor form object
            synonymEditor = SynonymEditor.SynonymEditor(self, self.itemDataMap[mapItem])
            # Display the Synonym Editor and get the results from the user
            result = synonymEditor.ShowModal()
            # If the user pressed OK ...
            if result == wx.ID_OK:
                # ... get the old synonym group name and list of synonyms as well as the changed values from the form
                oldGroupName = self.itemDataMap[mapItem][0]
                oldSynonyms = self.itemDataMap[mapItem][2]
                newGroupName = synonymEditor.synonymGroup.GetValue()
                newSynonyms = synonymEditor.GetSynonymValues()

                # If the Synonym Group Name has changed ...
                if oldGroupName != newGroupName:
                    # ... iterate through the OLD synonym list ...
                    for synonym in oldSynonyms:
                        # ... and delete each synonym from the database ...
                        DBInterface.DeleteSynonym(oldGroupName, synonym)
                        # ... and remove it from the synonym lookup dictionary
                        del(self.synonymLookups[synonym])
                    # Remove the old synonym group from the synonyms dictionary
                    del(self.synonyms[oldGroupName])
                    # Now iterate through the NEW synonyms list ...
                    for synonym in newSynonyms:
                        # ... adding the new synonyms to the database ...
                        DBInterface.AddSynonym(newGroupName, synonym)
                        # ... and the synonym lookup dictionary
                        self.synonymLookups[synonym] = newGroupName
                # If the Synonym Group Name did NOT change ...
                else:
                    # ... iterate through the OLD synonym list ...
                    for synonym in oldSynonyms:
                        # ... and if the synonym is not in the NEW synonym list ...
                        if not synonym in newSynonyms:
                            # ... delete it from the database ...
                            DBInterface.DeleteSynonym(oldGroupName, synonym)
                            # (Start exception handling in case the synonym isn't in the Lookup dictionary)
                            try:
                                # ... and try to remove it from the synonym Lookup dictionary
                                del(self.synonymLookups[synonym])
                            except KeyError:
                                pass

                # If there is only one entry in the new Synonym List and it's the same as the new Synonym Group Name ...
                if (len(newSynonyms) == 1) and (newGroupName == newSynonyms[0]):
                    # ... we can remove this entry from the database
                    DBInterface.DeleteSynonym(newGroupName, newSynonyms[0])
                    # We can also remove the Synonym Lookup and Synonyms Dictionary entries
                    del(self.synonymLookups[newSynonyms[0]])
                    del(self.synonyms[newGroupName])
                # If there are more than one entries in the new Synonym List ...
                elif len(newSynonyms) >= 1:
                    # ... we should set the Synonyms List dictionary list to this new list, replacing the old one
                    self.synonyms[newGroupName] = newSynonyms
                # If the new Synonyms List is empty ...
                else:
                    # ... start exception handling
                    try:
                        # Try to delete the Synonyms Dictionary entry
                        del(self.synonyms[newGroupName])
                    # If an exception is raised, we can ignore it.
                    except KeyError:
                        pass

                # Clear the itemDataMap so that the table will be properly refeshed
                # This actually requires using the ControlObject to signal ALL Word Frequency Reports!
                self.ControlObject.SignalWordFrequencyReports()
                
                # Populate the Word Frequencies table
                self.PopulateWordFrequencies()
                
            # Destroy the Synonmy Editor dialog
            synonymEditor.Destroy()

    def OnSynonymExtension(self, event):
        """ Handle Synonym Seeking requests """
        # Clear the Synonym Seeking Results display
        self.synonymResults.ClearAll()
        # Add Column Heading
        self.synonymResults.InsertColumn(0, _("Word"))
        self.synonymResults.InsertColumn(1, _("Words in Group"))

        # Check that a synonym extension has been defined
        if self.synonymExtension.GetValue().strip() == '':
            # If not, report it to the user
            index = self.synonymResults.InsertStringItem(sys.maxint, _("Please enter a Word Extension."))
            # We can't do any more here
            return
        # If a Synonym Extension is defined ...
        else:
            # ... remember it!
            synonymExtension = self.synonymExtension.GetValue().strip().lower()

        # We need a list of all the words in our display.
        words = []
        # The itemDataMap has all the words that are in the control.  Iterate through it ...
        for key in self.itemDataMap.keys():
            # ... and grab the words it contains
            words.append(self.itemDataMap[key][0])

        # Sort the Words list for better display to the user            
        words.sort()
        # Iterate through all the words
        for word in words:
            # If there's a version of the word WITH the extension ...
            if word + synonymExtension in words:
                # ... report that it has been found to the user

                # Create a new row in the Results List, adding the word or Synonym Group text
                index = self.synonymResults.InsertStringItem(sys.maxint, word)
                # Add the Count information
                self.synonymResults.SetStringItem(index, 1, word + synonymExtension)

                # If our word already HAS a synonym entry ...
                if word in self.synonymLookups.keys():
                    # ... add the extended version to the synonyms list for the synonym group
                    self.synonyms[self.synonymLookups[word]].append(word + synonymExtension)
                # If our word does NOT have a synonym entry ...
                else:
                    # ... add the original word to the synonym lookup data
                    self.synonymLookups[word] = word
                    # ... and add it to the database
                    DBInterface.AddSynonym(word, word)
                    # ... and add the extended version to the synonyms table, creating a synonyms group for it
                    self.synonyms[word] = [word, word + synonymExtension]
                # Now add the extended version to teh synonym lookup data
                self.synonymLookups[word + synonymExtension] = word
                # ... and add the extended version to the database
                DBInterface.AddSynonym(word, word + synonymExtension) 

        # This code comes here, as width is calculated immediately, so must be AFTER data population.        
        self.synonymResults.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        self.synonymResults.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # If the Word Column is too narrow to show the column header ...
        if self.synonymResults.GetColumnWidth(0) < 150:
            # ... widen the column
            self.synonymResults.SetColumnWidth(0, 150)
        # If the Words in Group Column is too narrow to show the column header ...
        if self.synonymResults.GetColumnWidth(1) < 200:
            # ... widen the column
            self.synonymResults.SetColumnWidth(1, 200)
        # Clear the itemDataMap so that the table will be properly refeshed
        # This actually requires using the ControlObject to signal ALL Word Frequency Reports!
        self.ControlObject.SignalWordFrequencyReports()
        # We need to update the Word Frequencies table before the next Synonym Lookup occurs, so call this
        # no matter what!
        self.PopulateWordFrequencies()

        # Clear the extension specification control
        self.synonymExtension.SetValue('')
        # Set the focus to the extension specification control, ready for the next synonym seek.
        self.synonymExtension.SetFocus()

    def OnDeleteSynonym(self, event):
        """ Event Handler for the Delete Synonym button on the Synonym Patterns Page """
        # Get the first selected item in the control
        sel = self.synonymResults.GetFirstSelected()
        # We need to delete items from the bottom up to avoid problems with changing item numbers.  So let's
        # remember the items to delete
        itemsToDelete = []
        # While there are selections in the control we have not yet processed ...
        while sel > -1:
            # ... add the item number to the list of items to delete ...
            itemsToDelete.append(sel)
            # ... and get the next selected item, if there is one
            sel = self.synonymResults.GetNextSelected(sel)

        # Sort the list of items to delete in reverse order
        itemsToDelete.sort(reverse=True)
        # Iterate through the list of items to delete
        for item in itemsToDelete:
            # Determine the synonym group and synonym value for each item
            synonymGroup = self.synonymResults.GetItemText(item, 0)
            synonym = self.synonymResults.GetItemText(item, 1)

            # It's possible that the Synonym Group is actually a synonym itself!  Check for that.
            if synonymGroup in self.synonymLookups.keys():
                # If so, use the value from the lookup dictionary rather than from the control
                synonymGroup = self.synonymLookups[synonymGroup]
            # Delete the synonym from the database
            DBInterface.DeleteSynonym(synonymGroup, synonym)
            # Start exception handling
            try:
                # Delete the value from the Synonym Lookup table
                del(self.synonymLookups[synonym])
            # If the synonym wasn't in the Lookup Table, we can ignore it!
            except KeyError:
                pass

            # Get the current Synonyms for the Synonym Group
            newSynonyms = self.synonyms[synonymGroup]
            # Delete the current synonym from this group
            del(newSynonyms[newSynonyms.index(synonym)])

            # If there is a single new synonym left in the list ...            
            if (len(newSynonyms) == 1) and (synonymGroup == newSynonyms[0]):
                # Delete the synonym record from the database
                DBInterface.DeleteSynonym(synonymGroup, newSynonyms[0])
                # Remove the synonym from the Lookups dictionary
                del(self.synonymLookups[newSynonyms[0]])
                # Remove the synonym from the Synonyms dictionary
                del(self.synonyms[synonymGroup])
            # If there's nore than one synonym record left in the list ...
            elif len(newSynonyms) > 1:
                # ... replace the list of synonyms for this group with the new list
                self.synonyms[synonymGroup] = newSynonyms

            else:
                # ... start exception handling
                try:
                    # Try to delete the Synonyms Dictionary entry
                    del(self.synonyms[synonymGroup])
                # If an exception is raised, we can ignore it.
                except KeyError:
                    pass

            # Remove the item from the Results table
            self.synonymResults.DeleteItem(item)                    

        # This code comes here, as width is calculated immediately, so must be AFTER data population.        
        self.synonymResults.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        self.synonymResults.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # If the Word Column is too narrow to show the column header ...
        if self.synonymResults.GetColumnWidth(0) < 150:
            # ... widen the column
            self.synonymResults.SetColumnWidth(0, 150)
        # If the Words in Group Column is too narrow to show the column header ...
        if self.synonymResults.GetColumnWidth(1) < 200:
            # ... widen the column
            self.synonymResults.SetColumnWidth(1, 200)

        # Clear the itemDataMap so that the table will be properly refeshed
        # This actually requires using the ControlObject to signal ALL Word Frequency Reports!
        self.ControlObject.SignalWordFrequencyReports()
        # Repopulate the Results table based on the new Synonyms data
        self.PopulateWordFrequencies()

    def PopulateSynonyms(self):
        """ Set up the Synonyms data structures and populate them initially """
        # Initialize a dictionary for holding the Synonym Groups (key = title, value = list of synonmyms)
        self.synonyms = {}
        # Initialize a Synonym Lookup table (key = synonyms, value = Synonym Group name)
        self.synonymLookups = {}

        # If we're testing ...
        if self.tree == None or self.startNode == None:
            # ... use fake synonym data
            self.synonyms = {'a and the'    : ['a', 'and', 'the'],
                             'name'  : ['ellen', 'feiss'],
                             'pronouns'      : ['i', "i'm", 'my']}
        # If we're NOT testing ...
        else:
            # Load synonyms from database
            self.synonyms = DBInterface.GetSynonyms()

        # Now iterate though the items in the Synonym Groups
        for key in self.synonyms.keys():
            # For each synonym in the group ...
            for synonym in self.synonyms[key]:
                # ... add an entry to the Synonym Lookup table
                self.synonymLookups[synonym] = key

    def PopulateWordFrequencies(self):
        """ Clear and Populate the Word Frequency Results """
        # Signal that the control's itemDataMap IS being updated right now!
        self.isUpdating = True
        # Clear the Control
        self.resultsList.ClearAll()
        # Add Column Heading
        self.resultsList.InsertColumn(0, _("Word"))
        self.resultsList.InsertColumn(1, _("Frequency"), wx.LIST_FORMAT_RIGHT)
        self.resultsList.InsertColumn(2, _("Word Group"))

        # If we haven't read the data yet, or need to refresh it ...
        if self.needsUpdate:
            # We can reset the needsUpdate Flag
            self.needsUpdate = False
            # Re-initialize the itemDataMap
            self.itemDataMap = {}
            # We need to populate the Synonyms BEFORE we Populate the Word Frequencies!!
            self.PopulateSynonyms()
            # ... initialize a data structure
            data = {}

            # If we're testing ...
            if self.tree == None or self.startNode == None:
                # ... define sample text
                sampleText = u""" I was writing a paper on the PC,
and it was like   beep   beeps   (beep)   beeping ....  beeps ???  beep   (beep.)
And then, like, half of my [paper] was gone.
And I was like  "Huh?"
It \u201cdevoured\u201d my paper.
"It was a really good paper."
And then I had to write it again and 1,000 I had to do it fast, so it wasn't as good.
http://www.spurgeonwoods.com/test
It's kind of  (2.2)   a bummer.
I'm Ellen Feiss, and I'm a student!"""
                # Prepare the sample text
                text = self.PrepareText(sampleText)
                # Count the words in the sample text
                data = self.CountWords(text, data)

            # If we're in Transana ...
            else:
                # ... pass the tree and node information into the method that recursively extracts text from
                #     tree nodes and counts words in it
                data = self.ExtractDataFromTree(self.tree, self.startNode)

            # Initialize the item data dictionary
            itemData = {}
            # Initialize a counter, which serves as the itemData Key required by the ColumnSorterMixin
            counter = 0
            # for all the words in the data dictionary ...
            for text in data.keys():
                # See if the word has defined synonyms
                if self.synonyms.has_key(text):
                    synonyms = self.synonyms[text]
                else:
                    synonyms = ''
                # Add the data to the itemData dictionary
                itemData[counter] = (text, data[text], synonyms)
                # Increment the counter
                counter += 1

        # If the data is already defined ...
        else:
            # ... then we'll use that!
            itemData = self.itemDataMap

        # Convert the extracted data to the form needed for the ColumnSorterMixin
        for key, data in itemData.items():
            # Start exception handling
            try:
                # Convert the Minimum Frequency from the Options tab
                minFrequency = int(self.minFrequency.GetValue())
            # If it's not an integer, ignore it!
            except:
                minFrequency = 1
            # Start separate exception handling process
            try:
                # Try to convert the Minimum Word Length from the Options Tab
                minLength = int(self.minLength.GetValue())
            # If it's not an integer, ignore it!
            except:
                minLength = 1

            # If the item is not part of the "Do Not Show" Group, and it meets the minimum frequency and length requirements ...
            if (data[0] != "Do Not Show Group") and (data[1] >= minFrequency) and (len(data[0]) >= minLength):
                # ... create a new row in the Results List, adding the word or Synonym Group text
                index = self.resultsList.InsertStringItem(sys.maxint, data[0])
                # Add the Count information
                self.resultsList.SetStringItem(index, 1, str(data[1]))
                # Initialize a string for the synonyms
                synonymString = ''
                # If THIS record has Synonyms ...
                if data[0] in self.synonyms.keys():
                    # ... for each synonym ...
                    for word in self.synonyms[data[0]]:
                        # ... if this word is NOT the first synonym ...
                        if synonymString != '':
                            # ... in sert a spacer in the synonym string
                            synonymString += ' '
                        # Add the synonym to the synonym string
                        synonymString += word
                # Add the list of synonyms to the Results List
                self.resultsList.SetStringItem(index, 2, synonymString)
                # Set the Results List item's itemData to the integer Key, needed for sorting in the ColumnSorterMixin
                self.resultsList.SetItemData(index, key)

        # The ColumnSorterMixin requires this data structure with this variable name to function
        self.itemDataMap = itemData
        # This code comes here, as width is calculated immediately, so must be AFTER data population.        
        self.resultsList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        self.resultsList.SetColumnWidth(1, 100)
        self.resultsList.SetColumnWidth(2, wx.LIST_AUTOSIZE)

        # If the Text Column is too narrow to show the column header ...
        if self.resultsList.GetColumnWidth(0) < 200:
            # ... widen the column
            self.resultsList.SetColumnWidth(0, 200)
        # If the Synonyms Column is too narrow to show the column header ...
        if self.resultsList.GetColumnWidth(2) < 200:
            # ... widen the column
            self.resultsList.SetColumnWidth(2, 200)

        # Sort data by the correct column and direction
        self.SortListItems(self.sortColumn, self.sortAscending)
        # Update the Control to try to fix the first column header's appearance
        self.resultsList.Update()
        # Signal that the control's itemDataMap is no longer being updated right now!
        self.isUpdating = False

    def GetListCtrl(self):
        """ Pointer to the Results List, required for the ColumnSorterMixin """
        # Return a pointer to the ResultsList ListCtrl
        return self.resultsList

    def GetSortImages(self):
        """ Provide the images needed for headers of sort columns, part of the ColumnSorterMixin implementation """
        # Return the up and down arrow images to be displayed in the column headers
        return (self.sortUpImage, self.sortDownImage)

    def ExtractDataFromTree(self, tree, startNode):
        """ This routine takes the tree and startNode, figures out what Documents and Transcripts to load, and creates
            a data structure with all the individual words in the appropriate scope along with their counts. """
        # Ask the user to wait while the report is being assembled
        popupDlg = Dialogs.PopupDialog(self, _("Word Frequency Report"), _("Please wait ..."))
        # Initialize the data dictionary
        data = {}
        # Extract the data from the tree using this recursive method
        data = self.ExtractDataFromNode(tree, startNode, data)
        # Destroy the popup
        popupDlg.Destroy()
        # Return the data dictionary to the calling routine
        return data

    def ExtractDataFromNode(self, tree, startNode, data):
        """ This extracts data from a node, calling subnodes recursively as needed """
        # Get the Item Name and Item Data from the tree node passed in.
        itemName = tree.GetItemText(startNode)
        itemData = tree.GetPyData(startNode)

        # If the node passed in is one we should recurse into ...
        if itemData.nodetype in ['LibraryRootNode', 'LibraryNode', 'SearchLibraryNode', 'EpisodeNode', 'SearchEpisodeNode',
                                 'CollectionsRootNode', 'CollectionNode', 'SearchCollectionNode']:
            # ... get the first Child node
            (childNode, cookieItem) = tree.GetFirstChild(startNode)
            # As long as there are additional children to examine ...
            while childNode.IsOk():
                # ... get the child node's data
                childData = tree.GetPyData(childNode)
                # If the child node is a Snapshot or any type of Note ...
                if childData.nodetype in ['SnapshotNode', 'LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 
                                          'TranscriptNoteNode', 'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode']:
                    # ... we can skip it
                    pass
                # If the child node is soemthing we need to process ...
                else:
                    # ... process the node by calling this method recursively
                    data = self.ExtractDataFromNode(tree, childNode, data)
                # Try to get the next Child Node
                (childNode, cookieItem) = tree.GetNextChild(startNode, cookieItem)

        # If the node passed in is a Document Node ...
        elif itemData.nodetype in ['DocumentNode', 'SearchDocumentNode']:
            # ... load the Document ...
            record = Document.Document(num=itemData.recNum)
            # ... prepare the object's plain text ...
            text = self.PrepareText(record.plaintext)
            # ... and count the words in the prepared text
            data = self.CountWords(text, data)
        # If the node passed in is a Transcript Node ...
        elif itemData.nodetype in ['TranscriptNode', 'SearchTranscriptNode']:
            # ... load the Transcript ...
            record = Transcript.Transcript(itemData.recNum)
            # ... prepare the object's plain text ...
            text = self.PrepareText(record.plaintext)
            # ... and count the words in the prepared text
            data = self.CountWords(text, data)
        # If the node passed in is a Quote Node ...
        elif itemData.nodetype in ['QuoteNode', 'SearchQuoteNode']:
            # ... load the Quote ...
            record = Quote.Quote(num=itemData.recNum)
            # ... prepare the object's plain text ...
            text = self.PrepareText(record.plaintext)
            # ... and count the words in the prepared text
            data = self.CountWords(text, data)
        # If the node passed in is a Clip Node ...
        elif itemData.nodetype in ['ClipNode', 'SearchClipNode']:
            # ... load the Clip ...
            clipRecord = Clip.Clip(itemData.recNum)
            # ... for each Transcript associated with the Clip ...
            for record in clipRecord.transcripts:
                # ... prepare the object's plain text ...
                text = self.PrepareText(record.plaintext)
                # ... and count the words in the prepared text
                data = self.CountWords(text, data)
        # If we have a Note node ... (Does this ever happen??)
        elif itemData.nodetype in ['LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode']:
            pass
        # If we have a node of undefined type ...
        else:
            # ... we should NEVER see this, obviously!
            print "ERROR:  ", tree.GetItemText(startNode).encode('utf8'), " NOT PROCESSED.  Wrong Node Type.", itemData.nodetype

        # Return the extracted Word Count data
        return data

    def PrepareText(self, text):
        """ This method cleans up the messy PlainText that comes in, removing time codes, punctuation, etc. """

        # Strip Time Codes
##        regex = "%s<[\d]*>" % TransanaConstants.TIMECODE_CHAR
##        text = re.sub(regex, '', text
## INSTEAD OF...
##        reg = re.compile(regex)
##        pos = 0
##        for x in reg.findall(text):
##            pos = text.find(x, pos, len(text))
##            text = text[ : pos] + text[pos + len(x) : ]
## IF THIS WORKS, ALSO UPDATE PlainTextUpdate.py!!

        # strip multiple periods, questions marks, plus signs, exclamation points, hyphens
        text = re.sub('[\.?:\*+!-][\.?:\*+!-]+', ' ', text)
        # Strip parentheses, brackets, braces, quotation marks, slashes, ampersands, equal signs, asterisks, pound sign(hashtag),
        # less than, greater than
        text = re.sub('[()\[\]{}"/&=\*#<>]', ' ', text)
        # Apostrophes PRECEDED by white space
        text = re.sub("\s'", ' ', text)
        # Remove certain unicode characters that seem to cause problems (smart quotes, double-angle quotes, etc.)
        # Also, the 4 Jeffersonian special symbols
        text = re.sub(u'\u00ab|\u00b0|\u00bb|\u2018|\u2019|\u2022|\u201c|\u201d|\u2039|\u203a|\u2191|\u2193', u'', text)
        # Strip commas, periods, questions marks, exclamation points, colons, semicolons, and apostrophes
        # followed by whitespace (but NOT those NOT followed by white space, leaving "2.2", "1,000" and
        # "won't" intact.)
        text = re.sub('[,\.?!:;\']\s', ' ', text)
        # String final punctuation at tend of the string, not followed by whitespace
        text = re.sub('[\.?!\']$', '', text)
        # Strip multiple whitespace characters, replacing all whitespace with single spaces
        text = re.sub('\s+', '\n', text)

        # Return the prepared text
        return text

    def CountWords(self, text, words = {}):
        """ This method takes prepared text (see above) and adds it to existing WordCount data. """
        # DO NOT initialize a dictionary to hold word counts (key = word, value = count) here.
        # Instead, this structure can be passed in to allow additional text to be added.
        
        # For each line of the prepared text ...
        for line in text.split('\n'):
            # ... remove whitespace and compensate for different cases
            word = line.strip().lower()
            # If the word exists in the synonymLookup dictionary ...
            if word in self.synonymLookups.keys():
                # ... substitute the synonym for the original word
                word = self.synonymLookups[word]
            # There are certain "words" that should not be included.  These can slip past PrepareText.
            if not word in ['', '-', ':']:
                # If the word is already in the dictionary ...
                if words.has_key(word):
                    # ... increment the count
                    words[word] += 1
                # If the word is not in the dictionary ...
                else:
                    # ... add it with a count of one.
                    words[word] = 1
        # Return the word dictionary
        return words

    def OnCheck(self, event):
        """ Handle Check and Uncheck Buttons """
        # if we're on the Results tab of the Notebook ...
        if self.notebook.GetSelection() == 0:
            # ... if this was triggered by CheckAll, check is True, otherwise False
            check = event.GetId() == self.checkAll.GetId()

            # Get the first selected item
            sel = self.resultsList.GetFirstSelected()
            # As long as there are more selected items ...
            while sel > -1:
                # ... check the current item ...
                self.resultsList.CheckItem(sel, check)
                # ... and move on to the next selected item
                sel = self.resultsList.GetNextSelected(sel)

    def OnSearch(self, event):
        """ Request Text Search from the Word Frequency Report """
        # Initialize variables for counting the number of search terms and building a search name
        termCount = 0
        searchName = ''
        # Create an empty list for search terms
        searchTerms = []

        # Determine the number of items in the Results List
        count = self.resultsList.GetItemCount()
        # Iterate through the Results List
        for itemId in range(count):
            
            # If the item is checked ...
            if self.resultsList.IsChecked(itemId):

                # If there are no Synonyms ...
                if self.resultsList.GetItemText(itemId, 2) == '':
                    # ... add the word itself to the Search terms, using "Word Text" instead of "Item Text" to signal we
                    #     want whole words instead of plain text search.
                    searchTerms.append(u'Word Text contains "%s" OR ' % self.resultsList.GetItemText(itemId, 0))
                    # Increment the counter
                    termCount += 1
                    # If no Search Name has been defined yet ...
                    if searchName == '':
                        # ... add the first search term to it
                        searchName = u'"%s"' % self.resultsList.GetItemText(itemId, 0)
                        
                # If there ARE Synonmyms ...
                else:
                    # Add each word in the Synonyms List
                    for word in self.resultsList.GetItemText(itemId, 2).split(' '):
                        # Add the word to the search terms, using "Word Text" instead of "Item Text" to signal we
                        # want whole words instead of plain text search.
                        searchTerms.append(u'Word Text contains "%s" OR ' % word)
                        # Increment the counter
                        termCount += 1
                        # If no Search Name has been defined yet ...
                        if searchName == '':
                            # ... add the first search term to it
                            searchName = u'"%s"' % word

        # If there are 2 or more search terms, modify the search name
        if termCount == 2:
            searchName += _(' and %s other') % (termCount - 1)
        elif termCount > 2:
            searchName += _(' and %s others') % (termCount - 1)

        # if there's at least one search term ...
        if termCount > 0:
            # ... remove the " OR " from the final entry in the Search Terms
            searchTerms[-1] = searchTerms[-1][:-4]
            # ... finalize the Search Name
            searchName = 'Word Freq for %s' % searchName
            # .. and call ProcessSearch to execute the search terms we've assembled here
            search = ProcessSearch.ProcessSearch(self.tree, self.tree.searchCount, searchName=searchName, \
                                                 searchTerms=searchTerms, searchScope=self.startNode)

    def OnPrintReport(self, event):
        """ Handle requests for a printable report """
        # Create a Report Frame
        self.report = TextReport.TextReport(self, -1, _("Word Frequency Report"), self.ConstructReport)
        # To speed report creation, freeze GUI updates based on changes to the report text
        self.report.reportText.Freeze()
        # Trigger the ReportText method that causes the report to be displayed.
        self.report.CallDisplay()
        # Now that we're done, remove the freeze
        self.report.reportText.Thaw()

    def ConstructReport(self, reportText):
        """ Populate the TextReport RTFControl for the Word Frequency Report """
        # Make the control writable
        reportText.SetReadOnly(False)
        # Set the font for the Report Title
        reportText.SetTxtStyle(fontFace = 'Courier New', fontSize = 16, fontBold = True, fontUnderline = True,
                               parAlign = wx.TEXT_ALIGNMENT_CENTER, parSpacingAfter = 42)
        # Add the Report Title
        reportText.WriteText(self.title)
        reportText.Newline()
        # Set the Style for the main report header
        reportText.SetTxtStyle(fontFace = 'Courier New', fontSize = 14, fontBold = True, fontUnderline = True,
                               parAlign = wx.TEXT_ALIGNMENT_LEFT, parLeftIndent = (0, 900), parSpacingAfter = 12,
                               parTabs=[600, 900])
        # Write column headers
        reportText.WriteText("%s\t%s\t%s\n" % (_("Word"), _("Frequency"), _("Word Group") ))
        # Adjust the style for the main report data
        reportText.SetTxtStyle(fontSize = 12, fontBold = False, fontUnderline = False)
        # Iterate through the items in the Results List ...
        for count in range(self.resultsList.GetItemCount()):
            # ... and add the data to the report
            reportText.WriteText("%s\t%s\t%s\n" % (self.resultsList.GetItemText(count, 0), self.resultsList.GetItemText(count, 1),
                                                self.resultsList.GetItemText(count, 2) ))

    def OnHelp(self, event):
        # ... call Help!
        self.ControlObject.Help("Word Frequency Report")

    def OnOK(self, event):
        """ Handle the OK / Close button """
        # Close the form
        self.Close()

    def OnClose(self, event):
        """ Handle Form Closure """
        # If there is a defined Control Object ...
        if self.ControlObject != None:
            # ... remove this form from the Windows menu
            self.ControlObject.RemoveReportWindow(self.title, self.reportNumber)
        # Inherit the parent Close event so things will, you know, close.
        event.Skip()

    def OnSetSynonyms(self, event):
        """ Handle the Set Synonym Group buttons """
        # If called by the "Do Not Show" group buttong ...
        if event.GetId() == self.addNoShowBtn.GetId():
            # ... we should use the "Do Not Show Group" synonym group ...
            synonymGroup = 'Do Not Show Group'
        # If we are adding a visible synonym ...
        else:
            # ... get the name of the synonym group
            synonymGroup = self.synonymGroup.GetValue().strip()
            # If the synonym group name is empty ...
            if synonymGroup == '':
                # ... set the focus to the synonym group field ...
                self.synonymGroup.SetFocus()
                # ... and stop.  There's nothing to do!
                return

        # Initialize a List to hold Synonyms
        synonymData = []
        # Initialize a String to hold the string representation of the Synonyms List
        synonymString = ''

        # if the Synonym Group already exists ...
        if synonymGroup in self.synonyms.keys():
            # Get the current synonyms as a starting point
            synonymData = self.synonyms[synonymGroup]

        # We may need to remember an OLD name for a synonym group.  Initialize a variable for this.
        oldSynonymGroup = ''
        # Determine the number of items in the Results List
        count = self.resultsList.GetItemCount()
        # Iterate through the Results List
        for itemId in range(count):
            # Get the Item Data value for the current Results List item
            item = self.resultsList.GetItemData(itemId)
            
            # If the item matches the Synonym Group name or is checked ...
            if (self.itemDataMap[item][0] == synonymGroup) or \
               (self.resultsList.IsChecked(itemId)):

                # If the item matches the Synonym Group name AND this group is already known ...
                if (self.itemDataMap[item][0] == synonymGroup) and \
                   (synonymGroup in self.synonyms.keys()):
                    # ... we don't need to do anything!
                    pass
                # If not, do we already have synonyms for THIS item?
                elif self.itemDataMap[item][2] != '':
                    # If so, remember the original synonym group name for this group ...
                    oldSynonymGroup = self.itemDataMap[item][0]
                    # ... and add the synonyms to the NEW synonym DATA
                    synonymData += self.itemDataMap[item][2]
                    # For each of the OLD synonyms for this item ...
                    for tmpItem in self.itemDataMap[item][2]:
                        # ... update the synonym Lookup dictionary with the new synonym group name
                        self.synonymLookups[tmpItem] = synonymGroup
                        # ... and update the database to use the NEW synonym group name
                        DBInterface.UpdateSynonym(oldSynonymGroup, tmpItem, synonymGroup, tmpItem)
                    # Finally, delete the OLD group from the Synonyms List
                    del(self.synonyms[oldSynonymGroup])
                # If we have a NEW synonym Group ...
                else:
                    # ... add it to the synonym DATA
                    synonymData.append(self.itemDataMap[item][0])
                    # Add the item to the synonym lookup dictionary
                    self.synonymLookups[self.itemDataMap[item][0]] = synonymGroup
                    # add the item to the Database
                    DBInterface.AddSynonym(synonymGroup, self.itemDataMap[item][0]) 

        # If the user pressed the Do Not Show button without checking any items ...
        if (synonymGroup == 'Do Not Show Group') and (len(synonymData) == 0):
            # ... there's nothing to do!
            return

        # Sort the Synonym Data
        synonymData.sort()
        # Create a String version of the synonym list for display
        for synonym in synonymData:
            if len(synonymString) > 0:
                synonymString += ' '
            synonymString += synonym

        # Update the permanent Synonyms data
        self.synonyms[synonymGroup] = synonymData

        # Initialize a flag for the index of an item, assuming the item will not be found
        index = -1
        # Now iterate though the items in the Synonyms List (control)
        for val in range(self.resultsList.GetItemCount()):
            # If the Synonym Group is found, we can update it
            if self.resultsList.GetItem(val, 0).GetText() == synonymGroup:
                # Set the index to the Results List item number ...
                index = val
                # ... and stop looking
                break
        # If the Synonym Group was not found ...
        if index == -1:
            # ... add a new on to the Synonym List (control)
            index = self.resultsList.InsertStringItem(sys.maxint, synonymGroup)
        # Set the value to the string representation of the synonyms
        self.resultsList.SetStringItem(index, 2, synonymString)

        # Now we need to look for the synonyms that were just defined in the data, consolidating them.
        itemIndex = -1
        itemValue = 0
        # Iterate though the itemDataMap
        for key in self.itemDataMap.keys():
            # Here's the trick.  The Synonyms Group name may not be in the sysnonymData list!!
            # So ... if the itemDataMap item is in the synonym list OR
            #        if the itemDataMap item matches the new Synonym Group name OR
            #        if the itemDataMap item is in the OLD synonym list ...
            if (self.itemDataMap[key][0] in synonymData) or (self.itemDataMap[key][0] == synonymGroup) or \
               (self.itemDataMap[key][0] == oldSynonymGroup):
                # ... remember the item's Map value
                itemValue += self.itemDataMap[key][1]
                # If our Map Data item matchs the Synonym Group name ...
                if self.itemDataMap[key][0] == synonymGroup:
                    # ... remember the item index
                    itemIndex = key
                # If the item is not Synonym Group name
                else:
                    # ... then we want to remove this item from the Item Data Map, as it is becomeing a synonym for something else!
                    del(self.itemDataMap[key])
        # If we did not match the Synonym Group name ...
        if itemIndex == -1:
            # ... we add a NEW item to the Item Data Map
            self.itemDataMap[len(self.itemDataMap)] = (synonymGroup, itemValue, synonymData)
        # If we DID match the Synonym Group Name ...
        else:
            # ... then update the existing item for that group
            self.itemDataMap[itemIndex] = (synonymGroup, itemValue, synonymData)

        # Because we've changed synonyms, we need to tell ALL Word Frequency Reports that they
        # need to repopulate their list contents.  But we can skip THIS report, as that's already
        # been done above.
        self.ControlObject.SignalWordFrequencyReports(self.reportNumber)

        # Remember the current scroll position of the Results List
        scrollPos = self.resultsList.GetScrollPos(wx.VERTICAL)
        # Repopulate the Word Frequencies
        self.PopulateWordFrequencies()
        # Reset the Synonym Group name to blank
        self.synonymGroup.SetValue('')
        # Freeze the Results List
        self.resultsList.Freeze()
        # Scroll to the original position
        self.resultsList.ScrollLines(scrollPos)
        # Thaw the Control
        self.resultsList.Thaw()

    def OnClearAllWordGroupings(self, event):
        prompt = _("Are you SURE you want to delete all Word Groupings?")
        # Build an error dialog
        dlg = Dialogs.QuestionDialog(None, prompt)
        dlg.CentreOnScreen()
        # If the user chooses to overwrite ...
        if dlg.LocalShowModal() == wx.ID_YES:
            # Delete all of the Word Groupings        
            DBInterface.ClearAllSynonyms()
            # Clear the itemDataMap so that the table will be properly refeshed
            # This actually requires using the ControlObject to signal ALL Word Frequency Reports!
            self.ControlObject.SignalWordFrequencyReports()
            # Repopulate the Word Frequencies
            self.PopulateWordFrequencies()
        dlg.Destroy()
        

if __name__ == '__main__':
    class MyApp(wx.App):
       def OnInit(self):
          frame = WordFrequencyReport(None, None, None)
          self.SetTopWindow(frame)
          return True
          

    app = MyApp(0)
    app.MainLoop()
