#Copyright (C) 2002-2015  The Board of Regents of the University of Wisconsin System

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
# Import Transana's Document object
import Document
# Import Transana's Quote object
import Quote
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

    def OnItemActivated(self, evt):
        """ Handle item double-click """
        # Toggle the checked state of the selected item
        self.ToggleItem(evt.m_itemIndex)

    def OnColumnClick(self, event):
        """ Handle setting the sort by clicking the column header """
        # Determine which column has been selected
        col = event.GetColumn()
        # Column image is a proxy for sort direction!!  I don't seem to be able to get that information directly.
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
        # Remember the tree that is passed in
        self.tree = tree
        # Remember the start node passed in
        self.startNode = startNode
        # We need a specifically-named dictionary in a particular format for the ColumnSorterMixin.
        # Initialize that here.  (This also tells us if we've gotten the data yet!)
        self.itemDataMap = {}
        # We also need a synonyms list.  Initialize it here
        self.synonyms = {}
        # Let's keep track of what column and what direction our current sort is.  That way, we
        # can recreate it as we manipulate the list.  Start with a Descending sort of the Count
        self.sortColumn = 1
        self.sortAscending = False

        # Determine the screen size for setting the initial dialog size
        if __name__ == '__main__':
            rect = wx.Display(0).GetClientArea()  # wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
        else:
            rect = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
        width = rect[2] * .80
        height = rect[3] * .80
        # Create the basic Frame structure with a white background
        wx.Frame.__init__(self, parent, -1, _('Word Frequency Report'), size=wx.Size(width, height), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundColour(wx.WHITE)
        
        # Set the report's icon
        transanaIcon = wx.Icon(os.path.join(TransanaGlobal.programDir, "images", "Transana.ico"), wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)

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
        addSynonymBtn = wx.Button(resultsPanel, -1, _("Add Checked Items to Synonym Group"))
        addSynonymSizer.Add(addSynonymBtn, 0, wx.EXPAND | wx.RIGHT, 10)
        addSynonymBtn.Bind(wx.EVT_BUTTON, self.OnSetSynonyms)
        # Create a Text Control for naming the Synonym Group
        self.synonymGroup = wx.TextCtrl(resultsPanel, -1)
        # Let the ResultsList control know this is the field that needs updating for setting the synonmy group name
        self.resultsList.SetSynonymGroupCtrl(self.synonymGroup)
        addSynonymSizer.Add(self.synonymGroup, 1, wx.EXPAND | wx.RIGHT, 10)
        # Add a button for a hidden group of words, the "Do Not Show" group.
        self.addNoShowBtn = wx.Button(resultsPanel, -1, _('Add Checked Items to the "Do Not Show" Group'))
        addSynonymSizer.Add(self.addNoShowBtn, 0, wx.EXPAND | wx.RIGHT, 10)
        self.addNoShowBtn.Bind(wx.EVT_BUTTON, self.OnSetSynonyms)
        # Add button for editing the selected Synonym Group
        editSynonymsBtn = wx.Button(resultsPanel, -1, _('Edit Synonyms'))
        addSynonymSizer.Add(editSynonymsBtn, 0, wx.EXPAND)
        # Add the Synonyms Controls' Sizer to the notebook Panel's sizer
        pnl1Sizer.Add(addSynonymSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # Add a page to the Notebook control for the Results tab
        self.notebook.AddPage(resultsPanel, _("Results"))
        

##        # Add a Panel for a Word Cloud output tab
##        self.wordCloudPanel = wx.Panel(self.notebook, -1)
##        self.wordCloudPanel.SetBackgroundColour(wx.BLUE)
##        
##        # Add a page to the Notebook control for the Word Cloud tab
##        self.notebook.AddPage(self.wordCloudPanel, _("Word Cloud"))
        

##        # Add a Panel for an Options tab
##        self.optionsPanel = wx.Panel(self.notebook, -1)
##        self.optionsPanel.SetBackgroundColour(wx.RED)

        # Minimum Word Frequency

        # Minimum Word Length
##        
##        # Add a page to the Notebook control for the Options tab
##        self.notebook.AddPage(self.optionsPanel, _("Options"))
        
        # Add the Notebook control to the form's Main Sizer
        mainSizer.Add(self.notebook, 1, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 10)

        # Add a Horizontal Sizer for the form buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer on the left to expand, allowing the buttons to be right-justified
        btnSizer.Add((10, 1), 1)
        # Add an OK button and a Help button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        btnOK.Bind(wx.EVT_BUTTON, self.OnOK)
        btnHelp = wx.Button(self, -1, _("Help"))

        # Put the buttons in the  Button sizer
        btnSizer.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        btnSizer.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        # Add the Button Size to the Dialog Sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)

        # Add in the Column Sorter mixin to make the results panel sortable
        ListCtrlMixins.ColumnSorterMixin.__init__(self, 3)

        # We need to populate the Synonyms BEFORE we Populate the Word Frequencies!!
        self.PopulateSynonyms()

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

        # Display the form
        self.Show(True)

    def PopulateSynonyms(self):
        """ Set up the Synonyms data structures and populate them initially """
        # Initialize a dictionary for holding the Synonym Groups (key = title, value = list of synonmyms)
        self.synonyms = {}
        # Initialize a Synonym Lookup table (key = synonyms, value = Synonym Group name)
        self.synonymLookups = {}

        # If we're testing ...
        if self.tree == None or self.startNode == None:

            self.synonyms = {'a and the'    : ['a', 'and', 'the'],
                             'name'  : ['ellen', 'feiss'],
                             'pronouns'      : ['i', "i'm", 'my']}

        else:

            print "WordFrequencyReport.PopulateSynonyms():  Load synonyms from database"

            self.synonyms = DBInterface.GetSynonyms()

            

        # Now iterate though the items in the Synonym Groups
        for key in self.synonyms.keys():
            # For each synonym in the group ...
            for synonym in self.synonyms[key]:
                # ... add an entry to the Synonym Lookup table
                self.synonymLookups[synonym] = key


    def PopulateWordFrequencies(self):
        """ Clear and Populate the Word Frequency Results """
        # Clear the Control
        self.resultsList.ClearAll()
        # Add Column Heading
        self.resultsList.InsertColumn(0, _("Text"))
        self.resultsList.InsertColumn(1, _("Count"), wx.LIST_FORMAT_RIGHT)
        self.resultsList.InsertColumn(2, _("Synonyms"))

        # If we haven't read the data yet, or need to refresh it ...
        if len(self.itemDataMap) == 0:
            # ... initialize a data structure
            data = {}

            # If we're testing ...
            if self.tree == None or self.startNode == None:
                # ... define sample text
                sampleText = u""" I was writing a paper on the PC,
and it was like   beep   beep   (beep)   beep ....  beep ???  beep   (beep.)
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

##        print
##        print "WordFrequencyReport.PopulateWordFrequencies():"
        
        # Convert the extracted data to the form needed for the ColumnSorterMixin
        for key, data in itemData.items():

##            print key, data
            
            if data[0] != "Do Not Show Group":
                # Create a new row in the Results List, adding the word or Synonym Group text
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

##        print '---------------------------------------------------------------'
##        print

        # The ColumnSorterMixin requires this data structure with this variable name to function
        self.itemDataMap = itemData
        # This code comes here, as width is calculated immediately, so must be AFTER data population.        
        self.resultsList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        self.resultsList.SetColumnWidth(1, 80)
        self.resultsList.SetColumnWidth(2, wx.LIST_AUTOSIZE)

        # If the Text Column is too narrow to show the column header ...
        if self.resultsList.GetColumnWidth(0) < 200:
            # ... widen the column
            self.resultsList.SetColumnWidth(0, 200)
        # If the Synonyms Column is too narrow to show the column header ...
        if self.resultsList.GetColumnWidth(2) < 200:
            # ... widen the column
            self.resultsList.SetColumnWidth(2, 200)

        # Sort data by Count, Descending
        self.SortListItems(self.sortColumn, self.sortAscending)
        # Update the Control to try to fix the first column header's appearance
        self.resultsList.Update()

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
        # Initialize the data dictionary
        data = {}
        # Extract the data from the tree using this recursive method
        data = self.ExtractDataFromNode(tree, startNode, data)
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

##        print text.encode('utf8')

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

#            print "CountWords:", type(word), type(self.synonymLookups.keys()[0])
            
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

    def OnOK(self, event):
        """ Handle the OK button """
        # Close the form
        self.Close()

    def OnSetSynonyms(self, event):
        """ Handle the Set Synonym Group buttons """
        
        if event.GetId() == self.addNoShowBtn.GetId():
            synonymGroup = 'Do Not Show Group'
            usingAlias = True
        else:
            synonymGroup = self.synonymGroup.GetValue()
            if self.synonyms.has_key(synonymGroup):
                usingAlias = not (synonymGroup in self.synonyms[synonymGroup])
            else:
                usingAlias = True
                count = self.resultsList.GetItemCount()
                for itemId in range(count):
                    item = self.resultsList.GetItemData(itemId)
                    
##                    print self.itemDataMap[item], synonymGroup, self.resultsList.IsChecked(itemId)

                    if (self.itemDataMap[item][0] == synonymGroup) and (self.resultsList.IsChecked(itemId)):
                        usingAlias = False
                        break

##        print "WordFrequencyReport.OnSetSynonyms():", synonymGroup, usingAlias
##        if self.synonyms.has_key(synonymGroup):
##            print self.synonyms[synonymGroup]
##        print
        
        synonymData = []
        synonymString = ''

        if synonymGroup == '':

            print "ERROR - No Synonym Group Defined"

            self.synonymGroup.SetFocus()
            return

        # if the Synonym Group already exists ...
        if synonymGroup in self.synonyms.keys():
            # Get the current synonyms as a starting point
            synonymData = self.synonyms[synonymGroup]

        # Iterate through the Results List
        count = self.resultsList.GetItemCount()
        oldSynonymGroup = ''
        for itemId in range(count):
            item = self.resultsList.GetItemData(itemId)
            
#            print self.itemDataMap[item],

            # If the item is checked and not already in the list ...

            #not((self.itemDataMap[item][0] == synonymGroup): or

            # Synonym Group we're adding to, whether it's checked or not! OR
            # Any checked item

            if (self.itemDataMap[item][0] == synonymGroup) or \
               (self.resultsList.IsChecked(itemId)):


                if (self.itemDataMap[item][0] == synonymGroup) and \
                   (synonymGroup in self.synonyms.keys()):

                    pass

                elif self.itemDataMap[item][2] != '':

                    oldSynonymGroup = self.itemDataMap[item][0]
                    
                    # ... add it to the synonym DATA
                    synonymData += self.itemDataMap[item][2]

##                    print ' XXX', self.itemDataMap[item][2], "added"

                    for tmpItem in self.itemDataMap[item][2]:
#                        synonymData.append(tmpItem)
                        self.synonymLookups[tmpItem] = synonymGroup

                        DBInterface.UpdateSynonym(oldSynonymGroup, tmpItem, synonymGroup, tmpItem)

                else:
                    # ... add it to the synonym DATA
                    synonymData.append(self.itemDataMap[item][0])

##                    print "added"

                    self.synonymLookups[self.itemDataMap[item][0]] = synonymGroup

                    DBInterface.AddSynonym(synonymGroup, self.itemDataMap[item][0]) 

##                print "synonymLookup:", self.itemDataMap[item][0].encode('utf8'), synonymGroup.encode('utf8')
                
##            else:
##                print

        # If the user pressed the Do Not Show button without checking any items ...
        if (synonymGroup == 'Do Not Show Group') and (len(synonymData) == 0):
            # ... there's nothing to do!

            print "EXIT"
            
            return

        # Sort the Synonym Data
        synonymData.sort()
        # Create a String version of the list for display
        for synonym in synonymData:
            if len(synonymString) > 0:
                synonymString += ' '
            synonymString += synonym

        # Update the permanent Synonyms data
        self.synonyms[synonymGroup] = synonymData

        index = -1
        # Now iterate though the items in the Synonyms List (control)
        for val in range(self.resultsList.GetItemCount()):
            
##            print val, self.resultsList.GetItem(val, 0).GetText().encode('utf8'), synonymGroup, self.resultsList.GetItem(val, 0).GetText() == synonymGroup

            # If the Synonym Group is found, we can update it
            if self.resultsList.GetItem(val, 0).GetText() == synonymGroup:
                index = val
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
        for key in self.itemDataMap.keys():

##            print self.itemDataMap[key][0].encode('utf8'), synonymGroup.encode('utf8'), oldSynonymGroup.encode('utf8')
            
            # Here's the trick.  The Synonyms Group name may not be in the sysnonymData list!!
            if (self.itemDataMap[key][0] in synonymData) or (self.itemDataMap[key][0] == synonymGroup) or \
               (self.itemDataMap[key][0] == oldSynonymGroup):
                itemValue += self.itemDataMap[key][1]
                if self.itemDataMap[key][0] == synonymGroup:
                    itemIndex = key
                else:
                    del(self.itemDataMap[key])
        if itemIndex == -1:
            self.itemDataMap[len(self.itemDataMap)] = (synonymGroup, itemValue, synonymData)
        else:
            self.itemDataMap[itemIndex] = (synonymGroup, itemValue, synonymData)

        scrollPos = self.resultsList.GetScrollPos(wx.VERTICAL)

        self.PopulateWordFrequencies()

        self.synonymGroup.SetValue('')


        print "ScrollPos =", scrollPos

        self.resultsList.Freeze()        
        self.resultsList.ScrollLines(scrollPos)
        self.resultsList.Thaw()

##        print
##        print self.synonymLookups
##        print
##        print


if __name__ == '__main__':
    class MyApp(wx.App):
       def OnInit(self):
          frame = WordFrequencyReport(None, None, None)
          self.SetTopWindow(frame)
          return True
          

    app = MyApp(0)
    app.MainLoop()
