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
from wx.lib.mixins.listctrl import CheckListCtrlMixin



# If running stand-alone ...
if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation



# Import Transana's Clip Object
import Clip
# Import Transana's Document object
import Document
# Import Transana's Quote object
import Quote
# Import Transana's Constants
import TransanaConstants
# Import Transana's Globals
import TransanaGlobal
# Import Transana's Transcript Object
import Transcript


class CheckListCtrl(wx.ListCtrl, CheckListCtrlMixin):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_REPORT | wx.LC_SORT_ASCENDING)
        self.parent = parent
        CheckListCtrlMixin.__init__(self)
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnItemActivated)
        self.Bind(wx.EVT_LIST_COL_CLICK, self.OnColClick)
        self.OnColClick(None)

    def OnItemActivated(self, evt):
        self.ToggleItem(evt.m_itemIndex)

    # this is called by the base class when an item is checked/unchecked
    def OnCheckItem(self, index, flag):
        if len(self.data) == 0:
            self.OnColClick(None)
        data = self.GetItemData(index)
        if flag:
            what = "checked"

            print "Checked:", index, data
        else:
            what = "unchecked"

            print "Unchecked:", index, data

        print self.data[index]

    def OnColClick(self, event):
        count = self.GetItemCount()
        self.data = [self.GetItem(itemId=row, col=0).GetText() for row in xrange(count)]

##        print len(self.data)


class WordFrequencyReport(wx.Frame, ListCtrlMixins.ColumnSorterMixin):
    """ This is the main Window for the Word Frequency Reports for Transana """

    def __init__(self, parent, tree, startNode):
        self.parent = parent

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

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.notebook = wx.Notebook(self, -1)

        self.resultsPanel = wx.Panel(self.notebook, -1)
        # self.resultsPanel.SetBackgroundColour(wx.RED)

        pnlVSizer = wx.BoxSizer(wx.VERTICAL)
        
        self.resultsList = CheckListCtrl(self.resultsPanel)
        
        self.resultsList.InsertColumn(0, _("Text"))
        self.resultsList.InsertColumn(1, _("Count"), wx.LIST_FORMAT_RIGHT)

        if tree == None and startNode == None:
            # FAKE DATA
            itemData = {  1 : ('Dog', 100),
                          2 : ('Cat', 50),
                          3 : ('Elephant', 220),
                          4 : ('Human', 110) }
        else:

            itemData = {}
            counter = 0

#            print "Extract data from Tree information!"
            
            data = self.ExtractDataFromTree(tree, startNode)

#            print data
            
            # Convert the extracted data to the form needed for the ColumnSorterMixin

            for text in data.keys():
                itemData[counter] = (text, data[text])
                counter += 1

#            print
#            print itemData
            

        for key, data in itemData.items():
            index = self.resultsList.InsertStringItem(sys.maxint, data[0])
            self.resultsList.SetStringItem(index, 1, str(data[1]))
            self.resultsList.SetItemData(index, key)

##            print data[0].encode('utf8'),
##            for x in data[0]:
##                print ord(x),
##            print


        self.itemDataMap = itemData


        ListCtrlMixins.ColumnSorterMixin.__init__(self, 2)
##        self.resultsList.GetColumnSorter = self.sortColumn

        # Sort data by Count, Descending
        self.SortListItems(1, False)

        # Comes here, as width is calculated immediately, so must be after introduction of data.        
        self.resultsList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        self.resultsList.SetColumnWidth(1, 100)

        
        pnlVSizer.Add(self.resultsList, 1, wx.EXPAND | wx.ALL, 5)
        self.resultsPanel.SetSizer(pnlVSizer)

        addSynonymSizer = wx.BoxSizer(wx.HORIZONTAL)
        addSynonmyBtn = wx.Button(self.resultsPanel, -1, _("Add Checked Items to Synonym Group"))
        addSynonymSizer.Add(addSynonmyBtn, 0, wx.EXPAND | wx.RIGHT, 10)
#        tmpText = wx.StaticText(self.resultsPanel, -1, _("Synonym Group:"))
#        addSynonymSizer.Add(tmpText, 0)
        self.synonymGroup = wx.TextCtrl(self.resultsPanel, -1)
        addSynonymSizer.Add(self.synonymGroup, 1, wx.EXPAND)

        pnlVSizer.Add(addSynonymSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        self.notebook.AddPage(self.resultsPanel, _("Results"))

        self.synonymPanel = wx.Panel(self.notebook, -1)
        self.synonymPanel.SetBackgroundColour(wx.GREEN)

        self.notebook.AddPage(self.synonymPanel, _("Synonyms"))

        self.wordCloudPanel = wx.Panel(self.notebook, -1)
        self.wordCloudPanel.SetBackgroundColour(wx.BLUE)
        
        self.notebook.AddPage(self.wordCloudPanel, _("Word Cloud"))

        mainSizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 10)

        # Add a Horizontal Sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer on the left to expand, allowing the buttons to be right-justified
        btnSizer.Add((10, 1), 1)
        # Add an OK button and a Cancel button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        btnOK.Bind(wx.EVT_BUTTON, self.OnOK)
        btnHelp = wx.Button(self, -1, _("Help"))

        # Put the buttons in the  Button sizer
        btnSizer.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        btnSizer.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

        # Add the Button Size to the Dialog Sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)

        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)

        if __name__ != '__main__':
            TransanaGlobal.CenterOnPrimary(self)
        self.Show(True)

    def GetListCtrl(self):
        """ Pointer to the Results List, required for the ColumnSorterMixin """
        return self.resultsList

    def ExtractDataFromTree(self, tree, startNode):
        """ This routine takes the tree and startNode, figures out what Documents and Transcripts to load, and creates
            a data structure with all the individual words in the appropriate scope along with their counts. """
        data = {}

        data = self.ExtractDataFromNode(tree, startNode, data)

        return data

    def ExtractDataFromNode(self, tree, startNode, data):
        """ This extracts data from a node, calling subnodes recursively as needed """

        itemName = tree.GetItemText(startNode)
        itemData = tree.GetPyData(startNode)

##        print "Extract Data From", itemName, itemData.nodetype

        if itemData.nodetype in ['LibraryRootNode', 'LibraryNode', 'SearchLibraryNode', 'EpisodeNode', 'SearchEpisodeNode',
                                 'CollectionsRootNode', 'CollectionNode', 'SearchCollectionNode']:
            
            (childNode, cookieItem) = tree.GetFirstChild(startNode)
            while childNode.IsOk():
                childData = tree.GetPyData(childNode)

                if childData.nodetype in ['SnapshotNode', 'LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 
                                          'TranscriptNoteNode', 'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode']:
                    pass
                else:

                    data = self.ExtractDataFromNode(tree, childNode, data)

                (childNode, cookieItem) = tree.GetNextChild(startNode, cookieItem)
                
        elif itemData.nodetype in ['DocumentNode', 'SearchDocumentNode']:
            record = Document.Document(num=itemData.recNum)
            text = self.PrepareText(record.plaintext)
            data = self.CountWords(text, data)
        elif itemData.nodetype in ['TranscriptNode', 'SearchTranscriptNode']:
            record = Transcript.Transcript(itemData.recNum)
            text = self.PrepareText(record.plaintext)
            data = self.CountWords(text, data)
        elif itemData.nodetype in ['QuoteNode', 'SearchQuoteNode']:
            record = Quote.Quote(num=itemData.recNum)
            text = self.PrepareText(record.plaintext)
            data = self.CountWords(text, data)
        elif itemData.nodetype in ['ClipNode', 'SearchClipNode']:
            clipRecord = Clip.Clip(itemData.recNum)
            for record in clipRecord.transcripts:
                text = self.PrepareText(record.plaintext)
                data = self.CountWords(text, data)
        elif itemData.nodetype in ['LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode']:
            pass
        else:
            print "ERROR:  ", tree.GetItemText(startNode).encode('utf8'), " NOT PROCESSED.  Wrong Node Type.", itemData.nodetype

        return data

    def PrepareText(self, text):
        """ This method cleans up the messy PlainText that comes in, removing time codes, punctuation, etc. """

#        print
#        print text[:1000]
#        print


        # ************************** LOOK AT re.sub() ******************************


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

        # strip multiple periods and questions marks
        text = re.sub('[\.?:\*][\.?:\*]+', ' ', text)
        # Strip parentheses, brackets, quotation marks, slashes, and equal signs
        text = re.sub('[()\[\]"/=]', ' ', text)
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

        return text

    def CountWords(self, text, words = {}):
        """ This method takes prepared text (see above) and adds it to existing WordCount data. """
        # DO NOT initialize a dictionary to hold word counts (key = word, value = count)
        # Instead, this structure can be passed in to allow additional text to be added.
        # The first call should leave it off, or it should be initialized externally.
        # words = {}
        # For each line of the file ...
        for line in text.split('\n'):
           # ... remove whitespace and compensate for different cases
           word = line.strip().lower()
           if not word in ['', '-', ':']:
               # If the word is already in the dictionary ...
               if words.has_key(word):
                  # ... increment the count
                  words[word] += 1
               # If the word is not in the dictionary ...
               else:
                  # ... add it with a count of one.
                  words[word] = 1

        return words

    def OnOK(self, event):
        self.Close()

if __name__ == '__main__':
    class MyApp(wx.App):
       def OnInit(self):
          frame = WordFrequencyReport(None, None, None)
          self.SetTopWindow(frame)
          return True
          

    app = MyApp(0)
    app.MainLoop()
