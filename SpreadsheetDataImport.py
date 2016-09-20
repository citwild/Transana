# Copyright (C) 2002-2016 Spurgeon Woods LLC
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of version 2 of the GNU General Public License as
# published by the Free Software Foundation.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
#

""" This module implements the Spreadsheet Data Import function for Transana """

__author__ = 'David Woods <dwoods@transana.com>'

# Import wxPython
import wx
# Import wxPython's Wizard
import wx.wizard as wiz

# Import Transana's Document Object
import Document
# Import Transana's Keyword Object
import KeywordObject
# Import Transana's Library Object
import Library
# Import Transana's Constants
import TransanaConstants
# Import Transana's Exceptions
import TransanaExceptions
# Import Transana's Global Variables
import TransanaGlobal

# import Python's csv module, which reads Comma Separated Values files
import csv
# import Python's datetime module
import datetime
# Import Python's os module
import os


class EditBoxFileDropTarget(wx.FileDropTarget):
    """ This simple derived class let's the user drop files onto an edit box """
    def __init__(self, editbox):
        """ Initialize a File Drop Target """
        # Initialize the FileDropTarget object
        wx.FileDropTarget.__init__(self)
        # Make the Edit Box passed in a File Drop Target
        self.editbox = editbox
        
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        # Insert the FIRST file name into the edit box
        self.editbox.SetValue(files[0])


class WizPage(wiz.PyWizardPage):
    """ Base class for individual wizard pages.  Provides:
          Title
          Back / Forward / Cancel button """
    def __init__(self, parent, title):
        """ Initialize the Wizard Page """
        # Initialize the Previous and Next pointer to None
        self.prev = self.next = None
        # Remember the parent
        self.parent = parent

        # Initialize the PyWizardPage object
        wiz.PyWizardPage.__init__(self, parent)
        # Define a main Sizer for the page
        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # Display the Page Title
        title = wx.StaticText(self, -1, title)
        title.SetFont(wx.Font(14, wx.SWISS, wx.NORMAL, wx.BOLD))
        self.sizer.Add(title, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        self.sizer.Add(wx.StaticLine(self, -1), 0, wx.EXPAND|wx.ALL, 5)

        # Set the main Sizer
        self.SetSizer(self.sizer)

        # Identify the Previous and Next buttons so they can be manipulated
        self.prevButton = parent.FindWindowById(wx.ID_BACKWARD)
        self.nextButton = parent.FindWindowById(wx.ID_FORWARD)

    def GetPrev(self):
        """ Get Previous Page function """
        return self.prev

    def SetPrev(self, prev):
        """ Set Previous Page function """
        self.prev = prev

    def GetNext(self):
        """ Get Next Page function """
        return self.next

    def SetNext(self, next):
        """ Set Next Page function """
        self.next = next

    def IsComplete(self):
        """ Has this page been completed?  Defaults to False, must be over-ridden! """
        return False


class GetFileNamePage(WizPage):
    """ Get File Name wizard page """
    def __init__(self, parent, title):
        """ Define the Wizard Page that gets the File to be imported """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add the Source File label
        lblSource = wx.StaticText(self, -1, _("Source Data File:"))
        self.sizer.Add(lblSource, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)

        # Create the box1 sizer, which will hold the source file and its browse button
        box1 = wx.BoxSizer(wx.HORIZONTAL)

        # Create the Source File text box
        self.txtSrcFileName = wx.TextCtrl(self, -1)
        # Make the Source File a File Drop Target
        self.txtSrcFileName.SetDropTarget(EditBoxFileDropTarget(self.txtSrcFileName))

        # Handle ALL changes to the source filename
        self.txtSrcFileName.Bind(wx.EVT_TEXT, self.OnSrcFileNameChange)
        # Add the text box to the sizer
        box1.Add(self.txtSrcFileName, 1, wx.EXPAND)
        # Spacer
        box1.Add((4, 0))
        # Create the Source File Browse button
        self.srcBrowse = wx.Button(self, -1, _("Browse"))
        self.srcBrowse.Bind(wx.EVT_BUTTON, self.OnBrowse)
        box1.Add(self.srcBrowse, 0)
        # Add the Source Sizer to the Main Sizer
        self.sizer.Add(box1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

    def IsComplete(self):
        """ IsComplete signals whether an EXISTING file has been selected """
        return os.path.exists(self.txtSrcFileName.GetValue())

    def OnBrowse(self, event):
        """ Browse Button event handler """
        # Get Transana's File Filter definitions
        fileTypesString = _("All supported files (*.csv, *.txt)|*.csv;*.txt|Comma Separated Values files (*.csv)|*.csv|Tab Delimited Text files (*.txt)|*.txt|All files (*.*)|*.*")
        # Create a File Open dialog.
        fs = wx.FileDialog(self, _('Select a spreadsheet data file:'),
                        TransanaGlobal.configData.videoPath,
                        "",
                        fileTypesString, 
                        wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        # Select "All supported files" as the initial Filter
        fs.SetFilterIndex(0)
        # Show the dialog and get user response.  If OK ...
        if fs.ShowModal() == wx.ID_OK:
            # ... get the selected file name
            self.fileName = fs.GetPath()
        # If not OK ...
        else:
            # ... signal Cancel by return a blank file name
            self.fileName = ''

        # Destroy the File Dialog
        fs.Destroy()

        # Add the selected File Name to the File Name text box
        self.txtSrcFileName.SetValue(self.fileName)

    def OnSrcFileNameChange(self, event):
        """ Process changes to the File Name Text Box """
        # If we have a valid, existing file ...
        if self.IsComplete():
            # ... enable the Next button
            self.nextButton.Enable(True)
        # If we don't have a valid file ...
        else:
            # ... disable the Next button
            self.nextButton.Enable(False)


class GetRowsOrColumnsPage(WizPage):
    """ Wizard page to find out if the data is organized by Rows or by Columns """
    def __init__(self, parent, title):
        """ Define the Wizard Page that gets the File's Data Orientation """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add a Text Field for the File Name
        self.fileName = wx.StaticText(self, -1, '')
        self.sizer.Add(self.fileName, 0, wx.ALL, 5)
        self.sizer.Add((1, 5))

        # Add a Hozizontal Box Sizer
        boxH = wx.BoxSizer(wx.HORIZONTAL)
        # A Vertical Box Sizer 1 (left)
        boxV1 = wx.BoxSizer(wx.VERTICAL)
        # Add direction prompt
        prompt1 = wx.StaticText(self, -1, _('Contents of the first Column'))
        boxV1.Add(prompt1, 0)
        # Add a multi-line TextCtrl to hold first Column items
        self.txt1 = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE)
        boxV1.Add(self.txt1, 1, wx.EXPAND | wx.ALL, 3)
        # Add a checkbox for selecting COLUMNS
        self.chkColumns = wx.CheckBox(self, -1, " " + _("Prompts shown in first Column"), style=wx.CHK_2STATE)
        boxV1.Add(self.chkColumns, 0)

        # A Vertical Box Sizer 2 (right)
        boxV2 = wx.BoxSizer(wx.VERTICAL)
        # Add direction prompt
        prompt2 = wx.StaticText(self, -1, _('Contents of the first Row'))
        boxV2.Add(prompt2, 0)
        # Add a multi-line TextCtrl to hold first Row items
        self.txt2 = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE)
        boxV2.Add(self.txt2, 1, wx.EXPAND | wx.ALL, 3)
        # Add a checkbox for selecting ROWS
        self.chkRows = wx.CheckBox(self, -1, " " + _("Prompts shown in first Row"), style=wx.CHK_2STATE)
        boxV2.Add(self.chkRows, 0)

        # Add the Vertical Sizers to the Horizontal Sizer
        boxH.Add(boxV1, 1, wx.EXPAND)
        boxH.Add(boxV2, 1, wx.EXPAND)
        # Add the Horizontal Sizer to the Main Sizer
        self.sizer.Add(boxH, 1, wx.EXPAND)

        # Set the processor for CheckBoxes
        self.Bind(wx.EVT_CHECKBOX, self.OnCheckbox)

    def IsComplete(self):
        """ IsComplete signals whether either Checkbox has been checked """
        return self.chkColumns.GetValue() or self.chkRows.GetValue()

    def OnCheckbox(self, event):
        """ Process Checkbox Activity """
        # If the Columns Checkbox is clicked ...
        if event.GetId() == self.chkColumns.GetId():
            # ... and has been CHECKED ...
            if self.chkColumns.GetValue():
                # ... un-check the Rows checkbox ...
                self.chkRows.SetValue(False)
                # ... and enable the Next Button
                self.nextButton.Enable(True)
            # ... and has been UN-CHECKED ...
            else:
                # ... disable the Next Button
                self.nextButton.Enable(False)
        # If the Rows Checkbox is clicked ...
        else:
            # ... and has been CHECKED ...
            if self.chkRows.GetValue():
                # ... un-check the Columns checkbox ...
                self.chkColumns.SetValue(False)
                # ... and enable the Next Button
                self.nextButton.Enable(True)
            # ... and has been UN-CHECKED ...
            else:
                # ... disable the Next Button
                self.nextButton.Enable(False)


class GetItemsToIncludePage(WizPage):
    """ Wizard page to find out which items to include and how to present them """
    def __init__(self, parent, title):
        """ Define the Wizard Page that information about organizing and including data """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add the first Instruction Text
        instructions1 = _("Please select the item used to uniquely identify Participants")
        txtInstructions1 = wx.StaticText(self, -1, instructions1)
        self.sizer.Add(txtInstructions1, 0, wx.ALL, 5)

        # Add a Choice Box for Unique Identifier
        self.identifier = wx.Choice(self, -1)
        self.sizer.Add(self.identifier, 0, wx.EXPAND | wx.ALL, 5)

        # Add a spacer
        self.sizer.Add((1, 5))

        # Add the second Instruction Text
        instructions2 = _("Please select which items to include in the Transana Documents to be created.")
        txtInstructions2 = wx.StaticText(self, -1, instructions2)
        self.sizer.Add(txtInstructions2, 0, wx.ALL, 5)

        # Add a ListBox for the Prompts / Questions with multi-select enabled
        self.questions = wx.ListBox(self, -1, style = wx.LB_EXTENDED)
        self.sizer.Add(self.questions, 1, wx.EXPAND | wx.ALL, 5)
        # Bind a handler for Item Selection
        self.Bind(wx.EVT_LISTBOX, self.OnItemSelect)

        # Add a spacer
        self.sizer.Add((1, 5))

        # Add a Radio Box to specify organization by Participant or by Question
        self.organize = wx.RadioBox(self, -1, label = _('Create Documents to contain data for ...'),
                                    choices = [_('Participants (allows auto-coding)'), _('Questions (no auto-coding)')])
        self.sizer.Add(self.organize, 0, wx.EXPAND | wx.ALL, 5)
        # Bind a handler for Radio Box Selection
        self.organize.Bind(wx.EVT_RADIOBOX, self.OnOrganizationSelect)

    def IsComplete(self):
        """ IsComplete signals whether ANY questions have been selected """
        return len(self.questions.GetSelections()) > 0

    def OnItemSelect(self, event):
        """ Handler for Question / Prompt Selection """
        # Check Next Button Enabling on each change
        self.nextButton.Enable(self.IsComplete())

    def OnOrganizationSelect(self, event):
        """ Handler for Organization Radio Box """
        # If we are organizing by Participant ...
        if self.organize.GetSelection() == 0:
            # ... then we should go to Page 4 after this page ...
            self.SetNext(self.parent.AutoCodePage)
            # ... and the Next Button should be "Next"
            self.nextButton.SetLabel(_('Next >'))
        # If we are organizing by Question ...
        else:
            # ... then we should SKIP Page 4 and are done with the Wizard ...
            self.SetNext(None)
            # ... and the Next button should be labelled "Finish".
            self.nextButton.SetLabel(_('Finish'))

    
class GetAutoCodePage(WizPage):
    """ Wizard page to find out which items to auto-code """
    def __init__(self, parent, title):
        """ Define the Wizard Page that gets the auto-code information """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add the first Instruction Text
        instructions3 = _("Please select which items to auto-code at the Document level.")
        txtInstructions3 = wx.StaticText(self, -1, instructions3)
        self.sizer.Add(txtInstructions3, 0, wx.ALL, 5)

        # Add a ListBox that allows selection of what items to auto-code, with multi-select enabled.
        self.autocode = wx.ListBox(self, -1, style = wx.LB_EXTENDED)
        self.sizer.Add(self.autocode, 1, wx.EXPAND | wx.ALL, 5)
        # Add a handler for auto-code item selection
        self.Bind(wx.EVT_LISTBOX, self.OnItemSelect)

    def IsComplete(self):
        """ Selections are NOT required on this page, so it's always complete. """
        return True

    def OnItemSelect(self, event):
        """ Handler for Auto-Code Item Selection """
        # Check Next Button Enabling on each change
        self.nextButton.Enable(self.IsComplete())


class SpreadsheetDataImport(wiz.Wizard):
    """ This displays the main Spreadsheet Data Import Wizard window. """
    def __init__(self, parent, treeCtrl):
        """ Initialize the Spreadsheet Data Import Wizard """
        # Remember the TreeCtrl
        self.treeCtrl = treeCtrl

        # Initialize data for the Wizard
        # This list holds the data imported from the file in a list of lists
        self.all_data = []
        # This list holds the Questions / Prompts from the data file
        self.questions = []
        # This dictionary holds Keyword Groups (keys) and Keywords (data) for Auto-Codes
        self.all_codes = {}

        # Create the Wizard
        wiz.Wizard.__init__(self, parent, -1, _('Import Spreadsheet Data'))
        self.SetPageSize(wx.Size(600, 450))

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Define the individual Wizard Pages
        self.FileNamePage = GetFileNamePage(self, _("Select a Spreadsheet Data File"))
        self.RowsOrColumnsPage = GetRowsOrColumnsPage(self, _("Identify Data Orientation"))
        self.ItemsToIncludePage = GetItemsToIncludePage(self, _("Organize Data for Import"))
        self.AutoCodePage = GetAutoCodePage(self, _("Select Items for Auto-Coding."))

        # Define the page order / relationships
        self.FileNamePage.SetNext(self.RowsOrColumnsPage)
        self.RowsOrColumnsPage.SetPrev(self.FileNamePage)
        self.RowsOrColumnsPage.SetNext(self.ItemsToIncludePage)
        self.ItemsToIncludePage.SetPrev(self.RowsOrColumnsPage)
        self.ItemsToIncludePage.SetNext(self.AutoCodePage)
        self.AutoCodePage.SetPrev(self.ItemsToIncludePage)

        # Bind Wizard Events
        self.Bind(wiz.EVT_WIZARD_PAGE_CHANGED, self.OnPageChanged)
        self.Bind(wiz.EVT_WIZARD_FINISHED, self.OnWizardFinished)

        # We need to add a HELP button to the Wizard Object.
        # First, let's create a Help Button
        self.helpButton = wx.Button(self, -1, _("Help"))
        self.helpButton.Bind(wx.EVT_BUTTON, self.OnHelp)

        # We need to figure out where to insert it into the existing Wizard Infrastructure.  It should go into the LAST
        # sizer we find in the Wizard object.  Let's initialize a variable so we can seek that sizer out.
        lastSizer = None
        # Sizers can be nested inside sizers.  So we iterate through the Wizard's Sizer's Children's sizer's Children to find it!
        for x in self.GetSizer().GetChildren()[0].Sizer.GetChildren():
            # If the current child is a sizer ...
            if x.IsSizer():
                # ... it's a candidate for the LAST sizer.  Remember it.
                lastSizer = x.Sizer
        # Add the Help button to the last sizer found!!
        if lastSizer != None:
            lastSizer.Add(self.helpButton, 0, wx.ALL, 5)

        # Run the Wixard
        self.FitToPage(self.FileNamePage)
        self.RunWizard(self.FileNamePage)

    def strip_quotes(self, text):
        """ Process text items from a Streadsheet Data File to strip leading and trailing quotes """
        # Remove leading and trailing whitespace
        text = text.strip()
        # Replace triple quotes with single quotes
        text = text.replace('"""', '"')
        # Replace double quotes with single quotes
        text = text.replace('""', '"')
        # If we have text that starts and ends with quotes ...
        if (len(text) > 1) and (text[0] == '"') and (text[-1] == '"'):
            # Strip a quote from the first and last positions in the string
            text = text[1:]
            text = text[:-1]
        # Return the processed string
        return text

    def OnPageChanged(self, event):
        """ Process Wizard Page changes """
        # If we move BACKWARDS to the File Name Page ...
        if event.GetPage() == self.FileNamePage:
            # Reset the Questions and Codes
            self.questions = []
            self.all_codes = {}

        # If we move to the Rows or Columns Page ...
        elif event.GetPage() == self.RowsOrColumnsPage:
            # ... set the File Name to the file selected on the File Name Page
            self.RowsOrColumnsPage.fileName.SetLabel(_('File:  %s') % self.FileNamePage.txtSrcFileName.GetValue())
            # If we're moving FORWARD ...
            if event.GetDirection():
                # Initialize the File Data list
                self.all_data = []
                # Note the filename
                filename = self.FileNamePage.txtSrcFileName.GetValue()
                # Open the file
                with open(filename, 'r') as f:
                    # Use the csv Sniffer to determine the dialect of the file
                    dialect = csv.Sniffer().sniff(f.read(1024))
                    # Reset the file to the beginning
                    f.seek(0)
                    # use the csv Reader to read the data file
                    csvReader = csv.reader(f, dialect=dialect)
                    # For each row of data read ...
                    for row in csvReader:
                        # ... add that row to the data list
                        self.all_data.append(row)

                # Place the first item in each row (that is, the first COLUMN of data) in the Column TextCtrl
                self.RowsOrColumnsPage.txt1.Clear()
                for row in self.all_data:
                    self.RowsOrColumnsPage.txt1.AppendText(self.strip_quotes(row[0]) + '\n')
                self.RowsOrColumnsPage.txt2.Clear()
                # Place the items from the first row (that is,the first ROW of data) in the Row TextCtrl
                for col in self.all_data[0]:
                    self.RowsOrColumnsPage.txt2.AppendText(self.strip_quotes(col) + '\n')
                # Disable the Column and Row TextCtrls
                self.RowsOrColumnsPage.txt1.Enable(False)
                self.RowsOrColumnsPage.txt2.Enable(False)

        # If we move to the Organize and Include Items page ...
        elif event.GetPage() == self.ItemsToIncludePage:
            # If we're moving FORWARD ...
            if event.GetDirection():
                # Initialize the Questions list
                self.questions = []
                # Determine the Questions if the user selected Columns ...
                if self.RowsOrColumnsPage.chkColumns.GetValue():
                    # ... by iterating through each row of data ...
                    for row in self.all_data:
                        # ... and selecting the row's first item
                        self.questions.append(self.strip_quotes(row[0]))
                # Determine the Questoins if the user selected Rows ...
                else:
                    # ... by iterating through the first row of data
                    for col in self.all_data[0]:
                        # ... and selecting the column's header
                        self.questions.append(self.strip_quotes(col))

                # Populate the combo of Questions / Prompts used to select the Unique Identifier after adding an automatic creation option
                self.ItemsToIncludePage.identifier.SetItems([_('Create one automatically')] + self.questions)
                self.ItemsToIncludePage.identifier.SetSelection(0)
                # Populate the list of Questions / Prompts
                self.ItemsToIncludePage.questions.SetItems(self.questions)

        # If we move to the Auto-Code Page ...
        elif event.GetPage() == self.AutoCodePage:
            # If we're moving FORWARD ...
            if event.GetDirection():
                # ... set the Auto-Code options to match the Questions
                self.AutoCodePage.autocode.SetItems(self.questions)

        # Identify the Next button
        nextButton = self.FindWindowById(wx.ID_FORWARD)
        # Enable (or not) the Next Button depending the new page's "completeness"
        nextButton.Enable(event.GetPage().IsComplete())

    def OnWizardFinished(self, event):
        """ Process the data when the Wizard is completed """

        # The Wizard is triggered on one of the Tree's Library nodes.  Get the Record Number for this node.
        libraryNumber = self.treeCtrl.GetPyData(self.treeCtrl.GetSelections()[0]).recNum
        # Get the full Library record
        library = Library.Library(libraryNumber)

        # Get the Unique Identifier selected by the user
        id_col = self.ItemsToIncludePage.identifier.GetSelection() - 1

        # Initialize the Participant Counter to 1 (the first participant) rather than 0.
        participantCount = 1

        # We need to move through the data differently if the source file is organized by Columns or by Rows.
        # We need to present data differently if we're organizing output data by Participant or by Question.
        # It might be possible to do this more efficiently, but I don't have time to abstract that right now.
        # As a result, there's a lot of duplication in the following code.

        # If source data Questions / Prompts are in the first Column ...
        if self.RowsOrColumnsPage.chkColumns.GetValue():
            # ... and if we're organizing output by Participant ...
            if self.ItemsToIncludePage.organize.GetSelection() == 0:
                # ... initialize the auto-codes found for THIS participant
                codes = {}

                # For each COLUMN ...
                for x in range(1, len(self.all_data[0])):

                    # If the user requested automatic unique Participant IDs ...
                    if id_col == -1:
                        # ... create a unique Participant ID and increment the Participant Counter
                        participantID = _('Participant %04d') % participantCount
                        participantCount += 1
                    # Otherwise, use the data the user requested
                    else:
                        participantID = self.strip_quotes(self.all_data[id_col][x])

                    # Create Document by participantID
                    tmpDoc = Document.Document()
                    # Populate essential Document Properties
                    tmpDoc.id = participantID
                    tmpDoc.library_num = libraryNumber
                    tmpDoc.imported_file = self.FileNamePage.txtSrcFileName.GetValue()
                    tmpDoc.import_date = datetime.datetime.now().strftime('%Y-%m-%d')
                    # Initialize Document Text and PlainText
                    tmpDoc.text = 'txt\n'
                    tmpDoc.plaintext = ''

                    # For each Question that should be included in the output ...
                    for q in self.ItemsToIncludePage.questions.GetSelections():
                        # ... populate the Document Text and Plain Text with Question and Response
                        tmpDoc.text += '%s\n%s\n\n' % (self.strip_quotes(self.questions[q]), self.strip_quotes(self.all_data[q][x]))
                        tmpDoc.plaintext += '%s\n%s\n\n' % (self.strip_quotes(self.questions[q]), self.strip_quotes(self.all_data[q][x]))

                    # Remove trailing carriage returns
                    tmpDoc.text = tmpDoc.text[:-2]
                    tmpDoc.plaintext = tmpDoc.plaintext[:-2]

                    # For each selected Auto-Code category ...
                    for c in self.AutoCodePage.autocode.GetSelections():
                        # Define the Keyword Group
                        kwg = _('Auto-code')
                        # Define the Keyword
                        kw = "%s : %s" % (self.strip_quotes(self.questions[c]), self.strip_quotes(self.all_data[c][x]))
                        # Replace Parentheses (illegal in Keywords) with Brackets
                        kw = kw.replace('(', '[')
                        kw = kw.replace(')', ']')
                        
                        # If there was no missing data in the Keyword Definition ...
                        if (self.strip_quotes(self.questions[c]) != '') and (self.strip_quotes(self.all_data[c][x]) != ''):
                            # ... Add the Keyword to the Document
                            tmpDoc.add_keyword(kwg, kw)
                            # If the Keyword Group had not been defined ...
                            if not kwg in self.all_codes.keys():
                                # ... define the Keyword Group using a list containing the Keyword
                                self.all_codes[kwg] = [kw]
                                # Try to load the keyword to see if it already exists.
                                try:
                                    keyword = KeywordObject.Keyword(kwg, kw)
                                # If the Keyword doesn't exist yet ...
                                except TransanaExceptions.RecordNotFoundError:
                                    # ... create the Keyword.
                                    keyword = KeywordObject.Keyword()
                                    keyword.keywordGroup = kwg
                                    keyword.keyword = kw
                                    keyword.definition = _('Created during Spreadsheet Data Import for file "%s."') % self.FileNamePage.txtSrcFileName.GetValue()
                                    # Try to save the Keyword
                                    keyword.db_save()
                                    # Add the new Keyword to the database tree
                                    self.treeCtrl.add_Node('KeywordNode', (_('Keywords'), keyword.keywordGroup, keyword.keyword), 0, keyword.keywordGroup)

                                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                                    if not TransanaConstants.singleUserVersion:
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (keyword.keywordGroup, keyword.keyword))
                            # If the Keyword Group HAS been defined ...
                            else:
                                # ... if the Keyword has NOT been defined ...
                                if not kw in self.all_codes[kwg]:
                                    # ... add the Keyword to the Keyword Group's Keyword List
                                    self.all_codes[kwg].append(kw)
                                    # Try to load the keyword to see if it already exists.
                                    try:
                                        keyword = KeywordObject.Keyword(kwg, kw)
                                    # If the Keyword doesn't exist yet ...
                                    except TransanaExceptions.RecordNotFoundError:
                                        # ... create the Keyword.
                                        keyword = KeywordObject.Keyword()
                                        keyword.keywordGroup = kwg
                                        keyword.keyword = kw
                                        keyword.definition = _('Created during Spreadsheet Data Import for file "%s."') % self.FileNamePage.txtSrcFileName.GetValue()
                                        # Try to save the Keyword
                                        keyword.db_save()
                                        # Add the new Keyword to the database tree
                                        self.treeCtrl.add_Node('KeywordNode', (_('Keywords'), keyword.keywordGroup, keyword.keyword), 0, keyword.keywordGroup)

                                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                                        if not TransanaConstants.singleUserVersion:
                                            if TransanaGlobal.chatWindow != None:
                                                TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (keyword.keywordGroup, keyword.keyword))

                    # Try to save the new Document
                    try:
                        tmpDoc.db_save()
                    # If there is a SaveError (duplicate name is most likely) ...
                    except TransanaExceptions.SaveError:
                        # ... try to give it a unique identifier
                        tmpDoc.id += ' (%04d)' % participantCount
                        participantCount += 1
                        # Try to save it again
                        tmpDoc.db_save()
                    finally:
                        # Add the new Document to the Database Tree
                        nodeData = (_('Libraries'), library.id, tmpDoc.id)
                        self.treeCtrl.add_Node('DocumentNode', nodeData, tmpDoc.number, library.number)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("AD %s >|< %s" % (nodeData[-2], nodeData[-1]))

            # ... and if we're organizing output by Question ...
            elif self.ItemsToIncludePage.organize.GetSelection() == 1:
                # For each Question that should be included in the output ...
                for q in self.ItemsToIncludePage.questions.GetSelections():
                    # Note the Question
                    questionID = self.strip_quotes(self.all_data[q][0])

                    # Create Document by QuestionID
                    tmpDoc = Document.Document()
                    # Populate essential Document Properties
                    tmpDoc.id = questionID
                    tmpDoc.library_num = libraryNumber
                    tmpDoc.imported_file = self.FileNamePage.txtSrcFileName.GetValue()
                    tmpDoc.import_date = datetime.datetime.now().strftime('%Y-%m-%d')
                    # Initialize Document Text and PlainText
                    tmpDoc.text = 'txt\n'
                    tmpDoc.plaintext = ''

                    # We have to re-initialize the Participant Counter for each Question
                    participantCount = 1
                    
                    # For each COLUMN ...
                    for x in range(1, len(self.all_data[0])):

                        # If the user requested automatic unique Participant IDs ...
                        if id_col == -1:
                            # ... create a unique Participant ID and increment the Participant Counter
                            participantID = _('Participant %04d') % participantCount
                            participantCount += 1
                        # Otherwise, use the data the user requested
                        else:
                            participantID = self.strip_quotes(self.all_data[id_col][x])

                        # ... populate the Document Text and Plain Text with Participant ID and Response
                        tmpDoc.text += '%s\n%s\n\n' % (participantID, self.strip_quotes(self.all_data[q][x]))
                        tmpDoc.plaintext += '%s\n%s\n\n' % (participantID, self.strip_quotes(self.all_data[q][x]))

                    # Remove trailing carriage returns
                    tmpDoc.text = tmpDoc.text[:-2]
                    tmpDoc.plaintext = tmpDoc.plaintext[:-2]

                    # Try to save the new Document
                    try:
                        tmpDoc.db_save()
                    # If there is a SaveError (duplicate name is most likely) ...
                    except TransanaExceptions.SaveError:
                        # ... try to give it a unique identifier
                        tmpDoc.id += ' (%04d)' % participantCount
                        participantCount += 1
                        # Try to save it again
                        tmpDoc.db_save()
                    finally:
                        # Add the new Document to the Database Tree
                        nodeData = (_('Libraries'), library.id, tmpDoc.id)
                        self.treeCtrl.add_Node('DocumentNode', nodeData, tmpDoc.number, library.number)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("AD %s >|< %s" % (nodeData[-2], nodeData[-1]))

        # If source data Questions / Prompts are organized in Rows ...
        elif self.RowsOrColumnsPage.chkRows.GetValue():
            # ... and if we're organizing output by Participant ...
            if self.ItemsToIncludePage.organize.GetSelection() == 0:
                # ... initialize the auto-codes found for THIS participant
                codes = {}
                
                # For each ROW ...
                for x in range(1, len(self.all_data)):

                    # If the user requested automatic unique Participant IDs ...
                    if id_col == -1:
                        # ... create a unique Participant ID and increment the Participant Counter
                        participantID = _('Participant %04d') % participantCount
                        participantCount += 1
                    # Otherwise, use the data the user requested
                    else:
                        participantID = self.strip_quotes(self.all_data[x][id_col])
                        
                    # Create Document by participantID
                    tmpDoc = Document.Document()
                    # Populate essential Document Properties
                    tmpDoc.id = participantID
                    tmpDoc.library_num = libraryNumber
                    tmpDoc.imported_file = self.FileNamePage.txtSrcFileName.GetValue()
                    tmpDoc.import_date = datetime.datetime.now().strftime('%Y-%m-%d')
                    # Initialize Document Text and PlainText
                    tmpDoc.text = 'txt\n'
                    tmpDoc.plaintext = ''

                    # For each Question that should be included in the output ...
                    for q in self.ItemsToIncludePage.questions.GetSelections():
                        # ... populate the Document Text and Plain Text with Question and Response
                        tmpDoc.text += '%s\n%s\n\n' % (self.strip_quotes(self.questions[q]), self.strip_quotes(self.all_data[x][q]))
                        tmpDoc.plaintext += '%s\n%s\n\n' % (self.strip_quotes(self.questions[q]), self.strip_quotes(self.all_data[x][q]))

                    # Remove trailing carriage returns
                    tmpDoc.text = tmpDoc.text[:-2]
                    tmpDoc.plaintext = tmpDoc.plaintext[:-2]

                    # For each selected Auto-Code category ...
                    for c in self.AutoCodePage.autocode.GetSelections():
                        # Define the Keyword Group
                        kwg = _('Auto-code')
                        # Define the Keyword
                        kw = "%s : %s" % (self.strip_quotes(self.questions[c]), self.strip_quotes(self.all_data[x][c]))
                        # Replace Parentheses (illegal in Keywords) with Brackets
                        kw = kw.replace('(', '[')
                        kw = kw.replace(')', ']')
                        
                        # If there was no missing data in the Keyword Definition ...
                        if (self.strip_quotes(self.questions[c]) != '') and (self.strip_quotes(self.all_data[x][c]) != ''):
                            # ... Add the Keyword to the Document
                            tmpDoc.add_keyword(kwg, kw)
                            # If the Keyword Group had not been defined ...
                            if not kwg in self.all_codes.keys():
                                # ... define the Keyword Group using a list containing the Keyword
                                self.all_codes[kwg] = [kw]
                                # Try to load the keyword to see if it already exists.
                                try:
                                    keyword = KeywordObject.Keyword(kwg, kw)
                                # If the Keyword doesn't exist yet ...
                                except TransanaExceptions.RecordNotFoundError:
                                    # ... create the Keyword.
                                    keyword = KeywordObject.Keyword()
                                    keyword.keywordGroup = kwg
                                    keyword.keyword = kw
                                    keyword.definition = _('Created during Spreadsheet Data Import for file "%s."') % self.FileNamePage.txtSrcFileName.GetValue()
                                    # Try to save the keyword
                                    keyword.db_save()
                                    # Add the new Keyword to the database tree
                                    self.treeCtrl.add_Node('KeywordNode', (_('Keywords'), keyword.keywordGroup, keyword.keyword), 0, keyword.keywordGroup)

                                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                                    if not TransanaConstants.singleUserVersion:
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (keyword.keywordGroup, keyword.keyword))
                            # If the Keyword Group HAS been defined ...
                            else:
                                # ... if the Keyword has NOT been defined ...
                                if not kw in self.all_codes[kwg]:
                                    # ... add the Keyword to the Keyword Group's Keyword List
                                    self.all_codes[kwg].append(kw)
                                    # Try to load the keyword to see if it already exists.
                                    try:
                                        keyword = KeywordObject.Keyword(kwg, kw)
                                    # If the Keyword doesn't exist yet ...
                                    except TransanaExceptions.RecordNotFoundError:
                                        # ... create the Keyword.
                                        keyword = KeywordObject.Keyword()
                                        keyword.keywordGroup = kwg
                                        keyword.keyword = kw
                                        keyword.definition = _('Created during Spreadsheet Data Import for file "%s."') % self.FileNamePage.txtSrcFileName.GetValue()
                                        # Try to save the Keyword
                                        keyword.db_save()
                                        # Add the new Keyword to the database tree
                                        self.treeCtrl.add_Node('KeywordNode', (_('Keywords'), keyword.keywordGroup, keyword.keyword), 0, keyword.keywordGroup)

                                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                                        if not TransanaConstants.singleUserVersion:
                                            if TransanaGlobal.chatWindow != None:
                                                TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (keyword.keywordGroup, keyword.keyword))

                    # Try to save the new Document
                    try:
                        tmpDoc.db_save()
                    # If there is a SaveError (duplicate name is most likely) ...
                    except TransanaExceptions.SaveError:
                        # ... try to give it a unique identifier
                        tmpDoc.id += ' (%04d)' % participantCount
                        participantCount += 1
                        # Try to save it again
                        tmpDoc.db_save()
                    finally:
                        # Add the new Document to the Database Tree
                        nodeData = (_('Libraries'), library.id, tmpDoc.id)
                        self.treeCtrl.add_Node('DocumentNode', nodeData, tmpDoc.number, library.number)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("AD %s >|< %s" % (nodeData[-2], nodeData[-1]))

            # ... and if we're organizing output by Question ...
            elif self.ItemsToIncludePage.organize.GetSelection() == 1:
                # For each Question that should be included in the output ...
                for q in self.ItemsToIncludePage.questions.GetSelections():
                    # Note the Question
                    questionID = self.strip_quotes(self.all_data[0][q])

                    # Create Document by QuestionID
                    tmpDoc = Document.Document()
                    # Populate essential Document Properties
                    tmpDoc.id = questionID
                    tmpDoc.library_num = libraryNumber
                    tmpDoc.imported_file = self.FileNamePage.txtSrcFileName.GetValue()
                    tmpDoc.import_date = datetime.datetime.now().strftime('%Y-%m-%d')
                    # Initialize Document Text and PlainText
                    tmpDoc.text = 'txt\n'
                    tmpDoc.plaintext = ''
                    # We have to re-initialize the Participant Counter with each Question
                    participantCount = 1

                    # For each ROW ...
                    for x in range(1, len(self.all_data)):

                        # If the user requested automatic unique Participant IDs ...
                        if id_col == -1:
                            # ... create a unique Participant ID and increment the Participant Counter
                            participantID = _('Participant %04d') % participantCount
                            participantCount += 1
                        # Otherwise, use the data the user requested
                        else:
                            participantID = self.strip_quotes(self.all_data[x][id_col])

                        # ... populate the Document Text and Plain Text with Participant ID and Response
                        tmpDoc.text += '%s\n%s\n\n' % (participantID, self.strip_quotes(self.all_data[x][q]))
                        tmpDoc.plaintext += '%s\n%s\n\n' % (participantID, self.strip_quotes(self.all_data[x][q]))

                    # Remove trailing carriage returns
                    tmpDoc.text = tmpDoc.text[:-2]
                    tmpDoc.plaintext = tmpDoc.plaintext[:-2]

                    # Try to save the new Document
                    try:
                        tmpDoc.db_save()
                    # If there is a SaveError (duplicate name is most likely) ...
                    except TransanaExceptions.SaveError:
                        # ... try to give it a unique identifier
                        tmpDoc.id += ' (%04d)' % participantCount
                        participantCount += 1
                        # Try to save it again
                        tmpDoc.db_save()
                    finally:
                        # Add the new Document to the Database Tree
                        nodeData = (_('Libraries'), library.id, tmpDoc.id)
                        self.treeCtrl.add_Node('DocumentNode', nodeData, tmpDoc.number, library.number)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("AD %s >|< %s" % (nodeData[-2], nodeData[-1]))

        # If there are auto-codes ...
        if len(self.all_codes) > 0:
            # Now let's sort the Keywords Node
            keywordsNode = self.treeCtrl.select_Node((_('Keywords'),), 'KeywordRootNode')
            self.treeCtrl.SortChildren(keywordsNode)
            # Now let's sort the Keywords Auto-Code Node
            keywordsNode = self.treeCtrl.select_Node((_('Keywords'), _('Auto-code')), 'KeywordGroupNode')
            self.treeCtrl.SortChildren(keywordsNode)

    def OnHelp(self, evt):
        """ Method to use when the Help Button is pressed """
        # If the Menu Window is defined ...
        if TransanaGlobal.menuWindow != None:
            # ... call it's Help() method for THIS control.
            TransanaGlobal.menuWindow.ControlObject.Help('Import Spreadsheet Data')
