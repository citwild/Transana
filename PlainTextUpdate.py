# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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

"""This file implements the Plain Text Update mechanism for Transana 3.1."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx

# Import Transana's Clip object
import Clip
# Import Transana's Database Interface
import DBInterface
# Import Transana's Document object
import Document
# import Transana's Quote object
import Quote
# Import Transana's Constants
import TransanaConstants
# Import Transana's global module
import TransanaGlobal
# import Transana's Transcript object
import Transcript
# import Transana's Rich Text Edit Control
import TranscriptEditor_RTC

class PlainTextUpdate(wx.Dialog):
    def __init__(self, parent, numRecords = 0):
        """ Create a Dialog Box to process the Database Conversion.
              Parameters:  parent       Parent Window
                           numRecords   The number of records that will be updated  """
        # Remember the total number of records passed in, or obtain that number if needed
        if numRecords == 0:
            self.numRecords = DBInterface.CountItemsWithoutPlainText()
        else:
            self.numRecords = numRecords

        # Create a Dialog Box
        wx.Dialog.__init__(self, parent, -1, _("Plain Text Conversion"), size=(600, 350))
        # Add the Main Sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Add a gauge based on the number of records to be handled
        self.gauge = wx.Gauge(self, -1, numRecords)
        mainSizer.Add(self.gauge, 0, wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT, 5)
        # Add a TextCtrl to provide user information
        self.txtCtrl = wx.TextCtrl(self, -1, "", style=wx.TE_LEFT | wx.TE_MULTILINE)
        mainSizer.Add(self.txtCtrl, 1, wx.EXPAND | wx.ALL, 5)
        # Add a hidden RichTextCrtl used for the conversion process
        self.richTextCtrl = TranscriptEditor_RTC.TranscriptEditor(self)
        self.richTextCtrl.SetReadOnly(True)
        self.richTextCtrl.Enable(False)
        # The user doesn't need to see this.
        self.richTextCtrl.Show(False)
        # Even though we add it to the sizer, it won't show up because of the Show(False).
        mainSizer.Add(self.richTextCtrl, 1, wx.EXPAND | wx.ALL, 5)

        # Finalize the Dialog layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        # Center the Dialog on the Screen
        TransanaGlobal.CenterOnPrimary(self)

    def OnConvert(self):
        """ Perform the Plain Text Extraction operation """
        # Get the database connection (required for Transactions)
        db = DBInterface.get_db()
        # Get a Database Cursor
        dbCursor = db.cursor()

        # Initialize a Record Counter
        counter = 0

        # Get a list of the Documents that need Plain Text extraction
        documents = DBInterface.list_of_documents(withoutPlainText=True)
        # Update User Info
        self.txtCtrl.AppendText("%5d Document Records\n" % len(documents))
        # Iterate through the list
        for document in documents:
            # Update User Info
            self.txtCtrl.AppendText("%5d  Document:  %s\n" % (self.numRecords - counter, document[1]))
            # Load the Document Object
            tmpDocument = Document.Document(num=document[0])

            # I'm not sure why I have to use a Transaction here.  But records are remaining locked
            # without this.  This at least makes things work!
            dbCursor.execute("BEGIN")
            # Lock the record
            tmpDocument.lock_record()
            # Load the record into the hidden RichTextCtrl.  Don't show the popup, as too many of these crashes the program!
            self.richTextCtrl.load_transcript(tmpDocument, showPopup=False)
            # Save the record.  This causes the PlainText to be added!  Don't show the popup, as too many of these crashes the program!
            self.richTextCtrl.save_transcript(use_transactions=False, showPopup=False)
            # Unlock the record
            tmpDocument.unlock_record()
            # Commit the Transaction
            dbCursor.execute("COMMIT")

            # Update the Record Counter
            counter += 1
            # Update the Progress Bar
            self.gauge.SetValue(counter)
            # This form can freeze up and appear non-responsive.  Every 20 items, we should avoid that
            if counter % 20 == 0:
                wx.YieldIfNeeded()

        # Get a list of the Episode Transcripts that need Plain Text extraction
        episodeTranscripts = DBInterface.list_of_episode_transcripts(withoutPlainText=True)
        # Update User Info
        self.txtCtrl.AppendText("%5d Episode Transcript Records\n" % len(episodeTranscripts))
        # Iterate through the list
        for episodeTranscript in episodeTranscripts:
            # Update User Info
            self.txtCtrl.AppendText("%5d  Episode Transcript:  %s\n" % (self.numRecords - counter, episodeTranscript[1]))
            # Load the Transcript Object
            tmpEpisodeTranscript = Transcript.Transcript(id_or_num=episodeTranscript[0])

            # I'm not sure why I have to use a Transaction here.  But the transaction from the Transcript
            # object doesn't seem to be cutting it.  This at least makes things work!
            dbCursor.execute("BEGIN")
            # Lock the record
            tmpEpisodeTranscript.lock_record()
            # Load the record into the hidden RichTextCtrl.  Don't show the popup, as too many of these crashes the program!
            self.richTextCtrl.load_transcript(tmpEpisodeTranscript, showPopup=False)
            # Save the record.  This causes the PlainText to be added!  Don't show the popup, as too many of these crashes the program!
            self.richTextCtrl.save_transcript(use_transactions=False, showPopup=False)
            # Unlock the record
            tmpEpisodeTranscript.unlock_record()
            # Commit the Transaction
            dbCursor.execute("COMMIT")

            # Update the Record Counter
            counter += 1
            # Update the Progress Bar
            self.gauge.SetValue(counter)
            # This form can freeze up and appear non-responsive.  We should avoid that with every Episode Transcript (they are big).
            wx.YieldIfNeeded()


        # Get a list of the Quotes that need Plain Text extraction
        quotes = DBInterface.list_of_quotes(withoutPlainText=True)
        # Update User Info
        self.txtCtrl.AppendText("%5d Quote Records\n" % len(quotes))
        # Iterate through the list
        for quote in quotes:
            # Update User Info
            self.txtCtrl.AppendText("%5d  Quote:  %s\n" % (self.numRecords - counter, quote[1]))
            # Load the Quote Object
            tmpQuote = Quote.Quote(num=quote[0])

            # I'm not sure why I have to use a Transaction here.  But the transaction from the Transcript
            # object doesn't seem to be cutting it.  This at least makes things work!
            dbCursor.execute("BEGIN")
            # Lock the record
            tmpQuote.lock_record()
            # Load the record into the hidden RichTextCtrl.  Don't show the popup, as too many of these crashes the program!
            self.richTextCtrl.load_transcript(tmpQuote, showPopup=False)
            # Save the record.  This causes the PlainText to be added!  Don't show the popup, as too many of these crashes the program!
            self.richTextCtrl.save_transcript(use_transactions=False, showPopup=False)
            # Unlock the record
            tmpQuote.unlock_record()
            # Commit the Transaction
            dbCursor.execute("COMMIT")

            # Update the Record Counter
            counter += 1
            # Update the Progress Bar
            self.gauge.SetValue(counter)
            # This form can freeze up and appear non-responsive.  Every 20 items, we should avoid that
            if counter % 20 == 0:
                wx.YieldIfNeeded()

        # Get a list of the Clips that need Plain Text extraction
        clips = DBInterface.list_of_clips(withoutPlainText=True)
        # Update User Info
        self.txtCtrl.AppendText("%5d Clip Records\n" % len(clips))
        # Iterate through the list
        for clip in clips:
            # Update User Info
            self.txtCtrl.AppendText("%5d  Clip:  %d  %s\n" % (self.numRecords - counter, clip[0], clip[1]))
            # Load the Clip Object
            tmpClip = Clip.Clip(id_or_num=clip[0])
            # For each Transcript in the Clip ...
            for tmpClipTranscript in tmpClip.transcripts:

                # I'm not sure why I have to use a Transaction here.  But the transaction from the Transcript
                # object doesn't seem to be cutting it.  This at least makes things work!
                dbCursor.execute("BEGIN")
                # Lock the record
                tmpClipTranscript.lock_record()
                # Load the record into the hidden RichTextCtrl.  Don't show the popup, as too many of these crashes the program!
                self.richTextCtrl.load_transcript(tmpClipTranscript, showPopup=False)
                # Save the record.  This causes the PlainText to be added!  Don't show the popup, as too many of these crashes the program!
                self.richTextCtrl.save_transcript(use_transactions=False, showPopup=False)

                ## This alternate method also works, and is slightly faster, but is less object friendly.
                ## The speed difference is about 10 - 15 seconds in a 3 minute processing of 1400 clips.

## import re

##                    plaintext = self.richTextCtrl.GetValue()
##                    # Strip Time Codes
##                    regex = "%s<[\d]*>" % TransanaConstants.TIMECODE_CHAR
##                    reg = re.compile(regex)
##                    pos = 0
##                    for x in reg.findall(plaintext):
##                        pos = plaintext.find(x, pos, len(plaintext))
##                        plaintext = plaintext[ : pos] + plaintext[pos + len(x) : ]
##                    query = "UPDATE Transcripts2 SET PlainText = %s WHERE TranscriptNum = %s"
##                    values = (plaintext.encode('utf8'), tmpClipTranscript.number)
##                    query = DBInterface.FixQuery(query)
##                    result = dbCursor.execute(query, values)

                # Unlock the record
                tmpClipTranscript.unlock_record()
                # Commit the Transaction
                dbCursor.execute("COMMIT")

            # Update the Record Counter
            counter += 1
            # Update the Progress Bar
            self.gauge.SetValue(counter)
            # This form can freeze up and appear non-responsive.  Every 20 items, we should avoid that
            if counter % 20 == 0:
                wx.YieldIfNeeded()

        dbCursor.close()
