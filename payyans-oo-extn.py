#!/usr/bin/env python
#
# Copyright 2008-2009 Santhosh Thottingal <santhosh.thottingal@gmail.com>,
# Nishan Naseer <nishan.naseer@gmail.com>, Manu S Madhav <manusmad@gmail.com>,
# Rajeesh K Nambiar <rajeeshknambiar@gmail.com>
# Copyright (C) 2009 Anoop John  <www.zyxware.com>
# http://www.smc.org.in
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# at your option) any later version.
#       
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#       
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
# MA 02110-1301, USA.

import sys
import os
from optparse import OptionParser
import codecs 
import uno
import unohelper
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, \
BUTTONS_ABORT_IGNORE_RETRY, BUTTONS_YES_NO_CANCEL, BUTTONS_YES_NO, BUTTONS_RETRY_CANCEL, \
DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, \
DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE

# http://www.oooforum.org/forum/viewtopic.phtml?t=59580
class MessageBox:
  '''Message box for OpenOffice.org, like the one in the Basic macro language. To specify a MsgBox type, use the named constants of this class or the equivalent numbers described in the StarBasic online help. Specify a parent window on initialization.'''

  # Named constants for ease of use:
  OK = 0
  OK_CANCEL = 1
  ABORT_RETRY_IGNORE = 2
  YES_NO_CANCEL = 3
  YES_NO = 4
  RETRY_CANCEL = 5
  ERROR = 16
  QUERY = 32
  WARN = 48
  INFO = 64
  DEFAULT_FIRST = 128
  DEFAULT_SECOND = 256
  DEFAULT_THIRD = 512
  RESULT_OK = 1
  RESULT_CANCEL = 2
  RESULT_ABORT = 3
  RESULT_RETRY = 4
  RESULT_IGNORE = 5
  RESULT_YES = 6
  RESULT_NO = 7

  # Mapping above StarBasic MsgBox constants to awt.MessageBoxButtons and icons:
  dInput = {
    OK_CANCEL : BUTTONS_OK_CANCEL,
    # the following constant should be named BUTTONS_ABORT_RETRY_IGNORE:
    ABORT_RETRY_IGNORE : BUTTONS_ABORT_IGNORE_RETRY,
    YES_NO_CANCEL : BUTTONS_YES_NO_CANCEL,
    YES_NO : BUTTONS_YES_NO,
    RETRY_CANCEL : BUTTONS_RETRY_CANCEL,
    ERROR : 'errorbox',
    QUERY : 'querybox',
    WARN : 'warningbox',
    INFO : 'infobox', # info always shows one OK button alone!
    129 : BUTTONS_OK_CANCEL + DEFAULT_BUTTON_OK,
    130 : BUTTONS_ABORT_IGNORE_RETRY + DEFAULT_BUTTON_CANCEL,
    131 : BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_YES,
    132 : BUTTONS_YES_NO + DEFAULT_BUTTON_YES,
    133 : BUTTONS_RETRY_CANCEL + DEFAULT_BUTTON_RETRY,
    257 : BUTTONS_OK_CANCEL + DEFAULT_BUTTON_CANCEL,
    258 : BUTTONS_ABORT_IGNORE_RETRY + DEFAULT_BUTTON_RETRY, # retry is 2nd!
    259 : BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO,
    260 : BUTTONS_YES_NO + DEFAULT_BUTTON_NO,
    261 : BUTTONS_RETRY_CANCEL + DEFAULT_BUTTON_CANCEL,
    # DEFAULT_BUTTON_IGNORE doesn't work at all. Use retry in this case:
    # 514 : BUTTONS_ABORT_IGNORE_RETRY  + DEFAULT_BUTTON_IGNORE,
    514 : BUTTONS_ABORT_IGNORE_RETRY  + DEFAULT_BUTTON_RETRY,
    515 : BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_CANCEL,
    }
  dOutput = {
    1 : RESULT_OK,
    2 : RESULT_YES,
    3 : RESULT_NO,
    4 : RESULT_RETRY,
    5 : RESULT_IGNORE,
    0 : RESULT_CANCEL, # there is no constant for ABORT
    }
     
  def __init__(self, XParentWindow):
    '''Set needed objects on init'''
    try:
      self.Parent = XParentWindow
      self.Toolkit = XParentWindow.getToolkit()
    except:
      raise AttributeError, 'Did not get a valid parent window'
           
  def msgbox(self, message='', flag=0, title=''):
    '''Wrapper for com.sun.star.awt.XMessageBoxFactory.'''
    rect = uno.createUnoStruct('com.sun.star.awt.Rectangle')
    stype, buttons = self.getFlags(flag)
    box = self.Toolkit.createMessageBox(self.Parent, rect, stype, buttons, title, message)
    e = box.execute()
    # the result of execute() does not distinguish between Cancel and Abort:
    if (e == 0) and (buttons & BUTTONS_ABORT_IGNORE_RETRY == BUTTONS_ABORT_IGNORE_RETRY):
      r  = self.RESULT_ABORT
    else:
      try:
        r = self.dOutput[e]
      except KeyError:
        raise KeyError, 'Lookup of message box result '+ str(e) +' failed'

    return r

  def getFlags(self, flag):
    s = self.dInput.get(flag & 112, 'messbox')
    try:
      b = self.dInput[flag & 903]
    except KeyError:
      b = self.dInput.get(flag & 7, BUTTONS_OK)
    # print 'B/D-Flag:', flag & 7, flag & 896, 'Return:', hex(b), s
    return s, b 

class Payyan:

  def __init__(self):
    self.input_filename =""
    self.output_filename=""
    self.mapping_filename=""
    self.rulesDict=None
    self.pdf=0
    
  def word2ASCII(self, unicode_text):
    index = 0
    prebase_letter = ""
    ascii_text=""
    self.direction = "u2a"
    self.rulesDict = self.LoadRules()
    while index < len(unicode_text):
      '''This takes care of conjuncts '''
      for charNo in [3,2,1]:
        letter = unicode_text[index:index+charNo]
        if letter in self.rulesDict:
          ascii_letter = self.rulesDict[letter]
          letter = letter.encode('utf-8')
          '''Fixing the prebase mathra'''
          '''TODO: Make it generic , so that usable for all indian languages'''
          if letter == 'ൈ':
            ascii_text = ascii_text[:-1] + ascii_letter*2 + ascii_text[-1:]
          elif (letter == 'ോ') | (letter == 'ൊ') | (letter == 'ൌ'):  #prebase+postbase mathra case
            ascii_text = ascii_text[:-1] + ascii_letter[0] + ascii_text[-1:] + ascii_letter[1]
          elif (letter == 'െ') | (letter == 'േ') |(letter == '്ര'):  #only prebase
            ascii_text = ascii_text[:-1] + ascii_letter + ascii_text[-1:]
          else:
            ascii_text = ascii_text + ascii_letter            
          index = index+charNo
          break
        else:
          if(charNo==1):
            index=index+1
            ascii_text = ascii_text + letter
            break;
          '''Did not get'''        
          ascii_letter = letter

    return ascii_text
    
  def Uni2Ascii(self):
    if self.input_filename :
      uni_file = codecs.open(self.input_filename, encoding = 'utf-8', errors = 'ignore')
    else :
      uni_file = codecs.open(sys.stdin, encoding = 'utf-8', errors = 'ignore')      
    text = ""
    if self.output_filename :
      output_file = codecs.open(self.output_filename, encoding = 'utf-8', errors = 'ignore',  mode='w+')      
    while 1:
      text =uni_file.readline()
      if text == "":
        break
      ascii_text = ""  
      ascii_text = self.word2ASCII(text)
                  
      if self.output_filename :
        output_file.write(ascii_text)
      else:
        print ascii_text.encode('utf-8')
    return 0
    
  def word2Unicode(self, ascii_text):
    index = 0
    post_index = 0
    prebase_letter = ""
    postbase_letter = ""
    unicode_text = ""
    next_ucode_letter = ""
    self.direction="a2u"
    self.rulesDict = self.LoadRules()
    while index < len(ascii_text):
      for charNo in [2,1]:
        letter = ascii_text[index:index+charNo]
        if letter in self.rulesDict:
          unicode_letter = self.rulesDict[letter]
          if(self.isPrebase(unicode_letter)):  
            prebase_letter = unicode_letter
          else:
            post_index = index+charNo
            if post_index < len(ascii_text):
              letter = ascii_text[post_index]
              if letter in self.rulesDict:
                next_ucode_letter = self.rulesDict[letter]
                if self.isPostbase(next_ucode_letter):
                  postbase_letter = next_ucode_letter
                  index = index + 1
            if  ((unicode_letter.encode('utf-8') == "എ") |
                ( unicode_letter.encode('utf-8') == "ഒ" )):
              unicode_text = unicode_text + postbase_letter + self.getVowelSign(prebase_letter , unicode_letter)
            else:
              unicode_text = unicode_text + unicode_letter + postbase_letter + prebase_letter
            prebase_letter=""
            postbase_letter=""
          index = index + charNo
          break
        else:
          if charNo == 1:
            unicode_text = unicode_text + letter
            index = index + 1
            break
          unicode_letter = letter
    return unicode_text  
  
  def Ascii2Uni(self):
    if self.pdf :
      command = "pdftotext '" + self.input_filename +"'"
      process = os.popen(command, 'r')
      status = process.close()
      if status:
        print "The input file is a PDF file. To convert this the  pdftotext  utility is required. "
        print "This feature is available only for GNU/Linux Operating system."
        return 1  # Error - no pdftotext !
      else:
        self.input_filename =  os.path.splitext(self.input_filename)[0] + ".txt"
    if self.input_filename :
      ascii_file = codecs.open(self.input_filename, encoding = 'utf-8', errors = 'ignore')
    else :
      ascii_file = codecs.open(sys.stdin, encoding = 'utf-8', errors = 'ignore')      
    
    text = ""
    if self.output_filename :
      output_file = codecs.open(self.output_filename, encoding = 'utf-8', errors = 'ignore',  mode='w+')      
  
    while 1:
      text =ascii_file.readline()
      if text == "":
        break
      unicode_text = ""
      unicode_text = self.word2Unicode(text)
      
      if self.output_filename :
        output_file.write(unicode_text)
      else:
        print unicode_text.encode('utf-8')
    return 0

  def getVowelSign(self, vowel_letter, vowel_sign_letter):
    vowel=  vowel_letter.encode('utf-8')
    vowel_sign=  vowel_sign_letter.encode('utf-8')
    if vowel == "എ":
      if vowel_sign == "െ":
        return "ഐ"
    if vowel == "ഒ":
      if vowel_sign == "ാ":
        return "ഓ"
      if vowel_sign =="ൗ":
        return "ഔ"
    return (vowel_letter+ vowel_sign_letter)

  def isPrebase(self, letter):
     unicode_letter = letter.encode('utf-8')
     if(   ( unicode_letter == "േ"  ) | (   unicode_letter ==  "ൈ" ) |   ( unicode_letter ==  "ൊ" )   | ( unicode_letter ==  "ോ"  ) |  ( unicode_letter == "ൌ"  )
           |  ( unicode_letter == "്ര"  )  |  ( unicode_letter == "െ"  ) 
            ):
      return True
     else:
      return False
      
  def isPostbase(self, letter):
    unicode_letter = letter.encode('utf-8')
    if ( (unicode_letter == "്യ") | (unicode_letter == "്വ") ):
      return True
    else:
      return False
          
  def LoadRules(self):  
    if(self.rulesDict):
      return self.rulesDict
    rules_dict = dict()
    line = []
    line_number = 0
    rules_file = codecs. open(self.mapping_filename,encoding='utf-8', errors='ignore')
    while 1:
      ''' Keep the line number. Required for error reporting'''
      line_number = line_number + 1 
      text = unicode( rules_file.readline())
      if text == "":
        break
      '''Ignore the comments'''
      if text[0] == '#': 
        continue 
      line = text.strip()
      if(line == ""):
        continue 
      if(len(line.split("=")) != 2):
        print "Error: Syntax Error in the Ascii to Unicode Map in line number ",  line_number
        print "Line: "+ text
        return 2  # Error - Syntax error in Mapping file 
      lhs = line.split("=") [ 0 ]  
      rhs = line.split("=") [ 1 ]  
      lhs=lhs.strip()
      rhs=rhs.strip()
      if self.direction == 'a2u':
        rules_dict[lhs]=rhs
      else:
        rules_dict[rhs]=lhs
    return rules_dict

def ConvertWithPayyans (direction):    
  import string
 
  xModel = XSCRIPTCONTEXT.getDocument()

  xSelectionSupplier = xModel.getCurrentController()

  xIndexAccess = xSelectionSupplier.getSelection()
  count = xIndexAccess.getCount();
  if(count>=1):
    i=0
    p = Payyan();
    while i < count :
      xTextRange = xIndexAccess.getByIndex(i);
      theString = xTextRange.getString();
      font = xTextRange.getPropertyValue("CharFontName");
      maps_folder = '/usr/share/payyans/maps/';
      if font == "ML-TTKarthika": 
        mapping_file = 'karthika.map';
      elif font == "ML-TTAmbili":   
        mapping_file = 'ambili.map';
      else:
        # will be used for u2a conversion and also for any unmapped fonts
        mapping_file = 'karthika.map'  
      mapping_file = maps_folder + mapping_file;
      p.mapping_filename = mapping_file
      if len(theString)==0 :
        xText = xTextRange.getText();
        xWordCursor = xText.createTextCursorByRange(xTextRange);
        if not xWordCursor.isStartOfWord():
          xWordCursor.gotoStartOfWord(False);
        xWordCursor.gotoNextWord(True);
        theString = xWordCursor.getString();
        if (direction == "a2u") : 
          newString = p.word2Unicode(theString);
        else : 
          newString = p.word2ASCII(theString);
        if newString :
          xWordCursor.setString(newString);
          xSelectionSupplier.select(xWordCursor);
      else :
        if (direction == "a2u") : 
          newString = p.word2Unicode(theString);
        else : 
          newString = p.word2ASCII(theString);
        if newString:
          xTextRange.setString(newString);
          xSelectionSupplier.select(xTextRange);
      i+= 1

def MsgBox( message='', flag=0, title='' ):
  doc = XSCRIPTCONTEXT.getDocument()
  parentwin = doc.CurrentController.Frame.ContainerWindow
  a = MessageBox(parentwin)
  a.msgbox(message, flag, title)
   
def A2U (event = None):    
  """Convert ASCII to Unicode..."""
  ConvertWithPayyans ("a2u")
    
def U2A (event = None):
  """Convert Unicode text to ASCII..."""
  ConvertWithPayyans ("u2a")

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
  None, \
  "org.openoffice.script.smc.payyans_oo_extn", \
  ("org.openoffice.script.smc.payyans_oo_extn-service",),)

# List of exported scripts available as macros
g_exportedScripts = U2A, A2U,

    
