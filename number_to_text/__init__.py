from __future__ import unicode_literals
from grammar_be.num2t4be import vars
from grammar_be.num2t4be import num2text_ordinal
from ru_number_to_text.num2t4ru import num2text

from libreoffice_snippets import dev # remove before deployment

def getModel():
  """
  just before deployment we can change this to
  return XSCRIPTCONTEXT.getDocument() # embedded
  """
  return dev.getModel() # via socket

# search patterns:
# '\d{1,}' - 50, 345
# '\d{1,}-[ях]' - 1980-я, 1990-х
# '\d{1,}-\w+' - 23-метровай (складаныя словы, першай (або апошняй) часткай якіх з’яўляецца лічба любога злічэння)

def ReplaceNumberToText():
    model = getModel()
    search = model.createSearchDescriptor()
    # dev.printObjectProperties(search) # explore the object
    search.setPropertyValue('SearchRegularExpression', True)

    search.setSearchString('(\d{1,})') # search numbers
    # get the XText interface
    text = model.Text
    # dev.printObjectProperties(text) # explore the object

    # create an XTextRange at the start of the document
    tRange = text.Start

    # the writer controller impl supports the css.view.XSelectionSupplier interface
    xSelectionSupplier = model.getCurrentController()

    oFound  = model.findNext(tRange, search)

    while oFound:
       print(oFound.getString())
       num = int(oFound.getString())

       xText = oFound.getText()
       xWordCursor = xText.createTextCursorByRange(oFound)
       xSelectionSupplier.select(xWordCursor)
       answer = input('Replace number to text? ("Yes/No" or "Y/N", default="No") \n') or 'No'
       answer = answer.lower()
       if answer == 'yes' or answer == 'y':
         num_type = input('Changer number to numerical (0) or ordinal (1)? (default=0) \n') or '0'
         if int(num_type) == 1:
           strTag=input('Input tag (examples: "MGS","PGP"; default="MNS"): ') or 'MNS'
           newString = num2text_ordinal(num,tag=strTag)
         else:
            newString = num2text(num)
         print(newString)
         oFound.setString(newString)

       oFound  = model.findNext(oFound.End, search)

if __name__ == '__main__':
  ReplaceNumberToText()

g_exportedScripts = ReplaceNumberToText,