from __future__ import unicode_literals
from grammar_be.num2t4be import vars
from grammar_be.num2t4be import num2text_ordinal, num2text_numerical
# from ru_number_to_text.num2t4ru import num2text

from libreoffice_snippets import dev # remove before deployment

import re

strFuncDescription=(
  u'\nСкланенне колькасных і парадкавых лічэбнікаў.\n'
  u'\ntag="<род><скланненне><лік>", options="<options>"\n'
  u'\nрод: "M" - мужчынскі, "F" - жаночы, "N" - ніякі, "P" - адсутны;\n'
  u'case (склон): "N" - Назоўны,"G" - Родны, "D" - Давальны, "A" - Вінавальны, "I" - Творны, "L" - Месны;\n'
  u'number (лік): "S" - адзіночны, "P" - множны;\n'
  u'options: "anim" - адуш., "inanim" - неадуш.\n')

patterns = (
  #  (u'()(\d+)-х()',1,u'PGP',None), #'90-x'
  #  (u'()(\d+)-я()',1,u'PNP',None), #'60-я'
  #  (u'()(\d+)-мі()',1,u'PIP',None), #'70-мі'
  #  (u'()(\d+)-га()',1,u'MGS',None),  #  66-га
  #  (u'([Ууў]\s)(\d+)(\sгодзе)',1,u'MLS',None), # у 1994 годзе
  #  (u'()(\d+)(\sгода)',1,u'MGS',None), # 1965 года
  #  (u'([Зз]\s)(\d{4})()',1,u'MGS',None), # з 1976
  #  (u'([Пп]а\s)(\d{4})()',1,u'MNS',None), # па 1982
  #  (u'()(\d+)-(\w{3,})',0,None,True), # '24-гадзіннага'
   ('()(\d+)()',None,None,None),
)
def getModel():
  """
  just before deployment we can change this to
  return XSCRIPTCONTEXT.getDocument() # embedded
  """
  return dev.getModel() # via socket

def ReplaceNumberToText(StartFromBegining=True):
    model = getModel()
    search = model.createSearchDescriptor()

    # dev.printObjectProperties(search) # explore the object
    search.setPropertyValue('SearchRegularExpression', True)

    for expr in patterns:
      print('expr: ' + expr[0])
      search.setSearchString(expr[0])
      # get the XText interface
      text = model.Text
      # dev.printObjectProperties(text) # explore the object

      # the writer controller impl supports the css.view.XSelectionSupplier interface
      xSelectionSupplier = model.getCurrentController()

      cursor = xSelectionSupplier.getViewCursor()
      # cursorPos = cursor.getPosition()
      # text.createTextCursorByRange(cursorPos)

      if StartFromBegining:
        # create an XTextRange at the start of the document
        tRange = text.Start
      else:
        # create an XTextRange at the start of the current cursor
        tRange = cursor

      # create an XTextRange at the start of the document
      tRange = text.Start

      oFound  = model.findNext(tRange, search)
      while oFound:
        xText = oFound.getText()
        xWordCursor = xText.createTextCursorByRange(oFound)
        xSelectionSupplier.select(xWordCursor)

        num_type = expr[1]
        strTag  = expr[2]
        found_str = oFound.getString()

        print('found string: ' + found_str, 'expr: ' + expr[0])

        answer = input('Replace number to text? ("Yes/No" or "Y/N", default="Yes") \n') or 'Yes'
        answer = answer.lower()
        if answer == 'yes' or answer == 'y':
          if not num_type:
            num_type = input('Changer number to numerical (0) or ordinal (1)? (default=0) \n') or '0'
          if not strTag:
            strTag=input(strFuncDescription + 'Input tag (default="MNS"): ') or 'MNS'

          m = re.search(expr[0], found_str)
          num = int(m.group(2))

          print('Number: ' + str(num), 'num_type: '+ str(num_type), 'Tag: ' + strTag)

          if int(num_type) == 1:
            newString = num2text_ordinal(num,tag=strTag)
          else:
             newString = num2text_numerical(num,tag=strTag)
          if expr[3] == True:
            newString = ''.join(newString.split(' '))
          print(newString)
          oFound.setString(m.group(1) + newString + m.group(3))

        oFound  = model.findNext(oFound.End, search)

if __name__ == '__main__':
  ReplaceNumberToText()

g_exportedScripts = ReplaceNumberToText,
