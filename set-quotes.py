import re
from libreoffice_snippets import dev

def getModel():
  """
  just before deployment we can change this to
  return XSCRIPTCONTEXT.getDocument() # embedded
  """
  return dev.getModel() # via socket

model = dev.getModel()

search = model.createSearchDescriptor()
search.setPropertyValue('SearchRegularExpression', True)
search.setPropertyValue('SearchCaseSensitive', True)

file = open('girl-with-d-tatoo/girl-with-d-tattoo-dic.csv', 'r')
words= file.readlines()

StartFromBegining = True
for line in words:
    list = line.strip().split(',')
    word = list[0]
    quoted = list[1]
    if quoted == 'True':
      print('word:','"'+word+'"','quoted:', quoted)
      expr = '(«?)(' + word + ')(»?)'

      search.setSearchString(expr)
      text = model.Text
      xSelectionSupplier = model.getCurrentController()
      cursor = xSelectionSupplier.getViewCursor()

      if StartFromBegining:
        # create an XTextRange at the start of the document
        tRange = text.Start
      else:
        # create an XTextRange at the current cursor
        tRange = cursor

      oFound  = model.findNext(tRange, search)
      while oFound:
        xText = oFound.getText()
        xWordCursor = xText.createTextCursorByRange(oFound)
        xSelectionSupplier.select(xWordCursor)

        found_str = oFound.getString()

        print('found string: ' + found_str, 'word: ' + word)

        m = re.search(expr, found_str)
        print(m.group(1), m.group(2), m.group(3))
        name = m.group(2)
        if not ( m.group(1) and  m.group(3) ):
          # print('not quoted')
          answer = input('Set quotes? ("Yes/No" or "Y/N", default="Yes") \n') or 'Yes'
          answer = answer.lower()
          if answer == 'yes' or answer == 'y':
            oFound.setString('«'+name+'»')
            input("Press Enter to continue...")

        oFound  = model.findNext(oFound.End, search)
