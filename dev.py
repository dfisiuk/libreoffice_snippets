import uno
def getModel():
  # get the uno component context from the PyUNO runtime
  localContext = uno.getComponentContext()

  # create the UnoUrlResolver
  resolver = localContext.ServiceManager.createInstanceWithContext(
  "com.sun.star.bridge.UnoUrlResolver", localContext)

  # connect to the running office
  context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
  manager = context.ServiceManager

  # get the central desktop object
  desktop = manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

  # access the current writer document
  return desktop.getCurrentComponent()

def printInterfaces(obj):
    """
    Thanks to Jamie Boyle https://documenthacker.wordpress.com
    """
    text = str(obj)
    interfacesBlock = [z for z in text.split(' ') if z.startswith('supportedInterfaces=')][0]
    interfaceNames = []
    for longName in interfacesBlock[interfacesBlock.find('{')+1:interfacesBlock.find('}')].split(','):
      interfaceName = longName[longName .rfind('.')+1:]
      if interfaceName[0]=='X':
          interfaceName = interfaceName[1:]

      interfaceNames.append(interfaceName)

    interfaceNames.sort() #Sort all the names alphabetically
    for interfaceName in interfaceNames:
      if interfaceName[0]=='X':
         print(interfaceName[1:])
      else:
         print(interfaceName)

def printObjectProperties(obj):
  #Get the properties
  properties = list(obj.getPropertySetInfo().getProperties())
  #Sort alphabetically by name
  properties.sort(key = lambda x:x.Name)
  longest_len = max([len(z.Name) for z in properties])
  for property in properties:
    print(property.Name.ljust(longest_len)+'  -'+str(property.Type))
