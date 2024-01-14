# dependency
from ru_number_to_text.num2t4be import num2text_ordinal, num2text_numerical

```bash
PYTHONPATH=. /Applications/LibreOffice.app/Contents/Resources/python
>>> import os
>>> os.chdir('/home/user/myprojects')
>>> os.getcwd()
'/home/user/myprojects'
>>> from libreoffice_snippets.number_to_text import ReplaceNumberToText
>>> ReplaceNumberToText()
```

Output:
```
7359327549237
Replace? (Yes/No or Y/N, default = No)
Yes
Input tag (tag=("MGS"|"PGP"), default tag="MNS"):
сем трыльёнаў трыста пяцьдзясят дзевяць мільярдаў трыста дваццаць сем мільёнаў пяцьсот сорак дзевяць тысяч дзвесце трыццаць сёмы
82365492300
```
