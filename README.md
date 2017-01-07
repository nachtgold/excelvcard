# excelvcard
Some scripts to link the two worlds: Spreadsheets (OpenDocumentXML) and vCards.

The code is built upon the two libraries [openpyxl](https://pypi.python.org/pypi/openpyxl) and [vobject](https://pypi.python.org/pypi/vobject/).

## contacts_excel_to_vcard.py

I was asked if it is possible to convert a Excel based contacts list to business cards, as you know them from mobile phones. In the internet I found no tool, which does not mean that there is none. So I started to practice something with python.

### Layout of the Excel file

The column layout of our sample file:

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| Name | Address | Mobile phone | Home phone | Work phone | Mail address | Birthday | Notes |
| World, Hello | Street, zipcode city | +1234 | +1234 | +1234 | a@b.cd | Excel date value | lorem ipsum |
|   | 2nd Street, another city |    |   |   |   |   |   |

### Data spefic details

The address value could have multiple formats:

* city
* street, city
* street, zipcode city

The birthday value could be a Excel datetime or only a date without year in the german "dd.mm." format.

If a row has no name value the script guess, that the data belongs to the previous row and combines boths records. 

### Result

After processing the first data row from the example table the rendered vCard could look like: 

![Demo of a resulting vCard](https://github.com/nachtgold/excelvcard/blob/master/demo.png?raw=true)

Note: Because I live in Germany, my script prefixes phone numbers with +49. You could change it to your home country.

### So feel free to use it for your own experiments.