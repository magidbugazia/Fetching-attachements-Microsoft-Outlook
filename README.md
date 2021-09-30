# Fetching attachments from Microsoft-365 Outlook mail using Python script 

We will be utilizing `Mac/Linux terminal`, `Python 3.9.6`, and `Jupyter Notebook`

The sole purpose of this project is to eliminate the need of logging into your email, downloading the attachments locally, then calling them into your Python script.

Instead, this code walks us through the processes of extracting attachments by iterating through your Outlook inbox, loading them into Python in the form of a byte string for further processing.


### Bonus Step. Setup a new virtual environment in terminal:

###### In Linux/Mac terminal, create an execusion environment with Python for packages and libraries control (always a good practice! Especially with package dependant projects 😉)

```
conda create -n jupytab-notebook-env python=3.9.6
conda activate jupytab-notebook-env
```

## 1. Installing `exchangelib` 

We will be using `exchangelibe` library (official documentation: [exchangelib](https://ecederstrand.github.io/exchangelib/#installation)) to coummnicate with `Office365`. 

- In the project environment run: `conda install exchangelib `

- Make sure you have a forge-channel already setup:

  - To display current list run: ```conda info``` ```conda config --show channels```

  - Add a forge channel (conda-forge) to your list with this command: ```conda config --append channels conda-forge```


## 2. Fetching attachments in `Jupyter Notebook`


 ``` Python
import io

from exchangelib import DELEGATE, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message, \
  CalendarItem, HTMLBody
import pandas as pd

credentials = Credentials('YOUR_EMAIL@host.com', 'YOUR_PASSWORD')
config = Configuration(server='outlook.office365.com', credentials=credentials)
account = Account(
    primary_smtp_address='YOUR_EMAIL@host.com',
    config=config,
    autodiscover=False,
    access_type=DELEGATE
)
```

Find all attachments in the inbox matching the email subject you specify:

```Python
item = account.inbox.all().get(subject='TARGET_EMAIL_SUBJECT')
 ```
Iterate through the attachments and match with the attachment name and extension you specify:

```python
for attachment in item.attachments:
    if attachment.name == 'filename.extension':
        my_excel_file_in_bytes = attachment.content
        break
else:
    assert False, 'No attachment with that name'
```

###### The attachment content will be the excel file in the form of a byte string

Convert to a file-like object and read the excel file in memory:

 - In terminal in our project environment run the command: `conda install xlrd`
 - Back to Jupyter Notebook run the code:


``` Python
my_excel_file_io = io.BytesIO(my_excel_file_in_bytes)
df = pd.read_excel(io=my_excel_file_io)
```

```python
print(df)
```

## Further Work to Follow:

- Directly expose real-time dataframe from Jupyter NoteBook into Tableau [Teableau_JupyterNotebook_Dynamic_Connection](https://github.com/magidbugazia/Realtime_Teableau_JupyterNotebook.git)
