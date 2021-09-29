# Fetching attachments from Microsoft-365 Outlook mail using Python script 

The sole purpose of this project is to eliminate the need of logging into your email, downloading the attachments locally, then calling them into your Python script.

Instead, this code walks us through the processes of extracting attachments by iterating through your Outlook inbox, loading them into Python in the form of a byte string for further processing.

We will be utilizing `Mac/Linux terminal`, `Python 3.9.6`, and `Anaconda`

### Bonus Step. Setup a new virtual environment in terminal:

###### In Linux/Mac terminal, create an execusion environment with Python for packages and libraries control (always a good practice! Especially with package dependant projects 😉)

```
conda create -n jupytab-notebook-env python=3.9.6
conda activate jupytab-notebook-env
```

## `exchangelib` Installation



We will be using `exchangelibe` library to coummnicate with `Office365`. Official documentation: [exchangelib](https://ecederstrand.github.io/exchangelib/#installation).

###### In the project environment run: 

```
conda install exchangelib
```
##### *depends on your setup, this might through an error that:*

>  PackagesNotFoundError: The following packages are not available from current channels:

> (it will print out a list of urrent channels)

###### Fix: start by displaying a list of active channels inside conda by typing either one of the following commands:

```
conda info
```
```
conda config --show channels
```

###### *make sure you have a forge channel in the list, if not we can easily add one*

###### Add a forge channel `conda-forge` to your list of channels with this command:

```
conda config --append channels conda-forge
```

*It tells conda to also look on the conda-forge channel when you search for packages*

#### You can then simply install the two packages with:

```
conda install exchangelib
```










