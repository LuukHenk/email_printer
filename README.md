# email_printer
Version 0.0.1 <br />
Print email .pdf attachment('s) and email body in .pdf format

## Quick start
### Requirements
* Python3.6
* pdfkit (verified for v0.6.1) - [How to install](https://github.com/JazzCore/python-pdfkit)

### Installing
First clone the repository to your local computer: `$ git clone https://github.com/LuukHenk/email_printer.git`

To set the mail host/port and to set other options, open main.py in the repository and find the 'Input data' header. Replace the input data with the setup preferred (the arguments in the script give a quick explaination about each variable set).


## Usage
Run `$ python3 main.py` and login. The attachments and mail text will be downloaded to the saving path location.

## Delevopment
### Issues
* I-002: There is no error handling yet
* I-003: Send files to the printer after saving them (or instead of saving them)
* I-004: Make email prefixes optional
* I-005: Generate output folder if it does not exist yet

### Changelog
* C-I-001: Fixed downloading of images in HTML code

### To do
* T-001: Add detection of the mail header and add it to the mail text output
* T-002: Also check for plain mail text instead of only HTML text
* T-003: Replace the Input data header with input arguments
* T-004: Make the email downloader executable
