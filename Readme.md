## Whatsapp Multisend

Whatsapp Multisend is an unofficial application that allows you to send text messages and images to as many contacts as you want without using the Whatsapp API. It reads contacts from a `contacts.xlsx` file where the first column is the contact name and the second column is the phone number. The phone numbers should be saved on your phone. If they are not saved, you can use another application named ExcelToVcard to generate a vCard and then import it to save those contacts.

### Installation

1. Clone or download this repository.
2. Install the required dependencies by running installing all libraries.
3. Save your contacts in the `contacts.xlsx` file.
4. Run the `whatsapp_multisend.py` script by running the following command:
```python whatsapp_multisend.py```
5. Wait for the script to open Whatsapp Web and scan the QR code to login.

### Usage

After you've logged in, the script will start pushing messages one by one to the contacts in the `contacts.xlsx` file. If the contact is not on Whatsapp, it will be skipped. You can send text messages or images by specifying the message type when prompted.

### Disclaimer

Please note that this application should not be used for spamming. Use it responsibly and within the limits of Whatsapp's terms of service. We are not responsible for any misuse of this application.
