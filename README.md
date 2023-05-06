## About 

This is a Python script that allows the modular construction and sending of emails.

The construction of the email can be broken down into mutliple parts. Conditional statements allow for adaptive control of the content based on the category of recipient.

Email recipients are retrieved from [recipients.csv](https://github.com/callumr00/microsoft-outlook-emailer/blob/main/recipients.csv), a csv file that takes ```category```, ```name```, and ```email```.\
All fields are required. Any row without a name or valid email will be removed to avoid errors at run time. Rows will duplicate emails are also removed.

An email is then constructed for each of the recipients.\
The body of the email consists of the [Email Style](https://github.com/callumr00/microsoft-outlook-emailer/blob/main/Email%20Style.txt), [Email Introduction](https://github.com/callumr00/microsoft-outlook-emailer/blob/main/Email%20Introduction.txt), Email Details, and [Email Signature](https://github.com/callumr00/microsoft-outlook-emailer/blob/main/Email%20Signature.txt).\
The middle section of the email  - Email Details - is determined based on the category of the recipient. A category of ```1``` sets this to the text present in [Category 1.txt](https://github.com/callumr00/microsoft-outlook-emailer/blob/main/Category%201.txt) and a category of ```2``` sets this to the text present in [Category 2.txt](https://github.com/callumr00/microsoft-outlook-emailer/blob/main/Category%202.txt). 

Optionality can be increased by increasing the number of categories. It is also possible to adapt such that more than Email Details is conditional, increasing versatility.


The number of recipients to send to can be altered when calling the ```mass_email()``` function by changing the ```batch_size``` argument. By default, this is set to ```batch_size=500```. This means that a maximum of 500 emails will be sent out at one time, with fewer being sent if the ending is reached prematurely.
There is also an argument for the last email sent, ```last_email```. This allows for continuation from the previous batch. By default, this is set to ```last_email=None``` and results in starting from the beginning.
