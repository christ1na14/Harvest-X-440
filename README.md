"# Harvest-X-440" 
Before First Run:
1) Install Python 3
2) Configure Server To Handle Python CGI
3) Download or Clone

Can be ran locally, or from htdocs.

Change line 83 in upload.py depending on your excel file.

GETTING STARTED:
Languages/Tools: Python, Excel, PyCharm, Google, GitHub
Task Manager: Google Keep
The integrated development environment (IDE) is the software that you use to code. The Recommended IDE for this project is PyCharm but anyone (such as Spyder) will do just fine. 
PyCharm Installation Instructions:
Go to Jet Brains for Students and register using your aggie email.
Verify your email
Download PyCharm
Detailed Windows Instructions here
Detailed Mac Instructions here
The latest version of Python must be downloaded. Python 3.7. You do not have to download this to your personal laptop, but that means you will only be able to use the Engineering Building computers.
Latest Python Installation Instructions:
Go to https://www.python.org/downloads/
Choose your respective operating system
Download
To run Python Scripts (aka python files aka python code) in PyCharm, you must configure PyCharm. It is not too complicated, but if done wrong it could mess up Python or PyCharm. Google ‘configure pycharm to run python’ if you want to try to do this on your own.
Installing GitHub:
Go to GitHub and create an account
Go to GitHub Desktop
Follow installation instructions
Sign in to GitHub Desktop
Make sure to accept the (harvest-x) invitation, which should be in your aggie email
———————————————————————————————————————
METHODOLOGY:
How Each Report Category Is Defined:
Total Individuals/Total households: All account IDs
Unduplicated Individuals: Unique account IDs
Unduplicated households: Off-campus with same address, any on campus address for each unique submission ID
PsuedoCode (english instructions) Execution:
Total Households/Individuals (Participation File):
Import the participation excel file
Extract Account IDs
Count number of IDs
Unduplicated Individuals (Participation File):
Import the participation excel files
Extract Account IDs
Count unique number of IDs
Unduplicated households (Application File & Participation File):
Create a dictionary of submission IDs (key) and Addresses (value)
Find all unique submission IDs, then find unique address (cities at the moment) from that list, this gives total number of unique households - Application File
The existing values will equal to one household
Compare against total individuals/households from participation file
 


