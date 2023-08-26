from sender_class import Sender
from getpass import getpass
from colorama import Fore


def get_input():
    print("version: 1.1")
    url = input("Please enter the url of the Excel file. > ")
    sender_email = input("Please enter the sender email adress. > ")
    sender_password = getpass(
        "Please enter the password for the entered email address. > "
    )
    year = input("Please enter the year for the invoice. > ")

    smtp_server = input(
        "Please enter the smtp_webserver. Standard is 'webmail.esperanto.de'. Press 'enter' to use this standard value. > "
    )
    if smtp_server == "":
        smtp_server = "webmail.esperanto.de"

    smtp_port = input(
        "Please enter the smtp port. Standard is 587. Press enter to use this standard value. > "
    )
    if smtp_port == "":
        smtp_port = 587
    else:
        smtp_port = int(smtp_port)

    return url, sender_email, sender_password, year, smtp_server, smtp_port


def main():
    url, sender_email, sender_password, year, smtp_server, smtp_port = get_input()
    print(Fore.YELLOW + "---Got input---")
    Sender(url, sender_email, sender_password, year, smtp_server, smtp_port)
    print("---Sending-Log---")
    print("Press enter to exit the program...")


if __name__ == "__main__":
    main()
