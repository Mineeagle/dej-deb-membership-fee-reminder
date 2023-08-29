from helper_classes import ExcelHandler, EmailSender
from colorama import Fore


class Sender:
    # Cols of the data in the excel
    LAST_NAME_COL = "d"
    EMAIL_COL = "i"
    INVOICE_NUMBER_COL = "j"
    FEE_COL = "k"
    COMPANY_COL = "B"

    # Col and marks for sent and not sent emails
    CHECK_COL = "l"
    CHECK_TRUE = "✔"
    CHECK_FALSE = "X"

    # Row with the first data entry
    START_ROW = 2
    """ Sender class
    This class handles the sending of the emails and extracting the files
    
    :param excel_url: URL of the excel file
    :param sender_email: email address of the sender
    :param sender_password: password to the email
    :param year: year of the invoice
    
    Nothing needs to be changed.
    """

    def __init__(
        self,
        excel_url,
        sender_email,
        sender_password,
        year,
        smtp_server="webmail.esperanto.de",
        smtp_port=587,
    ):
        self.year = year
        self.smtp_port = smtp_port
        self.smtp_server = smtp_server

        self.eh = ExcelHandler(excel_url)
        self.sender_email = sender_email
        self.sender_password = sender_password
        print(Fore.YELLOW + "Parameter set.")

        self.start_sending()

    def start_sending(self):
        """Core method of the class
        This is the core method of the class. It first extracts everything from the
        excel file, and then sends emails.

        Nothing needs to be changed.
        """
        # Handle excel extraction and add check col
        excel_information = self.get_information_from_excel()
        print("Data from excel extracted.")
        self.eh.write(f"{Sender.CHECK_COL}1", "E-Mail sent")
        print("Added 'E-Mail sent' column.")

        print("---Sending-Log---")

        # Start Handling
        for row, element in enumerate(excel_information):
            row += 2

            # Handle if no email exists
            if element[Sender.EMAIL_COL] is None:
                self.eh.write(f"{Sender.CHECK_COL}{row}", Sender.CHECK_FALSE)
                print(
                    Fore.RED
                    + f"Mr./Mrs. {element[Sender.LAST_NAME_COL]} has no email address."
                )
                continue

            # Send email if email exists
            email_sender = EmailSender(
                self.sender_email,
                self.sender_password,
                self.smtp_server,
                self.smtp_port,
            )
            if element[Sender.LAST_NAME_COL] is not None:
                body = self.get_email_body_person(
                    element[Sender.LAST_NAME_COL],
                    element[Sender.INVOICE_NUMBER_COL],
                    element[Sender.FEE_COL],
                    self.year
                )
            else:
                body = self.get_email_body_company(
                    element[Sender.COMPANY_COL],
                    element[Sender.INVOICE_NUMBER_COL],
                    element[Sender.FEE_COL],
                    self.year
                )
            subject = self.get_email_subject(element[Sender.INVOICE_NUMBER_COL])
            email_sender.send_mail(body, subject, element[Sender.EMAIL_COL])
            self.eh.write(f"{Sender.CHECK_COL}{row}", Sender.CHECK_TRUE)
            print(Fore.GREEN + f"Sent email to: {element[Sender.EMAIL_COL]}")

    def get_email_body_person(self, last_name, invoice_number, fee, year):
        """Email body creation for singular people
        This method creates the email body.

        get_email_body() -> str: Body of the email

        Adapt the body in this method.
        """
        email_body = f"""Sehr geehrte(r) Frau/Herr {last_name},

uns ist aufgefallen, dass Sie Ihren Mitgliedsbeitrag für die DEB/DEJ für das Jahr {year} noch nicht gezahlt haben.
Aus diesem Grund bitten wir Sie den Mitgliedsbeitrag von {fee}€ unter der Rechnungsnummer {invoice_number} auf folgendes Konto zu überweisen:

[Daten eines Kontos]

Wir bedanken uns herzlich, dass Sie als Mitglied mit Ihrem Beitrag die Esperanto-Kultur in Deutschland unterstützen.

Mit freundlichen Grüßen,

Deutscher Esperanto-Bund/Deutsche Esperanto-Jugend
"""
        return email_body

    def get_email_body_company(self, company, invoice_number, fee, year):
        """Email body creation for a company
        This method creates the email body.

        get_email_body() -> str: Body of the email

        Adapt the body in this method.
        """
        email_body = f"""Sehr geehrte Damen und Herren des Unternehmens {company},

uns ist aufgefallen, dass Sie Ihren Mitgliedsbeitrag für die DEB/DEJ für das Jahr {year} noch nicht gezahlt haben.
Aus diesem Grund bitten wir Sie den Mitgliedsbeitrag von {fee}€ unter der Rechnungsnummer {invoice_number} auf folgendes Konto zu überweisen:

[Daten eines Kontos]

Wir bedanken uns herzlich, dass Sie als Mitglied mit Ihrem Beitrag die Esperanto-Kultur in Deutschland unterstützen.

Mit freundlichen Grüßen,

Deutscher Esperanto-Bund/Deutsche Esperanto-Jugend
"""
        return email_body

    def get_email_subject(self, invoice_number):
        """Email subject creation
        This method creates the email subject.

        get_email_subject() -> str: Subject of the email

        Nothing needs to be changed.
        Adapt the subject in this method.
        """
        return f"Zahlungsaufforderung DEB/DEJ {invoice_number}"

    def get_information_from_excel(self):
        """Extract the information from the excel file
        This method extracts the information from the excel file.

        get_information_from_excel() -> []: List with dicts of the extracted information

        Nothing needs to be changed.
        To change the cols, please refer to class variables at the top!
        """
        max_row = self.eh.get_max_row()
        res = []

        for row_num in range(2, max_row + 1):
            email = self.eh.read(f"{Sender.EMAIL_COL}{row_num}")
            last_name = self.eh.read(f"{Sender.LAST_NAME_COL}{row_num}")
            invoice_number = self.eh.read(f"{Sender.INVOICE_NUMBER_COL}{row_num}")
            fee = self.eh.read(f"{Sender.FEE_COL}{row_num}")
            company = self.eh.read(f"{Sender.COMPANY_COL}{row_num}")
            res.append(
                {
                    Sender.EMAIL_COL: email,
                    Sender.LAST_NAME_COL: last_name,
                    Sender.INVOICE_NUMBER_COL: invoice_number,
                    Sender.FEE_COL: fee,
                    Sender.COMPANY_COL: company,
                }
            )
        return res
