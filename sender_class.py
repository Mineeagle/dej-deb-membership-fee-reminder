from helper_classes import ExcelHandler, EmailSender


class Sender:
    LAST_NAME_COL = "d"
    EMAIL_COL = "i"
    INVOICE_NUMBER_COL = "j"
    FEE_COL = "k"

    CHECK_COL = "l"
    CHECK_TRUE = "✔"
    CHECK_FALSE = "X"

    START_ROW = 2

    def __init__(self, excel_url, sender_email, sender_password, year=2023):
        self.year = year

        self.eh = ExcelHandler(excel_url)
        print("Excel file has been loaded.")
        self.sender_email = sender_email
        self.sender_password = sender_password

        self.start_sending()

    def start_sending(self):
        # Handle excel extraction and add check col
        excel_information = self.get_information_from_excel()
        print("Data from excel extracted.")
        self.eh.write(f"{Sender.CHECK_COL}1", "E-Mail sent")
        print("Added 'E-Mail sent' column.")

        print("---")

        # Start Handling
        for row, element in enumerate(excel_information):
            row += 2

            # Handle if no email exists
            if element[Sender.EMAIL_COL] is None:
                self.eh.write(f"{Sender.CHECK_COL}{row}", Sender.CHECK_FALSE)
                print(f"Mr./Mrs. {element[Sender.LAST_NAME_COL]} has no email address.")
                continue

            # Send email if email exists
            email_sender = EmailSender(self.sender_email, self.sender_password)
            body = self.get_email_body(
                element[Sender.LAST_NAME_COL],
                element[Sender.INVOICE_NUMBER_COL],
                element[Sender.FEE_COL],
                self.year,
            )
            subject = self.get_email_subject(element[Sender.INVOICE_NUMBER_COL])
            email_sender.send_mail(body, subject, element[Sender.EMAIL_COL])
            self.eh.write(f"{Sender.CHECK_COL}{row}", Sender.CHECK_TRUE)
            print(f"Sent email to: {element[Sender.EMAIL_COL]}")

    def get_email_body(self, last_name, invoice_number, fee, year):
        email_body = f"""Sehr geehrte(r) Frau/Herr {last_name},

uns ist aufgefallen, dass Sie Ihren Mitgliedsbeitrag für die DEB/DEJ für das Jahr {year} noch nicht gezahlt haben.
Aus diesem Grund bitten wir Sie den Mitgliedsbeitrag von {fee}€ unter der Rechnungsnummer {invoice_number} auf folgendes Konto zu überweisen:

[Daten eines Kontos]

Sollten Sie der Zahlungsaufforderung nicht nachkommen, sehen wir uns dazu verpflichtet anderweitige Maßnahmen einzuleiten.

Mit freundlichen Grüßen,

Deutscher Esperanto-Bund/Deutsche Esperanto-Jugend
"""
        return email_body

    def get_email_subject(self, invoice_number):
        return f"Zahlungsaufforderung DEB/DEJ {invoice_number}"

    def get_information_from_excel(self):
        max_row = self.eh.get_max_row()
        res = []

        for row_num in range(2, max_row + 1):
            email = self.eh.read(f"{Sender.EMAIL_COL}{row_num}")
            last_name = self.eh.read(f"{Sender.LAST_NAME_COL}{row_num}")
            invoice_number = self.eh.read(f"{Sender.INVOICE_NUMBER_COL}{row_num}")
            fee = self.eh.read(f"{Sender.FEE_COL}{row_num}")
            res.append(
                {
                    Sender.EMAIL_COL: email,
                    Sender.LAST_NAME_COL: last_name,
                    Sender.INVOICE_NUMBER_COL: invoice_number,
                    Sender.FEE_COL: fee,
                }
            )
        return res

