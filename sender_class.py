from helper_classes import ExcelHandler, EmailSender


class Sender:
    LAST_NAME_COL = "d"
    EMAIL_COL = "i"
    INVOICE_NUMBER_COL = "j"
    FEE_COL = "k"
    START_ROW = 2

    def __init__(self, excel_url, sender_email, sender_password):
        self.eh = ExcelHandler(excel_url)
        self.es = EmailSender(sender_email, sender_password)

        self.start_sending()

    def start_sending(self):
        excel_information = self.get_information_from_excel()

    def get_email_body(self, last_name, invoice_number, fee, year):
        email_body = f"""
Sehr geehrte(r) Frau/Herr {last_name},

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
