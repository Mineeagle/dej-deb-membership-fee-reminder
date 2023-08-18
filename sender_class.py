from helper_classes import ExcelHandler, EmailSender


class Sender:
    def __init__(self, excel_url, sender_email, sender_password):
        self.eh = ExcelHandler(excel_url)
        self.es = EmailSender(sender_email, sender_password)

        self.start_sending()

    def start_sending():
        pass
    
    def get_email_body(self, last_name, invoice_number, amt, year):
        email_body=f"""
Sehr geehrte(r) Frau/Herr {last_name},

uns ist aufgefallen, dass Sie Ihren Mitgliedsbeitrag für die DEB/DEJ für das Jahr {year} noch nicht gezahlt haben.
Aus diesem Grund bitten wir Sie den Mitgliedsbeitrag von {amt}€ unter der Rechnungsnummer {invoice_number} auf folgendes Konto zu überweisen:

[Daten eines Kontos]

Sollten Sie der Zahlungsaufforderung nicht nachkommen, sehen wir uns dazu verpflichtet anderweitige Maßnahmen einzuleiten.

Mit freundlichen Grüßen,

Deutscher Esperanto-Bund/Deutsche Esperanto-Jugend
"""
        return email_body

    def get_email_subject(self, invoice_number):
        return f"Zahlungsaufforderung DEB/DEJ {invoice_number}"
    


s = Sender()