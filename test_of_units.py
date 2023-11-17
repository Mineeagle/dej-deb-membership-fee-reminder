import pytest
from unittest.mock import patch
import os

"""Test of class Sender"""
from sender_class import Sender
from helper_classes import ExcelHandler


class Test_Sender:
    @classmethod
    def setup_class():
        Test_Sender.dummy_data = [
            {
                Sender.EMAIL_COL: "ulrike.kratschmann@esperanto.de",
                Sender.LAST_NAME_COL: "Kratschmann",
                Sender.INVOICE_NUMBER_COL: "609870",
                Sender.FEE_COL: "30",
                Sender.COMPANY_COL: None,
            },
            {
                Sender.EMAIL_COL: "rainer.zufall@esperanto.de",
                Sender.LAST_NAME_COL: "Zufall",
                Sender.INVOICE_NUMBER_COL: "609871",
                Sender.FEE_COL: "50",
                Sender.COMPANY_COL: None,
            },
            {
                Sender.EMAIL_COL: None,
                Sender.LAST_NAME_COL: "Michaelis",
                Sender.INVOICE_NUMBER_COL: "609872",
                Sender.FEE_COL: "1",
                Sender.COMPANY_COL: None,
            },
            {
                Sender.EMAIL_COL: "firma.grunnings@esperanto.de",
                Sender.LAST_NAME_COL: None,
                Sender.INVOICE_NUMBER_COL: "609873",
                Sender.FEE_COL: "42",
                Sender.COMPANY_COL: "Grunnings",
            },
        ]

    @patch("sender_class.Sender.start_sending")
    @patch("sender_class.print")
    def setup_method(self, method, mock_print, mock_start_sending):
        self.sender = Sender(
            "example_sheet.xlsx",
            "vorname.nachname@esperanto.de",
            "password",
            1984,
        )

    def teardown_method(self, method):
        max_row = self.sender.eh.get_max_row()
        check_col = self.sender.CHECK_COL

        for row in range(1, max_row+1):
            self.sender.eh.write(f"{check_col}{row}", None)
        
        del self.sender

    def test_get_information_from_excel(self):
        sender_res = self.sender.get_information_from_excel()
        for index, entry in enumerate(Test_Sender.dummy_data):
            assert entry[Sender.EMAIL_COL] == sender_res[index][Sender.EMAIL_COL]
            assert (
                entry[Sender.LAST_NAME_COL] == sender_res[index][Sender.LAST_NAME_COL]
            )
            assert (
                entry[Sender.INVOICE_NUMBER_COL]
                == sender_res[index][Sender.INVOICE_NUMBER_COL]
            )
            assert entry[Sender.FEE_COL] == sender_res[index][Sender.FEE_COL]
            assert entry[Sender.COMPANY_COL] == sender_res[index][Sender.COMPANY_COL]

    def test_get_email_subject(self):
        subject = (
            lambda invoice_number: f"Zahlungsaufforderung DEB/DEJ {invoice_number}"
        )
        for entry in Test_Sender.dummy_data:
            assert self.sender.get_email_subject(
                entry[Sender.INVOICE_NUMBER_COL]
            ) == subject(entry[Sender.INVOICE_NUMBER_COL])
        assert self.sender.get_email_subject(123) == subject(123)
        assert self.sender.get_email_subject("123") == subject(123)
        assert self.sender.get_email_subject("eĥoŝanĝo ĉiuĵaŭde") == subject(
            "eĥoŝanĝo ĉiuĵaŭde"
        )
        assert self.sender.get_email_subject(None) == subject(None)

    def test_get_email_body_company(self):
        body = (
            lambda company, invoice_number, fee, year: f"""Sehr geehrte Damen und Herren des Unternehmens {company},

uns ist aufgefallen, dass Sie Ihren Mitgliedsbeitrag für die DEB/DEJ für das Jahr {year} noch nicht gezahlt haben.
Aus diesem Grund bitten wir Sie den Mitgliedsbeitrag von {fee}€ unter der Rechnungsnummer {invoice_number} auf folgendes Konto zu überweisen:

[Daten eines Kontos]

Wir bedanken uns herzlich, dass Sie als Mitglied mit Ihrem Beitrag die Esperanto-Kultur in Deutschland unterstützen.

Mit freundlichen Grüßen,

Deutscher Esperanto-Bund/Deutsche Esperanto-Jugend
"""
        )

        for entry in Test_Sender.dummy_data:
            if entry[Sender.COMPANY_COL] is None:
                continue
            method_res = self.sender.get_email_body_company(
                entry[Sender.COMPANY_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                self.sender.year,
            )
            lambda_res = body(
                entry[Sender.COMPANY_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                self.sender.year,
            )
            assert lambda_res == method_res

    def test_get_email_body_person(self):
        body = (
            lambda last_name, invoice_number, fee, year: f"""Sehr geehrte(r) Frau/Herr {last_name},

uns ist aufgefallen, dass Sie Ihren Mitgliedsbeitrag für die DEB/DEJ für das Jahr {year} noch nicht gezahlt haben.
Aus diesem Grund bitten wir Sie den Mitgliedsbeitrag von {fee}€ unter der Rechnungsnummer {invoice_number} auf folgendes Konto zu überweisen:

[Daten eines Kontos]

Wir bedanken uns herzlich, dass Sie als Mitglied mit Ihrem Beitrag die Esperanto-Kultur in Deutschland unterstützen.

Mit freundlichen Grüßen,

Deutscher Esperanto-Bund/Deutsche Esperanto-Jugend
"""
        )

        for entry in Test_Sender.dummy_data:
            if entry[Sender.LAST_NAME_COL] is None:
                continue
            method_res = self.sender.get_email_body_person(
                entry[Sender.LAST_NAME_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                self.sender.year,
            )
            lambda_res = body(
                entry[Sender.LAST_NAME_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                entry[Sender.INVOICE_NUMBER_COL],
                self.sender.year,
            )
            assert lambda_res == method_res