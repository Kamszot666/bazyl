#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Testy jednostkowe dla funkcji ekstrakcji danych w skrypt.py."""

import unittest
from unittest.mock import patch, MagicMock
from bs4 import BeautifulSoup

# Import klasy scrapera - wymagane obejscie importu selenium
import sys

# Mock selenium i webdriver_manager zanim zaimportujemy skrypt
selenium_mock = MagicMock()
sys.modules["selenium"] = selenium_mock
sys.modules["selenium.webdriver"] = selenium_mock.webdriver
sys.modules["selenium.webdriver.chrome"] = selenium_mock.webdriver.chrome
sys.modules["selenium.webdriver.chrome.service"] = selenium_mock.webdriver.chrome.service
sys.modules["selenium.webdriver.chrome.options"] = selenium_mock.webdriver.chrome.options
sys.modules["selenium.webdriver.common"] = selenium_mock.webdriver.common
sys.modules["selenium.webdriver.common.by"] = selenium_mock.webdriver.common.by
sys.modules["selenium.webdriver.common.keys"] = selenium_mock.webdriver.common.keys
sys.modules["selenium.webdriver.support"] = selenium_mock.webdriver.support
sys.modules["selenium.webdriver.support.ui"] = selenium_mock.webdriver.support.ui
sys.modules["selenium.webdriver.support.expected_conditions"] = (
    selenium_mock.webdriver.support.expected_conditions
)
sys.modules["selenium.common"] = selenium_mock.common
sys.modules["selenium.common.exceptions"] = selenium_mock.common.exceptions
sys.modules["webdriver_manager"] = MagicMock()
sys.modules["webdriver_manager.chrome"] = MagicMock()

from skrypt import SkryptScraper


class TestWyodrebnijFacebook(unittest.TestCase):
    """Testy ekstrakcji linkow Facebook."""

    def test_link_facebook(self):
        html = '<html><body><a href="https://facebook.com/firma123">FB</a></body></html>'
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_facebook(soup)
        self.assertEqual(wynik, "https://facebook.com/firma123")

    def test_brak_facebooka(self):
        html = "<html><body><a href=\"https://twitter.com/firma\">TW</a></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_facebook(soup)
        self.assertIsNone(wynik)


class TestWyodrebnijLinkedin(unittest.TestCase):
    """Testy ekstrakcji linkow LinkedIn."""

    def test_link_linkedin(self):
        html = '<html><body><a href="https://linkedin.com/company/firma">LI</a></body></html>'
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_linkedin(soup)
        self.assertEqual(wynik, "https://linkedin.com/company/firma")

    def test_brak_linkedin(self):
        html = "<html><body><p>Brak linkow</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_linkedin(soup)
        self.assertIsNone(wynik)


class TestWyodrebnijEmail(unittest.TestCase):
    """Testy ekstrakcji adresow email."""

    def test_mailto_link(self):
        html = '<html><body><a href="mailto:kontakt@firma.pl">Mail</a></body></html>'
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_email(soup)
        self.assertEqual(wynik, "kontakt@firma.pl")

    def test_email_w_tekscie(self):
        html = "<html><body><p>Napisz do nas: biuro@firma.com</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_email(soup)
        self.assertEqual(wynik, "biuro@firma.com")

    def test_priorytet_kontakt(self):
        html = (
            "<html><body>"
            "<p>admin@firma.pl kontakt@firma.pl</p>"
            "</body></html>"
        )
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_email(soup)
        self.assertEqual(wynik, "kontakt@firma.pl")

    def test_filtracja_wykluczone(self):
        html = "<html><body><p>test@example.com</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_email(soup)
        self.assertIsNone(wynik)

    def test_brak_emaila(self):
        html = "<html><body><p>Brak danych</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_email(soup)
        self.assertIsNone(wynik)


class TestWyodrebnijTelefon(unittest.TestCase):
    """Testy ekstrakcji numerow telefonu."""

    def test_link_tel(self):
        html = '<html><body><a href="tel:+48123456789">Zadzwon</a></body></html>'
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_telefon(soup)
        self.assertIsNotNone(wynik)
        self.assertIn("123456789", wynik)

    def test_numer_w_tekscie(self):
        html = "<html><body><p>Tel: 123 456 789</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_telefon(soup)
        self.assertIsNotNone(wynik)

    def test_brak_telefonu(self):
        html = "<html><body><p>Brak numeru</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_telefon(soup)
        self.assertIsNone(wynik)


class TestWyodrebnijNip(unittest.TestCase):
    """Testy ekstrakcji NIP."""

    def test_nip_z_etykieta(self):
        html = "<html><body><p>NIP: 1234567890</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_nip(soup)
        self.assertEqual(wynik, "1234567890")

    def test_nip_z_myslnikami(self):
        html = "<html><body><p>NIP: 123-456-78-90</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_nip(soup)
        self.assertEqual(wynik, "1234567890")

    def test_brak_nip(self):
        html = "<html><body><p>Firma XYZ</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        wynik = SkryptScraper.wyodrebnij_nip(soup)
        self.assertIsNone(wynik)


class TestWyodrebnijOsobeKontaktowa(unittest.TestCase):
    """Testy ekstrakcji osoby kontaktowej."""

    def test_osoba_w_sekcji_kontakt(self):
        html = (
            "<html><body>"
            '<div class="kontakt"><p>Osoba kontaktowa: Jan Nowak</p></div>'
            "</body></html>"
        )
        soup = BeautifulSoup(html, "html.parser")
        imie, stanowisko = SkryptScraper.wyodrebnij_osobe_kontaktowa(soup)
        self.assertEqual(imie, "Jan Nowak")

    def test_brak_osoby(self):
        html = "<html><body><p>Informacje o firmie</p></body></html>"
        soup = BeautifulSoup(html, "html.parser")
        imie, stanowisko = SkryptScraper.wyodrebnij_osobe_kontaktowa(soup)
        self.assertIsNone(imie)


class TestHelpers(unittest.TestCase):
    """Testy funkcji pomocniczych."""

    def test_zapisz_wczytaj_checkpoint(self):
        from skrypt import zapisz_checkpoint, wczytaj_checkpoint, SCIEZKA_CHECKPOINT

        zapisz_checkpoint(42)
        wynik = wczytaj_checkpoint()
        self.assertEqual(wynik, 42)

        # Cleanup
        if SCIEZKA_CHECKPOINT.exists():
            SCIEZKA_CHECKPOINT.unlink()

    def test_zapisz_log_csv(self):
        from skrypt import zapisz_log_csv, SCIEZKA_LOGU

        zapisz_log_csv("Test zdarzenie")
        self.assertTrue(SCIEZKA_LOGU.exists())

        with open(SCIEZKA_LOGU, "r", encoding="utf-8") as f:
            tresc = f.read()
        self.assertIn("Test zdarzenie", tresc)

        # Cleanup
        if SCIEZKA_LOGU.exists():
            SCIEZKA_LOGU.unlink()


if __name__ == "__main__":
    unittest.main()
