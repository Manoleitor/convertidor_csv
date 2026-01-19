import unittest
from pathlib import Path
from process_signatures import (
    get_participants,
    create_sign_sheet_workbook,
    save_workbook,
    create_sign_pdf_html,
    save_pdf_from_html
)
import os
from openpyxl import load_workbook

class TestProcessSignatures(unittest.TestCase):
    def setUp(self):
        # 53 nombres de Pokémon (sin repeticiones excesivas)
        pokemon_names = [
            "Pikachu", "Bulbasaur", "Charmander", "Squirtle", "Jigglypuff", "Meowth", "Psyduck", "Machop", "Magnemite", "Eevee",
            "Snorlax", "Mewtwo", "Chikorita", "Cyndaquil", "Totodile", "Mareep", "Sudowoodo", "Espeon", "Umbreon", "Murkrow",
            "Wobbuffet", "Larvitar", "Treecko", "Torchic", "Mudkip", "Ralts", "Surskit", "Shroomish", "Makuhita", "Skitty",
            "Sableye", "Mawile", "Meditite", "Swablu", "Barboach", "Bagon", "Beldum", "Turtwig", "Chimchar", "Piplup",
            "Starly", "Kricketot", "Shinx", "Riolu", "Gible", "Hippopotas", "Snover", "Rotom", "Snivy", "Tepig",
            "Oshawott", "Zorua", "Axew"
        ]
        french_surnames = [
            "Dufour", "Lafitte", "Dumont", "Duchamp", "Bourdet", "Bideau", "Lissarrague", "Darrigrand", "Larrieu", "Larronde",
            "Larrarte", "Larrucea", "Etcheto", "Etcheverry", "Etchegoyen", "Etcheberry", "Larralde", "Darracq", "Bidegain", "Larranaga",
            "Lafargue", "Laborde", "Lafon", "Lafargue", "Lafitte", "Lafon", "Lafargue", "Lafitte", "Lafon", "Lafargue"
        ]
        basque_surnames = [
            "Aguirre", "Echeverria", "Goikoetxea", "Ibarra", "Irizar", "Mendieta", "Oteiza", "Urrutia", "Zabala", "Zubizarreta",
            "Arrieta", "Etxeberria", "Garate", "Garmendia", "Goñi", "Iturriaga", "Lasa", "Muguruza", "Olaizola", "Sarasola",
            "Ugalde", "Urkiza", "Zuloaga", "Zunzunegui", "Zubiri", "Zubia", "Zubeldia", "Zubimendi", "Zubizarreta", "Zuloaga", "Zubia"
        ]
        self.test_participants = [
            (pokemon_names[i], f"{french_surnames[i % len(french_surnames)]} {basque_surnames[i % len(basque_surnames)]}") for i in range(53)
        ]
        self.result_dir = Path("./result_test")
        self.result_dir.mkdir(exist_ok=True)
        self.excel_path = self.result_dir / "test_hoja_firmas.xlsx"
        self.pdf_path = self.result_dir / "test_hoja_firmas.pdf"

    def tearDown(self):
        if self.excel_path.exists():
            self.excel_path.unlink()
        if self.pdf_path.exists():
            self.pdf_path.unlink()
        if self.result_dir.exists():
            try:
                self.result_dir.rmdir()
            except Exception:
                pass

    def test_create_sign_sheet_workbook_and_save(self):
        wb = create_sign_sheet_workbook(self.test_participants)
        save_workbook(wb, self.excel_path)
        self.assertTrue(self.excel_path.exists())
        # Comprobar que la hoja tiene el nombre correcto
        loaded = load_workbook(self.excel_path)
        self.assertIn("Firmas", loaded.sheetnames)

    def test_create_sign_pdf_html_and_save(self):
        html = create_sign_pdf_html(self.test_participants, subject="Test", group="T1", week=1, day1="Lunes", day2="Viernes")
        self.assertIn("Hoja de Firmas", html)
        save_pdf_from_html(html, self.pdf_path)
        self.assertTrue(self.pdf_path.exists())

    def test_create_sign_pdf_html_placeholders(self):
        html = create_sign_pdf_html(self.test_participants)
        self.assertIn("primera", html)
        self.assertIn("segunda", html)

    def test_create_sign_sheet_workbook_empty(self):
        wb = create_sign_sheet_workbook([])
        save_workbook(wb, self.excel_path)
        self.assertTrue(self.excel_path.exists())

    def test_create_sign_pdf_html_special_characters(self):
        # Test PDF con caracteres especiales en nombres
        participants = [("José María", "Muñoz-Álvarez"), ("Zoë", "O'Connor")]
        html = create_sign_pdf_html(participants, subject="Español & Français", group="T2", week=2, day1="Lunes", day2="Miércoles")
        self.assertIn("Español & Français", html)
        save_pdf_from_html(html, self.pdf_path)
        self.assertTrue(self.pdf_path.exists())

    def test_create_sign_pdf_html_optional_args(self):
        html = create_sign_pdf_html(self.test_participants)
        self.assertIn("Hoja de Firmas", html)
        save_pdf_from_html(html, self.pdf_path)
        self.assertTrue(self.pdf_path.exists())

    def test_get_participants_empty(self):
        self.assertIsInstance(get_participants(), list)

if __name__ == "__main__":
    unittest.main()
