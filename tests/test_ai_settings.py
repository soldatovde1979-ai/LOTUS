# -*- coding: utf-8 -*-
"""Тесты настроек AI-провайдера (ai_settings): провайдер читается из mock-конфига,
реальный set.ini не затрагивается."""
import configparser
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_gemini


def make_config(provider):
    cfg = configparser.ConfigParser()
    cfg.add_section("AI")
    cfg.set("AI", "provider", provider)
    return cfg


class AiSettingsTest(unittest.TestCase):
    def setUp(self):
        self._orig_load_config = notes_gemini.load_config

    def tearDown(self):
        notes_gemini.load_config = self._orig_load_config

    def _use(self, cfg):
        notes_gemini.load_config = lambda: cfg

    def test_opencoder_defaults(self):
        self._use(make_config("opencoder"))
        ai = notes_gemini.ai_settings()
        self.assertEqual(ai["provider"], "opencoder")
        self.assertEqual(ai["model"], "pro")
        self.assertEqual(ai["base_url"], "http://127.0.0.1:8787/v1")

    def test_opencoder_from_section(self):
        cfg = make_config("opencoder")
        cfg.add_section("OPENCODER")
        cfg.set("OPENCODER", "model", "flash")
        cfg.set("OPENCODER", "base_url", "http://127.0.0.1:9999/v1")
        cfg.set("OPENCODER", "api_key", "local")
        notes_gemini.load_config = lambda: cfg
        ai = notes_gemini.ai_settings()
        self.assertEqual(ai["provider"], "opencoder")
        self.assertEqual(ai["model"], "flash")
        self.assertEqual(ai["base_url"], "http://127.0.0.1:9999/v1")
        self.assertEqual(ai["api_key"], "local")

    def test_fallback_is_opencoder(self):
        cfg = configparser.ConfigParser()
        cfg.add_section("AI")
        notes_gemini.load_config = lambda: cfg
        ai = notes_gemini.ai_settings()
        self.assertEqual(ai["provider"], "opencoder")

    def test_google_unchanged(self):
        cfg = make_config("google")
        cfg.add_section("GOOGLE")
        cfg.set("GOOGLE", "api_key", "k")
        cfg.set("GOOGLE", "model", "gemini-2.0-flash")
        notes_gemini.load_config = lambda: cfg
        ai = notes_gemini.ai_settings()
        self.assertEqual(ai["provider"], "google")
        self.assertEqual(ai["base_url"], notes_gemini.GEMINI_API_URL)


if __name__ == "__main__":
    unittest.main()
