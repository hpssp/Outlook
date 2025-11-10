from playwright.sync_api import sync_playwright
import os
import datetime
import time
import openpyxl
from openpyxl.styles import Font, PatternFill

with sync_playwright() as pw:
        resultados = {}
        robo = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36 Edg/140.0.0.0'
        browser = pw.chromium.launch(channel="msedge", headless=False)  # True= sem tela, False= com tela
        context = browser.new_context(user_agent=robo)
        page = context.new_page()
        page.goto("https://outlook.office.com/mail/", timeout=60000)
        page.wait_for_load_state('load')
        page.locator('input[type="email"]').click()
        page.locator('input[type="email"]').fill('hugo.silva.ciee@valor.com.br')
        page.locator('input[type="submit"]').click()
        page.locator('input[type="password"]').fill('40567997Hg10@')
        page.locator('input[type="submit"]').click()

        time.sleep(60)
        context.close()
        browser.close()