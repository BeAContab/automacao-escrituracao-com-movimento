import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, args=['--start-maximized'])
        context = await browser.new_context(viewport=None)
        page = await context.new_page()
        await page.goto('https://iss.fortaleza.ce.gov.br/grpfor/home.seam', wait_until='domcontentloaded', timeout=120000)
        await page.screenshot(path='brain/ultima_tela_playwright.png', full_page=True)
        print('BROWSER_READY')
        print(page.url)
        # Mantém a sessão viva para eu poder continuar no próximo passo.
        while True:
            await asyncio.sleep(1)

asyncio.run(main())
