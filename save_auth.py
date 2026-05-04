"""
Rode este script localmente para salvar a sessao autenticada do eLaw Anima.

Como usar:
  python save_auth.py

O browser vai abrir. Faca login normalmente (MFA incluido).
Quando a pagina inicial do eLaw carregar, volte aqui e pressione Enter.
O arquivo auth_state.json sera salvo nesta pasta — envie-o para ser
carregado no Render pelo endpoint /api/upload-auth ou pela interface.

ATENCAO: o auth_state.json contem cookies de sessao (nao contem senha).
Trate-o como informacao sigilosa e nao commite no git.
"""
import asyncio
from pathlib import Path

from playwright.async_api import async_playwright

URL = "https://anima.elaw.com.br/"
OUTPUT = Path(__file__).resolve().parent / "auth_state.json"


async def main():
    print("=== Salvar Sessao eLaw Anima ===\n")
    print("1. O browser vai abrir em modo normal (nao headless).")
    print("2. Faca login com usuario, senha e MFA normalmente.")
    print("3. Quando a pagina inicial do eLaw carregar, volte aqui.")
    print("4. Pressione Enter para salvar a sessao.\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=100)
        ctx = await browser.new_context()
        page = await ctx.new_page()

        print("Abrindo eLaw Anima...")
        await page.goto(URL, timeout=60000)

        input("Pressione Enter apos o login estar completo... ")

        await ctx.storage_state(path=str(OUTPUT))
        print(f"\nSessao salva em: {OUTPUT}")
        print("Envie este arquivo para ser carregado no Render antes de rodar o RPA.")
        print("Validade: geralmente 1-7 dias dependendo da politica do Azure AD da Viseu.")

        await browser.close()


asyncio.run(main())
