from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import os
import shutil

class BrowserFactory:
    @staticmethod
    def create_chrome(download_dir=None):
        print("🔧 Configurando opções do Chrome...")
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--start-maximized")
        #chrome_options.add_argument("--headless=new")
        
        # Adicionar argumentos para resolver problemas comuns
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        if download_dir:
            prefs = {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "safebrowsing.disable_download_protection": True,
                "plugins.always_open_pdf_externally": True
            }
            chrome_options.add_experimental_option("prefs", prefs)
        
        print("🔍 Procurando ChromeDriver...")
        
        # Tentar encontrar chromedriver no PATH primeiro
        driver_path = shutil.which("chromedriver")
        if driver_path:
            print(f"✅ ChromeDriver encontrado no PATH: {driver_path}")
            service = Service(driver_path)
        else:
            print("⚠️ ChromeDriver não encontrado no PATH. Deixando Selenium tentar automaticamente...")
            # Deixar o Selenium tentar encontrar automaticamente
            service = Service()
        
        print("🚀 Criando instância do Chrome...")
        try:
            driver = webdriver.Chrome(service=service, options=chrome_options)
            print("✅ Chrome criado com sucesso!")
            return driver
        except Exception as e:
            print(f"❌ Erro ao criar Chrome: {e}")
            print("💡 Dica: Certifique-se de que o Google Chrome está instalado e atualizado")
            raise