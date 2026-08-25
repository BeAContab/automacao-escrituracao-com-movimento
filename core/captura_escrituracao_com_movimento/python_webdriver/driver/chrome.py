from .element import (
    WebDriverElement,
    WebDriverElementLocator,
)
from .exceptions import (
    WebDriverException,
    WebDriverNotStartedException,
    WebDriverNotInstantiatedException,
    InvalidFileExtensionException,
)
from ..functions import retry_on_exception
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebElement, WebDriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from loguru import logger
import base64
from pathlib import Path
from enum import Enum
from typing import Any
from datetime import datetime
from time import sleep


class ChromeDriver:
    def __init__(
        self,
        driver_path: str | Path,
        driver_label: str | None = None,
    ):
        """
        Classe Wrapper que envelopa a classe WebDriver do selenium para automações utilizando o navegador Google Chrome

        Args:
            driver_path:
                Caminho para o executável do navegador
            driver_label:
                Label para identificar driver quando rodando mais de uma instância. Utilizado apenas em logs. Padrão é sem label.
        """
        self.driver_path = Path(driver_path)
        self._driver: WebDriver | None = None
        self._is_driver_started = False
        self._label = DriverLabel(driver_label)

    # ------------------------------------------------------------------------------------
    #                                       API Pública
    # ------------------------------------------------------------------------------------

    def is_driver_started(self) -> bool:
        """Retorna True se o Webdriver tiver sido inicializado, False caso contrário."""
        return self._is_driver_started

    def get_driver(self) -> WebDriver:
        """Retorna instância do WebDriver"""
        if self._driver:
            return self._driver
        raise WebDriverNotInstantiatedException(
            "ChromeDriver ainda não foi instanciado."
        )

    def start_driver(
        self,
        options: tuple[str, ...] = tuple(),
        experimental_options: dict[str, Any] | None = None,
    ) -> WebDriver:
        """
        Inicia o navegador localizado em **self.driver_path** e retorna uma instância de WebDriver.
        Args:
            options:
                Tupla de opções
            experimental_options:
                Dicionário com opções experimentais a serem passadas para o Chromium.
        Returns:
            Instância de WebDriver.
        """
        if not self.is_driver_started():
            logger.debug(f"{self._label}Iniciando ChromeDriver...")
            opt = Options()
            opt.binary_location = str(self.driver_path)
            for o in options:
                opt.add_argument(o)
            if experimental_options:
                opt.add_experimental_option("prefs", experimental_options)
            self._driver = webdriver.Chrome(opt)
            self._is_driver_started = True
        return self.get_driver()

    def quit_driver(self):
        """Encerra o navegador e marca o driver como não iniciado.

        Raises:
            WebDriverNotInstantiatedException: Se o ChromeDriver nunca foi instanciado.
        """
        if self.is_driver_started():
            logger.debug(f"{self._label}Fechando ChromeDriver...")
            self.get_driver().quit()
            self._is_driver_started = False

    def goto(self, url: str):
        """Navega para a URL informada.

        Args:
            url: Endereço da página a ser carregada.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
        """
        self._check_webdriver_started()
        logger.debug(f"{self._label}Navegando para {url}...")
        self.get_driver().get(url)

    def find_element(
        self, context: WebDriverElement | None = None
    ) -> WebDriverElementLocator:
        """Retorna um localizador para buscar um único elemento na página.

        Args:
            context: Elemento pai opcional para restringir a busca. Se None,
                a busca ocorre em toda a página.

        Returns:
            Instância de WebDriverElementLocator configurada com o driver atual.

        Raises:
            WebDriverNotInstantiatedException: Se o ChromeDriver ainda não foi
                instanciado.
        """
        return WebDriverElementLocator(self.get_driver(), context)

    def type(
        self,
        element: WebDriverElement,
        text: str,
        max_retries: int = 10,
        timeout: float = 1,
    ):
        """Insere um texto em um elemento da página.

        Args:
            element: Elemento no qual o texto será inserido.
            text: Texto a ser digitado no elemento.
            max_retries: Número máximo de tentativas em caso de falha. Padrão é 10.
            timeout: Intervalo em segundos entre cada tentativa. Padrão é 1.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
        """
        self._check_webdriver_started()
        logger.debug(f"{self._label}Inserindo texto no elemento {element}...")

        def _action():
            el = element.get_element()
            el.send_keys(text)

        retry_on_exception(
            func=_action,
            max_attempts=max_retries,
            polling_seconds=timeout,
        )

    def clear(
        self, element: WebDriverElement, max_retries: int = 10, timeout: float = 1
    ):
        """Limpa o texto de um elemento da página.

        Args:
            element: Elemento cujo conteúdo será limpo.
            max_retries: Número máximo de tentativas em caso de falha. Padrão é 10.
            timeout: Intervalo em segundos entre cada tentativa. Padrão é 1.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
        """
        self._check_webdriver_started()
        logger.debug(f"{self._label}Limpando o elemento {element}...")

        def _action():
            el = element.get_element()
            el.clear()

        retry_on_exception(
            func=_action,
            max_attempts=max_retries,
            polling_seconds=timeout,
        )

    def type_and_enter(
        self,
        element: WebDriverElement,
        text: str,
        max_retries: int = 10,
        timeout: float = 1,
    ):
        """Insere um texto em um elemento da página e pressiona a tecla Enter.

        Args:
            element: Elemento no qual o texto será inserido.
            text: Texto a ser digitado no elemento.
            max_retries: Número máximo de tentativas em caso de falha. Padrão é 10.
            timeout: Intervalo em segundos entre cada tentativa. Padrão é 1.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
        """
        self.type(element, text)
        logger.debug(f"{self._label}Apertando botão Enter...")

        def _action():
            element.get_element().send_keys(Keys.ENTER)

        retry_on_exception(
            func=_action,
            max_attempts=max_retries,
            polling_seconds=timeout,
        )

    def click(
        self, element: WebDriverElement, max_retries: int = 10, timeout: float = 1
    ):
        """Clica em um elemento da página.

        Args:
            element: Elemento a ser clicado.
            max_retries: Número máximo de tentativas em caso de falha. Padrão é 10.
            timeout: Intervalo em segundos entre cada tentativa. Padrão é 1.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
        """
        self._check_webdriver_started()
        logger.debug(f"{self._label}Clicando no elemento {element}...")

        def _action():
            el: WebElement = element.get_element()
            el.click()

        retry_on_exception(
            func=_action,
            max_attempts=max_retries,
            polling_seconds=timeout,
        )

    def scroll_element_into_view(
        self, element: WebDriverElement, max_retries: int = 10, timeout: float = 1
    ):
        """Rola a página até que o elemento especificado fique visível na tela.

        Args:
            element: Elemento até o qual a página será rolada.
            max_retries: Número máximo de tentativas em caso de falha. Padrão é 10.
            timeout: Intervalo em segundos entre cada tentativa. Padrão é 1.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
        """
        self._check_webdriver_started()
        logger.debug(f"{self._label}Rolando até o elemento: {element}")

        def _action():
            self.get_driver().execute_script(
                "arguments[0].scrollIntoView();", element.get_element()
            )

        retry_on_exception(
            func=_action,
            max_attempts=max_retries,
            polling_seconds=timeout,
        )

    def drag_and_drop_by_offset(
        self,
        element: WebDriverElement,
        x_offset: int,
        y_offset: int,
        max_retries: int = 10,
        timeout: float = 1,
    ):
        """Arrasta um elemento da página até a posição especificada pelos offsets X e Y.

        Args:
            element: Elemento a ser arrastado.
            x_offset: Deslocamento horizontal em pixels em relação à posição atual do elemento.
            y_offset: Deslocamento vertical em pixels em relação à posição atual do elemento.
            max_retries: Número máximo de tentativas em caso de falha. Padrão é 10.
            timeout: Intervalo em segundos entre cada tentativa. Padrão é 1.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
        """
        self._check_webdriver_started()
        logger.debug(
            f"{self._label}Arrastando elemento {element} até os offsets X: {x_offset}, Y: {y_offset}..."
        )

        def _action():
            el: WebElement = element.get_element()
            action_chains = ActionChains(self.get_driver())
            (
                action_chains.drag_and_drop_by_offset(
                    el, xoffset=x_offset, yoffset=y_offset
                ).perform()
            )

        retry_on_exception(
            func=_action,
            max_attempts=max_retries,
            polling_seconds=timeout,
        )

    def explicit_wait(self, seconds: float):
        """Aguarda um intervalo fixo de tempo antes de continuar a execução.

        Args:
            seconds: Duração da pausa, em segundos.
        """
        logger.debug(
            f"{self._label}Aguardando explicitamente por {seconds} segundo(s)..."
        )
        sleep(seconds)

    def hide_elements(self, elements: tuple[str, ...]):
        """Oculta, de forma síncrona, elementos da página atual.

            Utiliza JavaScript para alterar o estilo dos elementos, definindo `display: none`.
            O método utilizado é `document.querySelectorAll`, então será escondido **todos** os elementos que correspondam ao seletor CSS informado.

            Args:
                elements: Tupla com seletores CSS dos elementos a serem ocultados.
        """
        for el in elements:
            self.get_driver().execute_script(
                f"document.querySelectorAll('{el}').forEach((el) => el.style.display = 'none')"
            )

    def page_to_pdf(
        self,
        file_path: str | Path,
        file_name: str | None = None,
        hide_elements: tuple[str, ...] = tuple(),
        options: dict[str, Any] = {},
    ) -> bytes:
        """
        Renderiza a página atual do navegador como PDF e salva no caminho especificado em *file_path*.

        Utiliza o comando *Page.printToPDF* do Chrome DevTools Protocol (CDP) para gerar o PDF, portanto é
        compatível apenas com o Chrome.

        Args:
            file_path (str): Caminho completo do diretório de destino (sem o nome do arquivo).
                O diretório é criado automaticamente, caso não exista.
            file_name:
                Nome do arquivo, caso queira dar um nome específico. O nome padrão é o timestamp atual (formato "%Y%m%d_%H%M%S.%f")
            hide_elements:
                Tupla com seletores CSS dos elementos a esconder na página, para que não saiam no PDF.
            options:
                Dicionário com opções passadas ao método Page.printToPDF.
                @see https://chromedevtools.github.io/devtools-protocol/tot/Page/#method-printToPDF para a lista de opções.

        Returns:
            Bytes do arquivo PDF escrito.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
            OSError: Se não for possível criar o diretório ou gravar o arquivo.
            InvalidFileExtensionException caso a extensão do arquivo em *file_path* não for ".pdf".
            WebDriverException caso haja alguma falha na execução do comando Page.printToPDF.
        """
        self._check_webdriver_started()

        p = Path(file_path) if isinstance(file_path, str) else file_path

        if file_name:
            f = Path(file_name)
            expected = [".pdf", ".PDF"]
            if f.suffix not in expected:
                raise InvalidFileExtensionException(
                    path=f, expected=expected, received=f.suffix
                )
        else:
            f = Path(f"{datetime.now().strftime("%Y%m%d_%H%M%S.%f")}.pdf")

        pdf_path = p / f

        logger.debug(f"{self._label}Imprimindo PDF da página no caminho: {p}")

        # Create directory, if it does not exist
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        # Inject CSS to hide elements so they don't appear in the PDF
        for selector in hide_elements:
            self.get_driver().execute_script(
                f"document.querySelectorAll('{selector}').forEach((el) => el.style.display = 'none')"
            )

        # Use CDP to render page as PDF
        pdf_data = self.get_driver().execute_cdp_cmd(
            cmd="Page.printToPDF",
            cmd_args=options,
        )

        if not pdf_data:
            raise WebDriverException("Failed to print page to PDF.")

        # Decode and save
        with open(pdf_path, "wb") as f:
            decoded_data = base64.b64decode(pdf_data["data"])
            f.write(decoded_data)

        return decoded_data

    def save_screenshot_page(self, file_dir: str) -> Path | None:
        """
        Salva screenshot da página atual no caminho especificado em *file_dir*.

        Nota: Se save_screenshot_page for invocado antes que a página carregue propriamente, a imagem sairá vazia (ou toda em branco).

        Args:
            file_dir (str): Caminho completo do diretório de destino (sem o nome do arquivo).
                O diretório é criado automaticamente, caso não exista.

        Returns:
            Caminho do arquivo PNG escrito, ou None caso falhe.

        Raises:
            WebDriverNotStartedException: Se o WebDriver não foi iniciado.
            OSError: Se não for possível criar o diretório ou gravar o arquivo.
        """
        self._check_webdriver_started()
        path = Path(f"{file_dir}/{datetime.now().strftime("%Y%m%d_%H%M%S.%f")}.png")
        path.parent.mkdir(parents=True, exist_ok=True)
        result = self.get_driver().save_screenshot(path)
        if result:
            logger.debug(f"{self._label}Captura de tela sava no caminho: {path}")
            return path
        else:
            logger.error(
                f"{self._label}Falha ao salvar captura de tela no caminho: {path}"
            )
            return None

    def enable_network_throttling(
        self,
        throttling_profile: ChromeNetworkThrottlingProfile,
    ):
        """Habilita limitação de rede via Chrome DevTools Protocol (CDP).

        Args:
            throttling_profile: Perfil de throttling a ser emulado (por
                exemplo, SLOW_2G ou CELLULAR_4G).

        Raises:
            WebDriverNotInstantiatedException: Se o ChromeDriver ainda não foi
                instanciado.
        """
        logger.debug(
            f"{self._label}Habilitando limitação de rede com perfil: {throttling_profile.name}"
        )
        self.get_driver().execute_cdp_cmd("Network.enable", {})
        self.get_driver().execute_cdp_cmd(
            "Network.emulateNetworkConditions", throttling_profile.value
        )

    # ------------------------------------------------------------------------------------
    #                                       API Interna
    # ------------------------------------------------------------------------------------

    def _check_webdriver_started(self) -> None:
        """Levanta WebDriverNotStartedException caso o Webdriver não tenha sido inicializado.
        :return:
        """
        if self.get_driver() is None or not self.is_driver_started():
            raise WebDriverNotStartedException("ChromeDriver ainda não foi iniciado.")


# ------------------------------------------------------------------------------------
#                                          Outras Classes
# ------------------------------------------------------------------------------------


class DriverLabel:
    """Label para identificação de drivers quando utilizando múltiplas instâncias. Usado apenas para logging."""

    def __init__(self, label: str | None = None):
        self.label = label

    def __repr__(self) -> str:
        return "" if not self.label else f"({self.label}) "


# ------------------------------------------------------------------------------------
#                                              Enums
# ------------------------------------------------------------------------------------


class ChromeNetworkThrottlingProfile(Enum):
    OFFLINE = {
        "offline": True,
        "downloadThroughput": 0,
        "uploadThroughput": 0,
        "latency": 0,
    }
    SLOW_2G = {
        "offline": False,
        "downloadThroughput": 250 * 1024 / 8,
        "uploadThroughput": 50 * 1024 / 8,
        "latency": 2000,
    }
    CELLULAR_2G = {
        "offline": False,
        "downloadThroughput": 450 * 1024 / 8,
        "uploadThroughput": 150 * 1024 / 8,
        "latency": 300,
    }
    CELLULAR_3G = {
        "offline": False,
        "downloadThroughput": 750 * 1024 / 8,
        "uploadThroughput": 250 * 1024 / 8,
        "latency": 100,
    }
    CELLULAR_4G = {
        "offline": False,
        "downloadThroughput": 4 * 1024 * 1024 / 8,
        "uploadThroughput": 3 * 1024 * 1024 / 8,
        "latency": 20,
    }
