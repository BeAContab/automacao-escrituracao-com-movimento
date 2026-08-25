from selenium.webdriver.common.by import By, ByType
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webdriver import WebElement


class WebDriverElement:
    """Representa um elemento da página web localizado por um seletor do Selenium.

    Attributes:
        locator: Tupla contendo o tipo de seletor e o valor utilizado para localizar o elemento.
    """

    def __init__(self, locator: tuple[ByType, str], driver: WebDriver, context: WebDriverElement | None = None):
        """
        Args:
            locator: Tupla contendo o tipo de seletor (ex: By.ID) e o valor do seletor.
            driver: Instância do WebDriver utilizada para localizar o elemento.
            context: Contexto de busca para localizar o elemento.
            Se passado, todas as buscas serão feitas em cima deste elemento.
        """
        self.locator = locator
        self.context = context
        self._driver = driver

    def __repr__(self):
        return f"({self.locator[0]}) {self.locator[1]}"

    def get_element(self) -> WebElement:
        """Localiza e retorna o primeiro elemento correspondente ao localizador.

        Returns:
            Instância de WebElement correspondente ao localizador.
        """
        if self.context:
            return self.context.get_element().find_element(by=self.locator[0], value=self.locator[1])
        return self._driver.find_element(by=self.locator[0], value=self.locator[1])

    def get_all_elements(self) -> list[WebElement]:
        """Localiza e retorna todos os elementos correspondentes ao localizador.

        Returns:
            Lista de WebElements correspondentes ao localizador.
        """
        if self.context:
            return self.context.get_element().find_elements(by=self.locator[0], value=self.locator[1])
        return self._driver.find_elements(by=self.locator[0], value=self.locator[1])

    def get_coordinates(self) -> WebDriverElementCoordinates:
        """Retorna as coordenadas do elemento na página.

        Returns:
            Instância de WebDriverElementCoordinates com as coordenadas X e Y do elemento.
        """
        return WebDriverElementCoordinates(coords=self.get_element().location)


class WebDriverElementLocator:
    """Fábrica de WebDriverElements, responsável por localizar elementos na página web.

    Attributes:
        context: Elemento pai opcional utilizado para restringir a busca.
    """

    def __init__(self, driver: WebDriver, context: WebDriverElement | None = None):
        """
        Args:
            driver: Instância do WebDriver utilizada para localizar elementos.
            context: Elemento pai opcional para restringir a busca. Se None,
                a busca ocorre em toda a página.
        """
        self._driver = driver
        self.context = context

    def by_attribute(self, attr_name: str, attr_value: str, context: WebDriverElement | None = None) -> WebDriverElement:
        """Localiza um elemento pelo valor de um atributo HTML via XPath.

        Args:
            attr_name: Nome do atributo HTML a ser utilizado na busca.
            attr_value: Valor esperado do atributo.

        Returns:
            Instância de WebDriverElement configurada com o localizador XPath gerado.
        """
        return WebDriverElement(
            locator=(By.XPATH, f'//*[@{attr_name}="{attr_value}"]'), driver=self._driver, context=self.context
        )

    def by_tag(self, tag_name: str, context: WebDriverElement | None = None) -> WebDriverElement:
        """Localiza um elemento pelo nome da tag HTML.

        Args:
            tag_name: Nome da tag HTML a ser buscada (ex: "input", "div").

        Returns:
            Instância de WebDriverElement configurada com o localizador por tag.
        """
        return WebDriverElement(locator=(By.TAG_NAME, tag_name), driver=self._driver, context=self.context)

    def by_id(self, element_id: str, context: WebDriverElement | None = None) -> WebDriverElement:
        """Localiza um elemento pelo atributo id.

        Args:
            element_id: Valor do atributo id do elemento.

        Returns:
            Instância de WebDriverElement configurada com o localizador por id.
        """
        return WebDriverElement(locator=(By.ID, element_id), driver=self._driver, context=self.context)

    def by_class(self, class_name: str, context: WebDriverElement | None = None):
        """Localiza um elemento pelo nome da classe CSS.

        Args:
            class_name: Nome da classe CSS do elemento.

        Returns:
            Instância de WebDriverElement configurada com o localizador por classe.
        """
        return WebDriverElement(
            locator=(By.CLASS_NAME, class_name), driver=self._driver, context=self.context
        )

    def by_name(self, name: str, context: WebDriverElement | None = None):
        """Localiza um elemento pelo atributo name.

        Args:
            name: Valor do atributo name do elemento.

        Returns:
            Instância de WebDriverElement configurada com o localizador por name.
        """
        return WebDriverElement(locator=(By.NAME, name), driver=self._driver, context=self.context)

    def by_xpath(self, xpath: str, context: WebDriverElement | None = None) -> WebDriverElement:
        """Localiza um elemento por uma expressão XPath.

        Args:
            xpath: Expressão XPath utilizada para localizar o elemento.

        Returns:
            Instância de WebDriverElement configurada com o localizador XPath.
        """
        return WebDriverElement(locator=(By.XPATH, xpath), driver=self._driver, context=self.context)

    def by_link_text(self, text: str, context: WebDriverElement | None = None) -> WebDriverElement:
        """Localiza um elemento de link pelo seu texto visível.

        Args:
            text: Texto visível do link a ser buscado.

        Returns:
            Instância de WebDriverElement configurada com o localizador por texto de link.
        """
        return WebDriverElement(locator=(By.LINK_TEXT, text), driver=self._driver, context=self.context)

    def by_css_selector(self, selector: str, context: WebDriverElement | None = None) -> WebDriverElement:
        """Localiza um elemento por um seletor CSS.

        Args:
            selector: Seletor CSS utilizado para localizar o elemento.

        Returns:
            Instância de WebDriverElement configurada com o localizador por seletor CSS.
        """
        return WebDriverElement(
            locator=(By.CSS_SELECTOR, selector), driver=self._driver, context=self.context
        )


class WebDriverElementCoordinates:
    """Representa as coordenadas X e Y de um elemento na página web.

    Attributes:
        x: Coordenada horizontal do elemento em pixels.
        y: Coordenada vertical do elemento em pixels.
    """

    def __init__(self, coords: dict[str, int]):
        """
        Args:
            coords: Dicionário contendo as coordenadas do elemento,
                esperando as chaves "x" e "y" com valores inteiros em pixels.
        """
        self.x, self.y = coords.values()
