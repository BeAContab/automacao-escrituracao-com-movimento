
from core.captura_escrituracao_com_movimento.python_spreadsheet_reader.readers.xlsx import XLSXReader
from core.captura_escrituracao_com_movimento.python_webdriver.driver.chrome import ChromeDriver
from core.captura_escrituracao_com_movimento.python_webdriver.functions import retry_on_exception
from core.captura_escrituracao_com_movimento.functions import (
    aplicar_mascara_cnpj,
    remover_pontuacao_cpf_cnpj,
    remover_caracteres_escape,
    timestamp_as_file_name,
    verifica_arquivo_existe,
    safe_cast
)
from core.captura_escrituracao_com_movimento.classes import (
    Empresa,
    MensagemISSFortaleza,
    ThreadStoppedException,
)
from selenium.common import (
    TimeoutException,
    ElementNotInteractableException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from pydantic import SecretStr, ValidationError
from collections.abc import Callable
from datetime import date, datetime
from pathlib import Path
from loguru import logger
import shutil
import tempfile


def baixar_certificado_escrituracao_empresas_iss(
    url_iss_fortaleza: str,
    cpf_iss_fortaleza: SecretStr,
    senha_iss_fortaleza: SecretStr,
    competencia: date,
    encerramento: date,
    caminho_planilha_fiscal: Path,
    caminho_webdriver: Path,
    diretorio_saida: Path,
    check_thread_stopped_cb: Callable[[], None],
    callback_progresso: Callable[[int, int], None] | None = None
):
    # Iniciar webdriver
    driver = ChromeDriver(
        driver_path=caminho_webdriver,
    )

    try:
        # Usa uma pasta temporária real (fora do diretório de instalação) para os downloads
        # intermediários do Chrome — o diretório de instalação pode estar em Program Files,
        # sem permissão de escrita para o operador (mesmo motivo pelo qual app.py grava
        # configurações em %APPDATA%, não ao lado do executável).
        temp_path = tempfile.mkdtemp(prefix="captura_iss_com_movimento_")
        driver.start_driver(
            options=("--start-maximized",),
            experimental_options={
                "download.default_directory": temp_path,  # Arquivos serão baixados aqui
                "download.prompt_for_download": False,    # não perguntar onde baixar
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True              # evita bloqueio de downloads
            }
        )

        temp_path = Path(temp_path)  # Path é mais fácil de trabalhar com diretórios

        # * ---------------------------------- Login Form -----------------------------------

        logger.info("Estou entrando no portal da ISS Fortaleza com o CPF e a senha informados...")

        # Navegar ao site do ISS
        driver.goto(url_iss_fortaleza)

        driver.click(
            driver.find_element().by_attribute(attr_name="href", attr_value="/grpfor/oauth2/login")
        )

        # CNPJ
        driver.type(
            element=driver.find_element().by_id("username"),
            text=cpf_iss_fortaleza.get_secret_value()
        )

        # Senha
        driver.type(
            element=driver.find_element().by_id("password"),
            text=senha_iss_fortaleza.get_secret_value()
        )

        check_thread_stopped_cb()

        # Click no botão de login
        driver.find_element().by_id("kc-form-login").get_element().submit()

        # * -------------------------- Home / Modal de Busca de Empresas ---------------------------

        # Ler Planilha Fiscal
        reader = XLSXReader(
            workbook_path=Path(caminho_planilha_fiscal)
        )

        # * Nota: planilha fiscal do ISS tem muitos registros. Não tente carregar dados sem usar lazy_load.
        dados_planilha = reader.read_sheet(
            sheet_name=f"{competencia.month:02d}.{competencia.year}",
            lazy_load=True,  # Retorna Generator
            cell_values_only=True,
            read_locked=True  # Planilha é compartilhada na rede, outros usuários podem ter ela aberta.
        )

        # Contadores
        empresas_processadas: list[Empresa] = []
        qtd_empresas_ignoradas: int = 0
        qtd_empresas__sem_procuracao: int = 0
        qtd_empresas_erro: int = 0

        for idx, row in enumerate(dados_planilha):
            row_number = idx + 1
            coord_cod_empresa = f"A{row_number}"

            # Validações
            if coord_cod_empresa not in row:  # Pode existir um gap entre as chaves em row.
                continue

            cod_empresa = row[coord_cod_empresa]

            if not cod_empresa or (isinstance(cod_empresa, str) and cod_empresa.replace(" ", "") == ""):
                break  # Se cod_empresa é vazio ou nulo é porque chegou no fim dos dados relevantes da planilha

            municipio = row[f"W{row_number}"]

            if (not municipio) or (isinstance(municipio, str) and municipio.lower() != "fortaleza"):
                # Ignorar empresas que não são de Fortaleza
                qtd_empresas_ignoradas += 1
                continue

            dt_encerramento_planilha = row[f"Y{row_number}"]  # Campo vem da planilha como datetime

            if not isinstance(dt_encerramento_planilha, datetime):
                # Campo não é datetime ou não está preenchido
                continue

            dt_encerramento_planilha = dt_encerramento_planilha.date()  # Converter para date

            # Comparar com data de encerramento
            if dt_encerramento_planilha != encerramento:
                # Apenas nos interessa empresas com ISS encerrado na data em "encerramento"
                continue

            responsavel = row[f"H{row_number}"]
            socio_ativo_reinf = row[f"I{row_number}"] or ""
            responsavel_separacao = row[f"M{row_number}"] or ""

            if not (isinstance(responsavel, str) or isinstance(socio_ativo_reinf, str) or isinstance(responsavel_separacao, str)):
                logger.error("Não consegui descobrir se essa empresa é com ou sem movimento — os campos da planilha que uso para isso vieram vazios ou em formato inesperado. Vou pular para a próxima.")
                qtd_empresas_erro += 1
                continue

            try:
                empresa = Empresa(
                    codigo=safe_cast(cod_empresa, int),
                    nome=row[f"B{row_number}"],
                    cnpj=row[f"D{row_number}"],
                    responsavel=responsavel.upper(),
                    socio_ativo_reinf=socio_ativo_reinf,
                    responsavel_separacao=responsavel_separacao,
                )
            except ValidationError as v_err:
                logger.error(f"Os dados da empresa na linha {row_number} da planilha não bateram com o esperado, então vou pular essa linha. Detalhe do problema: {v_err}")
                qtd_empresas_erro += 1
                continue

            if empresa.eh_sem_movimento(): # Empresas sem movimento não nos interessa
                logger.info(f"A empresa {empresa} está marcada como sem movimento — essa não é a lista que estou processando agora, então vou seguir para a próxima.")
                qtd_empresas_ignoradas += 1
                continue

            logger.info(f"Agora é a vez da empresa {empresa}: vou procurar a inscrição dela no portal e baixar a escrituração.")

            check_thread_stopped_cb()

            # Selecionar a opção CNPJ
            driver.click(
                driver.find_element().by_xpath('//table[@id="alteraInscricaoForm:tipoPesquisa"]//input[@value="CNPJ"]'),
            )

            # Após clicar na opção CNPJ, é realizado uma chamada AJAX para aplicar máscara de CPF/CNPJ
            # o que faz com que a caixa de pesquisa re-renderize e o texto digitado seja apagado, se inserido rápido demais.
            driver.explicit_wait(1)

            cnpj_sem_pontuacao = remover_pontuacao_cpf_cnpj(empresa.cnpj.strip())

            driver.clear(driver.find_element().by_id("alteraInscricaoForm:cpfPesquisa"))
            driver.type(driver.find_element().by_id("alteraInscricaoForm:cpfPesquisa"), cnpj_sem_pontuacao)

            driver.explicit_wait(1)

            # * Nota: antes de apertar o botão "Pesquisar", temos que verificar se a empresa já está na barra de resultados.
            # * Se a empresa já estiver listada, é necessário esperar um pouco mais após apertar o botão "Pesquisar".
            bln_empresa_listada = False

            record_xpath = f'//tbody[@id="alteraInscricaoForm:empresaDataTable:tb"]//a[normalize-space()="{aplicar_mascara_cnpj(cnpj_sem_pontuacao)}"]'

            # Verificar se empresa já está listada
            if driver.find_element().by_xpath(record_xpath).get_all_elements():
                bln_empresa_listada = True

            # Clicar no botão "Pesquisar"
            driver.click(
                driver.find_element().by_id("alteraInscricaoForm:btnPesquisar")
            )

            try:
                if bln_empresa_listada:
                    driver.explicit_wait(5)

                # Procurar a linha da tabela de resultados que tem o CNPJ (com máscara) da empresa
                driver.click(
                    driver.find_element().by_xpath(record_xpath),
                    max_retries=5
                )
            except (TimeoutException, NoSuchElementException):
                logger.warning(
                    f"Não foi possível encontrar a inscrição da empresa {empresa}"
                )
                qtd_empresas__sem_procuracao += 1
                continue

            try:
                # * -------------------------- Modal de Mudança de Empresa ---------------------------
                driver.click(  # levanta exceção se modal não aparecer
                    driver.find_element().by_id("alteraInscricaoForm:botaoOk"),
                    max_retries=5
                )
            except (ElementNotInteractableException, NoSuchElementException, StaleElementReferenceException):
                pass

            try:
                # * ------------------------------- Modal de Mensagens Não-Lidas --------------------------------
                def _wait_for_messages_modal_to_popup():
                    return (driver
                          .find_element()
                          .by_xpath("//*[@id='mensagensForm:mensagemDataTable:tb']//tr")
                          .get_element())

                retry_on_exception(  # Se não tiver mensagens levanta exception, pulando o processo abaixo
                    func=_wait_for_messages_modal_to_popup,
                    max_attempts=5,
                    polling_seconds=1,
                    exception=NoSuchElementException
                )

                logger.info("Essa empresa tem mensagens não lidas no portal — vou abrir cada uma e dar ciência antes de continuar.")

                # driver.explicit_wait(1)

                # Mover modal 200 pixels para cima porque o programa quebra se tentar clicar em um
                # elemento fora de vista.
                driver.drag_and_drop_by_offset(
                    driver.find_element().by_id("mensagensModalHeader"),
                    x_offset=0,
                    y_offset=-200,
                )

                check_thread_stopped_cb()

                # Verificar se o modal de mensagens pendentes apareceu
                # Se tem mensagens, dar ciência em todas
                messages = (driver
                          .find_element()
                          .by_xpath("//*[@id='mensagensForm:mensagemDataTable:tb']//tr")
                          .get_all_elements())
                for msg in messages:
                    # ! Não usar "msg" pois a referência DOM não é renovada ao entrar e sair do modal de anexos.
                    driver.click(
                        driver.find_element().by_xpath(f'//a[contains(@id, "mensagensForm:mensagemDataTable:0:linkTitulo")]')
                    )

                    # * Decidi que seria útil coletar os dados da mensagem, caso no futuro
                    # * seja necessário utilizá-los.
                    def _collect_message_data() -> MensagemISSFortaleza:
                        # Coletar dados da mensagem
                        title = driver.find_element().by_id("mensagensForm:titulo").get_element()
                        data = driver.find_element().by_id("mensagensForm:dataRegistro").get_element()
                        conteudo = driver.find_element().by_id("mensagensForm:descricao").get_element()
                        return MensagemISSFortaleza(
                            title=title.text,
                            date=datetime.strptime(data.text, "%d/%m/%Y").date(),
                            content=conteudo.text
                        )

                    m = retry_on_exception(
                        func=_collect_message_data,
                        max_attempts=5,
                        polling_seconds=1,
                    )
                    empresa.mensagens.append(m)

                    # Verificar sem tem anexo.
                    # Precisa baixar todos os anexos antes de dar ciência.
                    try:
                        anexos = (driver.find_element()
                                  .by_xpath('//*[@id="mensagensForm:divAnexos"]//a[contains(@id, "linkVisualizarAnexo")]')
                                  .get_all_elements())
                        logger.info("Essa mensagem tem anexos — vou baixar todos antes de dar ciência.")
                        for idx, anexo in enumerate(anexos):
                            # ! Não usar "anexo", mesmo motivo de "msg".
                            # * Existem dois tipos de anexos: 1) que abrem o anexo em um modal para leitura, e 2) que só baixa o anexo.
                            try:
                                seletor_anexo = f"mensagensForm:anexos:{idx}:linkVisualizarAnexo"
                                # * Tentar clicar no 1° caso.
                                driver.click(
                                    driver.find_element().by_id(seletor_anexo),
                                    max_retries=5
                                )
                                # Baixar anexo
                                btn_baixar_anexo = driver.find_element().by_id("mensagensForm:botaoBaixarAnexo")
                                # driver.scroll_element_into_view(btn_baixar_anexo)
                                driver.click(btn_baixar_anexo)
                                driver.explicit_wait(.5)
                                # Voltar p/ form da mensagem
                                driver.click(
                                    driver.find_element().by_id("mensagensForm:botaoVoltarAnexo")
                                )
                            except (ElementClickInterceptedException, NoSuchElementException, StaleElementReferenceException):
                                # * Tentar clicar no 2° caso.
                                seletor_anexo = f"mensagensForm:anexos:{idx}:linkBaixarAnexo"
                                driver.click(
                                    driver.find_element().by_id(seletor_anexo),
                                    max_retries=5
                                )

                            driver.explicit_wait(1)
                            # * Em mensagems com muitos anexos, o botão de baixar anexo some da tela,
                            # * então precisamos esconder o link do anexo que foi baixado para poder clicar no botão de ciência.
                            driver.hide_elements((f"#{seletor_anexo}",))
                    except (TimeoutException, StaleElementReferenceException, NoSuchElementException):
                        pass # Não tem anexo

                    # Dar ciência
                    btn_dar_ciencia = driver.find_element().by_id("mensagensForm:botaoDarCiencia")
                    # driver.scroll_element_into_view(btn_dar_ciencia)
                    driver.click(btn_dar_ciencia)
            except NoSuchElementException:
                pass

            # * ------------------------------- Tela de Boas-Vindas Empresa --------------------------------

            # Pra ir até a tela de Escriturações, precisamos clicar nos botões do menu dropdown.
            # Nota: botões não tem nenhum identificador, então vamos procurar pelo texto interno.

            # Botão "Escriturações"
            driver.click(
                driver.find_element().by_xpath('//ul[contains(@class, "navbar-nav")]//a[normalize-space()="Escrituração"]')
            )

            # Opção "Manter Escrituração"
            driver.click(
                driver.find_element().by_xpath(
                    '//ul[contains(@class, "navbar-nav")]//li[@class="dropdown open"]//ul[@class="dropdown-menu"]//a[normalize-space()="Manter Escrituração"]'
                )
            )

            # * ----------------------------- Tela de Escrituração ------------------------------

            # Preencher campo de data com mês anterior ao atual
            logger.debug(
                "Preenchendo widget de calendário com a data do mês anterior..."
            )

            # Abrir widget de calendário
            driver.click(
                driver.find_element().by_class("rich-calendar-tool-btn")
            )

            # selecionar o mês no widget
            driver.click(
                driver.find_element().by_id(
                    f"manterEscrituracaoForm:dataInicialDateEditorLayoutM{competencia.month - 1}"  # Mês (0-based)
                )
            )

            # selecionar o ano no widget
            driver.click(
                driver.find_element().by_xpath(
                    f'//div[contains(@id, "manterEscrituracaoForm:dataInicialDateEditorLayoutY") and normalize-space()="{competencia.year}"]'
                )
            )

            # Clicar botão OK do widget
            driver.click(
                driver.find_element().by_id("manterEscrituracaoForm:dataInicialDateEditorButtonOk")
            )

            logger.debug("Consultando no portal se há inscrições em aberto para essa competência...")

            # Clicar botão "Consultar"
            driver.click(
                driver.find_element().by_id("manterEscrituracaoForm:btnConsultar")
            )

            driver.explicit_wait(2)

            # Procurar a linha da tabela de resultados que contém o registro relevante
            comp = competencia.strftime("%m/%Y")

            def _locate_element():
                return driver.find_element().by_xpath(
                    f'//tbody[@id="manterEscrituracaoForm:dataTable:tb"]/tr/td/span[normalize-space()="{comp}"]/../../self::tr'
                )

            search_result_row = retry_on_exception(
                func=_locate_element,
                max_attempts=10,
                polling_seconds=1
            )

            btn_escrituracao = (driver
                                .find_element(search_result_row)
                                .by_xpath(".//a[@title='Exportar Escrituração']"))

            btn_escrituracao.get_element().click()

            # * ----------------------------- Tela Exportar Escrituração ------------------------------

            # Selecionar formato XLS
            driver.click(
                driver.find_element().by_xpath(".//option[text()='XLS']")
            )

            # Clicar botão "Gerar"
            driver.click(
                driver.find_element().by_id("exportarEscrituracaoForm:btnGerar"),
            )

            download_btn = driver.find_element().by_id("exportarEscrituracaoForm:fileButton")

            check_thread_stopped_cb()

            try:
                driver.click(
                    download_btn,
                    max_retries=30,  # Tem que esperar ser gerado o XLS e aparecer o botão de download
                )
            except NoSuchElementException as no_element:
                # Botão não apareceu -- empresa provavelmente não tem movimento, então não foi gerado escrituração
                logger.error(
                    f"Falha ao baixar escrituração da empresa {empresa}: {no_element}"
                )

                driver.click(
                    driver.find_element().by_attribute(attr_name="title", attr_value="Alterar Inscrição Atual")
                )

                driver.explicit_wait(1)

                continue

            # Baixar arquivo e mover para diretório destino
            origin_file_name = remover_caracteres_escape(download_btn.get_element().text.strip())
            origin_file_path = temp_path / origin_file_name

            dest_file_prefix = f"EMP_{str(empresa.codigo)}_Escrituracao_ISS_{competencia.year}_{competencia.month:02d}"
            dest_file_name = Path(f"{dest_file_prefix}_{timestamp_as_file_name('xlsx')}")
            dest_file_dir = Path(f"{diretorio_saida}/{competencia.year}/{competencia.month:02d}")
            dest_file_path = dest_file_dir / dest_file_name

            try:
                # Aguardar arquivo ser baixado
                def _aguarda_arquivo_baixado():
                    verifica_arquivo_existe(  # Levanta FileNotFoundError se não achar arquivo
                        str(origin_file_name), str(temp_path)
                    )

                retry_on_exception(
                    _aguarda_arquivo_baixado,
                    max_attempts=30,
                    polling_seconds=1,
                    exception=OSError
                )

                # Criar diretório destino
                dest_file_dir.mkdir(parents=True, exist_ok=True)

                # Mover arquivo
                shutil.move(origin_file_path, dest_file_path)

                logger.success(f"Pronto! Escrituração da empresa salva em: {dest_file_path}")
                empresas_processadas.append(empresa)
            except OSError as oserr:
                logger.error(
                    f"Falha ao mover arquivo origem `{origin_file_path}` ao caminho destino `{dest_file_path}`: {oserr}"
                )

            driver.click(
                driver.find_element().by_attribute(attr_name="title", attr_value="Alterar Inscrição Atual")
            )

            driver.explicit_wait(1)

            if callback_progresso:
                callback_progresso(len(empresas_processadas) + qtd_empresas_ignoradas + qtd_empresas__sem_procuracao + qtd_empresas_erro, row_number)

            check_thread_stopped_cb()

        # Resultados
        logger.info(f"Terminei a captura. No total, processei {len(empresas_processadas)} empresa(s) com sucesso.")
        logger.info(f"{qtd_empresas_ignoradas} empresa(s) foram ignoradas (fora de Fortaleza ou sem movimento).")
        logger.info(f"{qtd_empresas__sem_procuracao} empresa(s) não tinham procuração no portal, então não consegui acessá-las.")
        logger.info(f"{qtd_empresas_erro} empresa(s) tiveram algum erro no meio do caminho.")
    except ThreadStoppedException:
        logger.warning("Recebi o pedido de cancelamento do operador e vou parar por aqui.")
    except Exception as exc:
        logger.critical(f"Algo deu muito errado e eu não sabia como continuar: {exc}")
    finally:
        driver.quit_driver()
