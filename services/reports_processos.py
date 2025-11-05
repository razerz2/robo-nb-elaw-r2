from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time


def _abrir_dialog_excel(driver, wait):
    """Abre o diálogo do Excel e muda para o iframe."""
    btn_excel = wait.until(EC.element_to_be_clickable((By.ID, "btnExcel")))
    btn_excel.click()
    print("📥 Botão Excel clicado.")

    wait.until(EC.visibility_of_element_located((By.ID, "btnExcel_dlg")))
    iframe = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#btnExcel_dlg iframe")))
    driver.switch_to.frame(iframe)
    print("🔄 Mudamos para o iframe do relatório.")


def _configurar_modelo(driver, wait):
    """Seleciona modelo pré-configurado e relatório 'Tarefas'."""
    # Modelo pré-configurado
    lbl = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//label[@for='elawReportForm:elawReportOption:0']"))
    )
    lbl.click()
    print("☑️ Selecionado: Modelos pré-configurados")

    btn_continuar = wait.until(EC.element_to_be_clickable((By.ID, "elawReportForm:continuarBtn")))
    btn_continuar.click()
    print("➡️ Continuar clicado.")

    # Dropdown Tarefas
    dd = wait.until(EC.element_to_be_clickable((By.ID, "elawReportForm:selectElawReport_label")))
    dd.click()
    opcoes = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR,
            "ul[id$='selectElawReport_items'] li.ui-selectonemenu-item"))
    )
    time.sleep(1)

    alvo = next((o for o in opcoes if o.text.strip() == "Recebimentos"), None)
    if not alvo:
        raise Exception("❌ Opção 'Recebimentos' não encontrada!")

    wait.until(EC.element_to_be_clickable(alvo)).click()
    print("✔️ Relatório selecionado: Recebimentos")

    # Botão gerar
    btn_gerar = wait.until(EC.element_to_be_clickable((By.ID, "elawReportForm:elawReportGerarBtn")))
    btn_gerar.click()
    print("📊 Gerar relatório clicado.")


def _capturar_id(driver, wait):
    """Captura o ID do relatório gerado."""
    id_elem = wait.until(EC.presence_of_element_located((
        By.XPATH, "//span[normalize-space()='ID']/ancestor::div[contains(@class,'ui-g')]/div[last()]"
    )))
    relatorio_id = id_elem.text.strip()
    print(f"🆔 Relatório solicitado com ID: {relatorio_id}")
    return relatorio_id


def gerar_relatorio(driver):
    """
    Fluxo completo:
    - Acessa a tela
    - Abre Excel -> Recebimentos
    - Gera relatório
    - Retorna o ID
    """
    wait = WebDriverWait(driver, 30)
    url_relatorio = "https://sicredi.elaw.com.br/processoList.elaw"
    driver.get(url_relatorio)
    time.sleep(2)

    # Pesquisar
    btn = wait.until(EC.element_to_be_clickable((By.ID, "btnPesquisar")))
    btn.click()
    print("🔎 Pesquisa disparada.")

    # Excel
    _abrir_dialog_excel(driver, wait)
    time.sleep(2)
    _configurar_modelo(driver, wait)
    time.sleep(2)
    relatorio_id = _capturar_id(driver, wait)

    # Volta pro principal e recarrega
    driver.switch_to.default_content()
    driver.refresh()
    print("🔄 Página recarregada.")

    return relatorio_id
