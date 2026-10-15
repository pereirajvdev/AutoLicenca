import argparse
import time
from pathlib import Path

import pyautogui
import pyperclip
import win32gui

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime


# ============================================================
# CONFIGURAÇÕES
# ============================================================

TITULO_MENU = "Menu: Arquivos / Cadastro de Funcionários"

X_INICIAL = 583
Y_INICIAL = 841

X_LICENCAS = 667
Y_LICENCAS = 495

X_PROXIMA_ACAO = 898
Y_PROXIMA_ACAO = 419

INTERVALO = 0.3
INTERVALO_VERIFICACAO = 0.5


# ============================================================
# IDENTIFICAÇÃO DE JANELAS
# ============================================================

def encontrar_janela_por_titulo(titulo_procurado):
    """
    Retorna o identificador da primeira janela visível
    cujo título contenha o texto procurado.
    """
    hwnd_encontrado = None
    titulo_procurado = titulo_procurado.casefold()

    def callback(hwnd, _):
        nonlocal hwnd_encontrado

        try:
            titulo = win32gui.GetWindowText(hwnd).strip()

            if (
                win32gui.IsWindowVisible(hwnd)
                and titulo_procurado in titulo.casefold()
            ):
                hwnd_encontrado = hwnd
                return False

        except Exception:
            pass

        return True

    win32gui.EnumWindows(callback, None)

    return hwnd_encontrado


def aguardar_estado_menu(timeout=30):
    """
    Aguarda até que o Menu esteja aberto.
    """
    limite = time.time() + timeout

    while time.time() < limite:
        if encontrar_janela_por_titulo(TITULO_MENU) is not None:
            return True

        time.sleep(.1)

    return False


def ativar_janela(titulo, timeout=15):
    """
    Localiza e ativa a janela pelo título.
    """
    print(f"Esperando janela: {titulo}")

    limite = time.time() + timeout

    while time.time() < limite:
        hwnd = encontrar_janela_por_titulo(titulo)

        if hwnd:
            try:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(INTERVALO)

                print(f"Janela ativa: {win32gui.GetWindowText(hwnd)}")
                return True

            except Exception as erro:
                print(f"Não foi possível ativar a janela: {erro}")
                return False

        time.sleep(0.2)

    print(f"Janela não encontrada: {titulo}")
    return False


# ============================================================
# LEITURA DA PLANILHA
# ============================================================

def ler_servidores(caminho_planilha):
    """
    Lê exclusivamente os nomes dos servidores na coluna B.
    Ignora completamente a coluna A (matrículas).
    """
    workbook = load_workbook(
        caminho_planilha,
        read_only=True,
        data_only=True,
    )

    try:
        planilha = workbook.active
        servidores = []

        for linha in planilha.iter_rows(
            min_col=2,
            max_col=2,
            values_only=True,
        ):
            valor = linha[0]

            if valor is None:
                continue

            nome = str(valor).strip()

            if nome:
                servidores.append(nome)

        return servidores

    finally:
        workbook.close()


# ============================================================
# PESQUISA DO SERVIDOR
# ============================================================

def pesquisar_servidor(nome, numero_linha=None, caminho_planilha=None):
    print(f"\nPesquisando servidor: {nome}")
    """
    Executa a sequência inicial:
    1. Aguarda o Menu.
    2. Ativa a janela.
    3. Clica na posição inicial.
    4. Cola o nome.
    5. Pressiona Enter duas vezes.
    """

    print(f"\n{'=' * 50}")
    print(f"Linha {numero_linha}: {nome}")

    print("Aguardando o Menu...")

    if not aguardar_estado_menu():
        print("Menu não encontrado. Servidor ignorado.")
        return False

    print(f"Janela encontrada: {TITULO_MENU}")

    if not ativar_janela(TITULO_MENU):
        return False

    print("Clicando na posição inicial...")
    pyautogui.click(X_INICIAL, Y_INICIAL)
    pyautogui.click(X_INICIAL, Y_INICIAL)
    time.sleep(INTERVALO)

    print(f"Pesquisando servidor: {nome}")
    pyperclip.copy(nome)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(.05)

    print("Pressionando primeiro Enter...")
    pyautogui.press("enter")
    time.sleep(.05)

    print("Pressionando segundo Enter...")
    pyautogui.press("enter")
    time.sleep(.5)

    print("Abrindo janela de Licenças...")
    pyautogui.click(X_LICENCAS, Y_LICENCAS)
    pyautogui.click(X_LICENCAS, Y_LICENCAS)
    time.sleep(.5)

    print("Executando próximo clique...")
    pyautogui.click(X_PROXIMA_ACAO, Y_PROXIMA_ACAO)
    pyautogui.click(X_PROXIMA_ACAO, Y_PROXIMA_ACAO)
    time.sleep(.05)

    print("Clique para abrir Licenças realizado.")

    obter_licencas(nome, caminho_planilha)
    
    time.sleep(.5)

    print(f"Sequência concluída para: {nome}")
    print(f"Voltando ao menu")
    pyautogui.hotkey("esc")
    
    time.sleep(.5)

    return True


def obter_licencas(nome_servidor, caminho_saida):
    """
    Lê os registros da grade e acrescenta as licenças
    ao Excel consolidado.
    """
    titulo = "Licenças / Afastamentos / Exames"
    limite_tentativas = 100

    hwnd = encontrar_janela_por_titulo(titulo)

    if not hwnd:
        print("Janela de Licenças não encontrada. (1/2)")
        time.sleep(.5)
        hwnd = encontrar_janela_por_titulo(titulo)
        if not hwnd:
            print("Janela de Licenças não encontrada. (2/2)")
            return

    print(f"\nObtendo licenças de: {nome_servidor}")

    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.5)

    pyautogui.click(800, 460)
    time.sleep(0.3)

    registros = []
    anterior = None

    for tentativa in range(1, limite_tentativas + 1):
        pyperclip.copy("")
        time.sleep(0.1)

        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.05)

        atual = pyperclip.paste()

        if not atual.strip():
            print("Nenhum texto foi copiado.")
            break

        if anterior is not None and atual == anterior:
            print("Fim da leitura: conteúdo repetido.")
            break

        registros.append(atual.strip())
        anterior = atual

        print(f"Registro {len(registros)} capturado.")

        pyautogui.press("down")
        time.sleep(0.01)

    if not registros:
        print("Nenhuma licença foi encontrada.")
        return

    workbook = load_workbook(caminho_saida)

    try:
        planilha = workbook["Licenças"]

        contador = sum(
            1
            for linha in planilha.iter_rows(
                min_row=2,
                min_col=1,
                max_col=1,
                values_only=True,
            )
            if linha[0] == nome_servidor
        )

        for registro in registros:
            dados = interpretar_licenca(registro)

            if dados is None:
                continue

            contador += 1

            planilha.append([
                nome_servidor,
                dados["tipo"],
                datetime.strptime(
                    dados["data_inicio"], "%d/%m/%Y"
                ).date(),
                datetime.strptime(
                    dados["data_fim"], "%d/%m/%Y"
                ).date(),
            ])

            linha = planilha.max_row

            for coluna in range(1, 5):
                planilha.cell(linha, coluna).alignment = Alignment(
                    vertical="top",
                    wrap_text=(coluna == 2),
                )

            for coluna in (3, 4):
                planilha.cell(linha, coluna).number_format = "dd/mm/yyyy"
                planilha.cell(linha, coluna).alignment = Alignment(
                    horizontal="center",
                    vertical="top",
                )

        planilha.auto_filter.ref = planilha.dimensions

        workbook.save(caminho_saida)

    finally:
        workbook.close()

    print("\nLicenças adicionadas ao Excel.")
    print(f"Servidor: {nome_servidor}")
    print(f"Registros adicionados: {len(registros)}")
    print(f"Arquivo: {Path(caminho_saida).resolve()}")


def criar_planilha_licencas(caminho_planilha):
    caminho_saida = (
        Path(caminho_planilha).resolve().parent / "Licencas.xlsx"
    )

    workbook = Workbook()
    planilha = workbook.active
    planilha.title = "Licenças"

    planilha.append([
        "Servidor",
        "Tipo Licença",
        "Data Início",
        "Data Fim",
    ])

    for celula in planilha[1]:
        celula.fill = PatternFill(
            fill_type="solid",
            fgColor="4472C4",
        )
        celula.font = Font(
            bold=True,
            color="FFFFFF",
        )
        celula.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    planilha.freeze_panes = "A2"
    planilha.auto_filter.ref = "A1:D1"

    planilha.column_dimensions["A"].width = 35
    planilha.column_dimensions["B"].width = 65
    planilha.column_dimensions["C"].width = 16
    planilha.column_dimensions["D"].width = 16

    workbook.save(caminho_saida)
    workbook.close()

    print(f"Planilha de saída criada: {caminho_saida}")

    return caminho_saida

import re


def interpretar_licenca(texto):
    """
    Interpreta o texto copiado da grade de licenças.

    Retorna os campos separados ou None se não conseguir
    identificar um registro válido.
    """

    texto = texto.replace("\r", "\n")

    # Separa o conteúdo por tabulações e remove espaços extras.
    campos = [
        re.sub(r"\s+", " ", campo).strip()
        for campo in re.split(r"[\t\n]+", texto)
        if campo.strip()
    ]

    # Remove os cabeçalhos da grade.
    cabecalhos = {
        "tipo licença",
        "data início",
        "data fim",
    }

    campos = [
        re.sub(r"\s+", " ", campo).strip()
        for campo in re.split(r"[\t\n]+", texto)
        if campo.strip()
    ]

    # Procura a sequência:
    # Controle | Tipo | Data inicial | Data final |
    # Dias afastado | Data de retorno
    for i in range(len(campos) - 5):
        controle = campos[i]
        tipo = campos[i + 1]
        data_inicio = campos[i + 2]
        data_fim = campos[i + 3]
        dias = campos[i + 4]
        data_retorno = campos[i + 5]

        if not controle.isdigit():
            continue

        formato_data = re.compile(r"\d{2}/\d{2}/\d{4}")

        if not all(
            formato_data.fullmatch(data)
            for data in (data_inicio, data_fim, data_retorno)
        ):
            continue

        if not dias.isdigit():
            continue

        return {
            "controle": int(controle),
            "tipo": tipo,
            "data_inicio": data_inicio,
            "data_fim": data_fim,
            "dias": int(dias),
            "data_retorno": data_retorno,
        }

    print(f"Não foi possível interpretar o registro: {texto!r}")
    return None

# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Executa a pesquisa inicial dos servidores de uma planilha."
    )

    parser.add_argument(
        "planilha",
        type=Path,
        help="Caminho da planilha Excel.",
    )

    args = parser.parse_args()

    if not args.planilha.exists():
        print(f"Planilha não encontrada: {args.planilha}")
        return

    try:
        servidores = ler_servidores(args.planilha)
    except Exception as erro:
        print(f"Erro ao ler a planilha: {erro}")
        return

    if not servidores:
        print("Nenhum nome encontrado na coluna B.")
        return

    print(f"Total de servidores encontrados: {len(servidores)}")

    caminho_saida = (
        args.planilha.resolve().parent / "Licencas.xlsx"
    )

    # Cria um único arquivo para todo o processamento.
    criar_planilha_licencas(args.planilha)

    for indice, nome in enumerate(servidores, start=1):
        try:
            pesquisar_servidor(
                nome,
                indice,
                caminho_saida,
            )
            time.sleep(0.5)

        except KeyboardInterrupt:
            print("\nProcessamento interrompido pelo usuário.")
            break

        except Exception as erro:
            print(f"Erro ao processar {nome}: {erro}")


if __name__ == "__main__":
    main()