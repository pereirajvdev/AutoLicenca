import argparse
import time
import re
from pathlib import Path

import pyautogui
import pyperclip
import win32gui
import ctypes

from ctypes import wintypes
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime


from pywinauto import Desktop
from functools import wraps


# ============================================================
# CONFIGURAÇÕES
# ============================================================

TITULO_MENU = "Menu: Arquivos / Cadastro de Funcionários"
TITULO_LICENCAS = "Licenças / Afastamentos / Exames"
TITULO_PESQUISANDO = "Pesquisando dados dos funcionários"

X_LUPA = 583
Y_LUPA = 841

X_LUPA_LICENCA = 629
Y_LUPA_LICENCA = 877

X_LICENCAS = 667
Y_LICENCAS = 495

X_PROXIMA_ACAO = 898
Y_PROXIMA_ACAO = 419

X_MATRICULA = 658
Y_MATRICULA = 628


INTERVALO = 0.3
INTERVALO_VERIFICACAO = 0.5
TEMPO_LIMITE = 60

# ======================================================================
# DESEMPENHO (--perf)
# ======================================================================
PERF_ATIVO = False
_TEMPOS = {}


def _registrar_tempo(nome, duracao):
    registro = _TEMPOS.setdefault(nome, [0, 0.0])
    registro[0] += 1
    registro[1] += duracao


class _SemMedicao:
    """Context manager nulo, compartilhado: custo praticamente zero com --perf desligado."""
    __slots__ = ()

    def __enter__(self):
        return None

    def __exit__(self, *_exc):
        return False


class _Medicao:
    __slots__ = ("nome", "inicio")

    def __init__(self, nome):
        self.nome = nome
        self.inicio = 0.0

    def __enter__(self):
        self.inicio = time.perf_counter()
        return self

    def __exit__(self, *_exc):
        _registrar_tempo(self.nome, time.perf_counter() - self.inicio)
        return False


_SEM_MEDICAO = _SemMedicao()


def medir(nome):
    """Mede o tempo de um bloco (`with medir(...)`). Sem efeito se --perf estiver desligado."""
    return _Medicao(nome) if PERF_ATIVO else _SEM_MEDICAO


def cronometrar(nome=None):
    """Decorator que mede o tempo total de uma função (só com --perf)."""
    def decorador(func):
        rotulo = nome or func.__name__

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not PERF_ATIVO:
                return func(*args, **kwargs)

            inicio = time.perf_counter()
            try:
                return func(*args, **kwargs)
            finally:
                _registrar_tempo(rotulo, time.perf_counter() - inicio)

        return wrapper
    return decorador


def imprimir_relatorio_tempos():
    if not PERF_ATIVO or not _TEMPOS:
        return

    print("\n" + "=" * 66)
    print(" RELATÓRIO DE DESEMPENHO (--perf)")
    print(" Obs.: tempos aninhados se sobrepõem; não some as linhas.")
    print("=" * 66)
    print(f"{'Etapa':<42}{'Chamadas':>9}{'Total (s)':>11}{'Média (s)':>11}")

    for nome, (chamadas, total) in sorted(
        _TEMPOS.items(), key=lambda item: item[1][1], reverse=True
    ):
        media = total / chamadas if chamadas else 0.0
        print(f"{nome:<42}{chamadas:>9}{total:>11.2f}{media:>11.3f}")

    print("=" * 66)


def _banner(texto):
    """Imprime o cabeçalho padrão de seção (linha em branco + moldura)."""
    print()
    print("=" * 80)
    print(texto)
    print("=" * 80)

# ======================================================================
# WIN32 (ctypes): enumeração rápida de janelas
# ======================================================================
_user32 = ctypes.windll.user32

_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

_user32.EnumWindows.argtypes = [_WNDENUMPROC, wintypes.LPARAM]
_user32.EnumWindows.restype = wintypes.BOOL
_user32.IsWindowVisible.argtypes = [wintypes.HWND]
_user32.IsWindowVisible.restype = wintypes.BOOL
_user32.IsHungAppWindow.argtypes = [wintypes.HWND]
_user32.IsHungAppWindow.restype = wintypes.BOOL
_user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
_user32.GetWindowTextLengthW.restype = ctypes.c_int
_user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
_user32.GetWindowTextW.restype = ctypes.c_int

# Instância única de Desktop, usada só para criar o wrapper da janela escolhida.
_DESKTOP = Desktop(backend="win32")

_AVISOS_TRAVADA = set()


def _listar_janelas_visiveis():
    """Enumera janelas de nível superior visíveis: [(hwnd, titulo), ...]."""
    resultado = []

    def callback(hwnd, _lparam):
        if _user32.IsWindowVisible(hwnd):
            tamanho = _user32.GetWindowTextLengthW(hwnd)
            if tamanho:
                buffer = ctypes.create_unicode_buffer(tamanho + 1)
                _user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
                resultado.append((hwnd, buffer.value))
        return True

    proc = _WNDENUMPROC(callback)
    _user32.EnumWindows(proc, 0)
    return resultado


def _buscar_janela(titulos):
    """
    Faz UMA enumeração e retorna (hwnd, titulo_real) da primeira janela
    válida, respeitando a ordem dos títulos informados.
    Comparação: texto idêntico (sem diferenciar maiúsculas/minúsculas).
    Janelas travadas são ignoradas.
    """
    candidatas = {}

    for hwnd, texto in _listar_janelas_visiveis():
        texto = texto.strip()
        candidatas.setdefault(texto.casefold(), []).append((hwnd, texto))

    for titulo in titulos:
        for hwnd, texto in candidatas.get(titulo.strip().casefold(), []):
            if _user32.IsHungAppWindow(hwnd):
                if hwnd not in _AVISOS_TRAVADA:
                    _AVISOS_TRAVADA.add(hwnd)
                    print(f"[AGUARDANDO] Janela sem resposta: {texto}")
                continue

            return hwnd, texto

    return None


def _criar_wrapper(hwnd):
    with medir("criar wrapper pywinauto"):
        return _DESKTOP.window(handle=hwnd).wrapper_object()


@cronometrar()
def localizar_janela(titulo_esperado):
    """Verificação única (sem espera). Retorna wrapper ou None."""
    achado = _buscar_janela([titulo_esperado])

    if achado is None:
        return None

    hwnd, titulo = achado
    print(f"[OK] Janela identificada: {titulo}")
    return _criar_wrapper(hwnd)


@cronometrar()
def aguardar_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    """
    Aguarda até que uma janela com um dos títulos apareça.
    Retorna o wrapper da primeira janela válida (seguindo a ordem dos
    títulos) ou None se o tempo limite estourar.
    """
    print(
        f"[INFO] Aguardando uma das janelas: "
        f"{', '.join(titulos_esperados)}"
    )

    limite = time.monotonic() + time_limit

    while True:
        with medir("enumerar janelas"):
            achado = _buscar_janela(titulos_esperados)

        if achado is not None:
            hwnd, titulo = achado
            print(f"[OK] Janela identificada: {titulo}")
            return _criar_wrapper(hwnd)

        if time.monotonic() >= limite:
            break

        time.sleep(INTERVALO_VERIFICACAO)

    print(
        f"[INFO] Nenhuma janela válida apareceu em "
        f"{time_limit} segundos: "
        f"{', '.join(titulos_esperados)}"
    )

    return None


def exigir_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    janela = aguardar_janela(*titulos_esperados, time_limit=time_limit)

    if janela is None:
        raise RuntimeError(
            "Janela não encontrada: " + ", ".join(titulos_esperados)
        )

    return janela


def procurar_controle(janela, titulo=None, classe=None, controles=None):
    """
    Localiza o primeiro descendente pelo título e/ou classe.
    """
    if controles is None:
        controles = janela.descendants()

    for i, controle in enumerate(controles):
        try:
            texto_ok = (
                titulo is None
                or controle.window_text() == titulo
            )

            classe_ok = (
                classe is None
                or controle.class_name() == classe
            )

            if texto_ok and classe_ok:
                return controle

        except Exception as e:
            print(f"[{i}] Erro: {e}")

    print(
        f"[ERRO] Controle não encontrado: "
        f"titulo={titulo!r}, classe={classe!r}"
    )

    return None


# ============================================================
# LEITURA DA PLANILHA
# ============================================================

def normalizar_matricula(valor):
    """
    Normaliza a matrícula para seis dígitos, completando com zeros à esquerda.

    Exemplos:
        2062    -> 002062
        31205   -> 031205
        002062  -> 002062
        2062.0  -> 002062
    """
    if valor is None:
        return None

    texto = str(valor).strip()

    if not texto:
        return None

    # Remove o .0 de valores numéricos lidos como float.
    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]

    # Mantém somente os números.
    numeros = re.sub(r"\D", "", texto)

    if not numeros:
        return None

    return numeros.zfill(6)


def ler_servidores(caminho_planilha, usar_matriculas=False):
    """
    Lê os nomes dos servidores na coluna B e, opcionalmente,
    as matrículas na coluna A.

    Retorna uma lista de tuplas:
        (nome, matricula)
    """
    workbook = load_workbook(
        caminho_planilha,
        read_only=True,
        data_only=True,
    )

    try:
        planilha = workbook.active
        servidores = []

        for numero_linha, linha in enumerate(
            planilha.iter_rows(
                min_col=1,
                max_col=2,
                values_only=True,
            ),
            start=1,
        ):
            valor_matricula, valor_nome = linha

            if valor_nome is None:
                continue

            nome = str(valor_nome).strip()

            if not nome:
                continue

            matricula = None

            if usar_matriculas:
                matricula = normalizar_matricula(valor_matricula)

                if matricula is None:
                    raise ValueError(
                        f"Matrícula inválida na linha {numero_linha}: "
                        f"{valor_matricula!r}"
                    )

            servidores.append((nome, matricula))

        return servidores

    finally:
        workbook.close()


# ============================================================
# PESQUISA DO SERVIDOR
# ============================================================

def primeira_execucao(
    nome,
    numero_linha=None,
    caminho_planilha=None,
    matricula=None,
):
    print(f"\nPesquisando servidor: {nome}")

    print(f"\n{'=' * 50}")
    print(f"Linha {numero_linha}: {nome}")

    if matricula is not None:
        print(f"Matrícula: {matricula}")

    print("Aguardando o Menu...")

    janela = aguardar_janela(TITULO_MENU)
    janela.set_focus()

    print("Clicando na Lupa")
    pyautogui.click(X_LUPA, Y_LUPA)
    pyautogui.click(X_LUPA, Y_LUPA)

    acao_pesquisar(
        nome,
        numero_linha,
        caminho_planilha,
        matricula,
    )

    janela = aguardar_janela(TITULO_MENU)
    janela.set_focus()

    print("Abrindo janela de Licenças...")
    pyautogui.click(X_LICENCAS, Y_LICENCAS)
    pyautogui.click(X_LICENCAS, Y_LICENCAS)

    operar_licencas(nome, numero_linha, caminho_planilha)


def acao_pesquisar(
    nome,
    numero_linha=None,
    caminho_planilha=None,
    matricula=None,
):
    janela = aguardar_janela(TITULO_PESQUISANDO)
    janela.set_focus()

    if matricula is not None:
        print(f"Pesquisando pela matrícula: {matricula}")

        radio_matricula = procurar_controle(
            janela,
            titulo="&Matrícula",
            classe="TRadioButton"
        )
        radio_matricula.set_focus()
        radio_matricula.click()

        campo = procurar_controle(
            janela,
            classe="TEdit"
        )
        campo.set_focus()

        pyperclip.copy(matricula)

    else:
        print(f"Pesquisando pelo nome: {nome}")

        pyperclip.copy(nome)

    pyautogui.hotkey("ctrl", "v")
    time.sleep(.05)

    print("Pressionando primeiro Enter...")
    pyautogui.press("enter")

    print("Pressionando segundo Enter...")
    pyautogui.press("enter")

    return True


def operar_licencas(nome, numero_linha=None, caminho_planilha=None):
    # Opera a janela de licencas

    janela = aguardar_janela(TITULO_LICENCAS)
    janela.set_focus()

    print("Executando próximo clique...")
    pyautogui.click(X_PROXIMA_ACAO, Y_PROXIMA_ACAO)
    pyautogui.click(X_PROXIMA_ACAO, Y_PROXIMA_ACAO)
    time.sleep(.05)

    print("Clique para abrir Grade Licenças realizado.")

    obter_licencas(nome, caminho_planilha) 

    return True


def pesquisar_servidor(
    nome,
    numero_linha=None,
    caminho_planilha=None,
    matricula=None,
):
    pyautogui.click(X_LUPA_LICENCA, Y_LUPA_LICENCA)
    pyautogui.click(X_LUPA_LICENCA, Y_LUPA_LICENCA)

    if not acao_pesquisar(
        nome,
        numero_linha,
        caminho_planilha,
        matricula,
    ):
        return False

    if not operar_licencas(nome, numero_linha, caminho_planilha):
        return False

    print(f"Sequência concluída para: {nome}")

    return True


def obter_licencas(nome_servidor, caminho_saida):
    """
    Lê os registros da grade e acrescenta as licenças
    ao Excel consolidado.
    """

    print(f"\nObtendo licenças de: {nome_servidor}")

    pyautogui.click(800, 460)
    time.sleep(0.05)

    registros = []
    anterior = None
    
    limite_tentativas = 100

    for tentativa in range(1, limite_tentativas + 1):
        pyperclip.copy("")

        pyautogui.hotkey("ctrl", "c")

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
        description="Executa a pesquisa dos servidores de uma planilha."
    )

    parser.add_argument(
        "planilha",
        type=Path,
        help="Caminho da planilha Excel.",
    )

    parser.add_argument(
        "--matriculas",
        action="store_true",
        help="Utiliza as matrículas da coluna A para pesquisar os servidores.",
    )

    args = parser.parse_args()

    if not args.planilha.exists():
        print(f"Planilha não encontrada: {args.planilha}")
        return

    try:
        servidores = ler_servidores(
            args.planilha,
            usar_matriculas=args.matriculas,
        )

    except Exception as erro:
        print(f"Erro ao ler a planilha: {erro}")
        return

    if not servidores:
        print("Nenhum servidor encontrado na planilha.")
        return

    print(f"Total de servidores encontrados: {len(servidores)}")

    if args.matriculas:
        print("Modo de pesquisa: MATRÍCULA")
    else:
        print("Modo de pesquisa: NOME")

    caminho_saida = (
        args.planilha.resolve().parent / "Licencas.xlsx"
    )

    criar_planilha_licencas(args.planilha)

    for indice, (nome, matricula) in enumerate(servidores, start=1):

        if indice == 1:
            primeira_execucao(
                nome,
                indice,
                caminho_saida,
                matricula,
            )

        else:
            pesquisar_servidor(
                nome,
                indice,
                caminho_saida,
                matricula,
            )


if __name__ == "__main__":
    main()