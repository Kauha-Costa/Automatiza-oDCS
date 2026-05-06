import os
import time
import threading
import win32com.client
import win32api
import win32print


def listar_impressoras() -> list[str]:
    """Retorna lista de impressoras disponíveis no sistema."""
    try:
        impressoras = []
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            impressoras.append(p[2])
        return impressoras
    except Exception:
        return []


def _confirmar_avisos_word(timeout_total: int = 10):
    """
    Fica rodando em background por timeout_total segundos tentando fechar
    quaisquer janelas de diálogo do Word (avisos de margens, etc.).
    Usa múltiplas estratégias para maior robustez.
    """
    import ctypes

    fim = time.time() + timeout_total
    confirmados = 0

    while time.time() < fim:
        try:
            # Estratégia 1: pywinauto - busca diálogos do Word pelo processo
            try:
                from pywinauto import Application, findwindows
                handles = findwindows.find_windows(class_name="#32770")  # class de diálogos Win32
                for h in handles:
                    try:
                        app = Application(backend="win32").connect(handle=h)
                        dlg = app.window(handle=h)
                        if dlg.exists(timeout=0.5):
                            # Tenta clicar em botão OK/Sim/Yes
                            for btn_title in ["OK", "Sim", "Yes", "&Sim", "&Yes", "&OK"]:
                                try:
                                    btn = dlg.child_window(title=btn_title, control_type="Button")
                                    if btn.exists(timeout=0.3):
                                        btn.click()
                                        confirmados += 1
                                        time.sleep(0.3)
                                        break
                                except Exception:
                                    pass
                    except Exception:
                        pass
            except ImportError:
                pass

            # Estratégia 2: SendKeys ENTER para janela em foreground (fallback)
            try:
                import win32gui
                import win32con
                hwnd = win32gui.GetForegroundWindow()
                titulo = win32gui.GetWindowText(hwnd)
                if "Word" in titulo or "Microsoft" in titulo:
                    win32api.keybd_event(0x0D, 0, 0, 0)   # ENTER down
                    win32api.keybd_event(0x0D, 0, 2, 0)   # ENTER up
                    confirmados += 1
                    time.sleep(0.3)
            except Exception:
                pass

        except Exception:
            pass

        time.sleep(0.4)

    return confirmados


def imprimir_certificados_em_pasta(
    pasta: str,
    impressora: str = "",
    log_callback=print,
    cancelar_flag=None
):
    """
    Imprime todos os .docx da pasta usando Word via COM.
    impressora: nome da impressora; "" = impressora padrão do sistema.
    cancelar_flag: lista [False]; vira [True] para cancelar.
    """
    if not os.path.exists(pasta):
        log_callback(f"❌ Pasta não encontrada: {pasta}")
        return 0

    arquivos = [arq for arq in os.listdir(pasta) if arq.lower().endswith(".docx")]
    if not arquivos:
        log_callback("⚠️  Nenhum arquivo .docx encontrado para impressão.")
        return 0

    total = len(arquivos)
    impressos = 0

    # Salva impressora padrão atual para restaurar depois
    impressora_original = win32print.GetDefaultPrinter() if impressora else ""

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone — suprime alertas automáticos

    try:
        # Muda impressora padrão temporariamente se solicitado
        if impressora and impressora != impressora_original:
            try:
                win32print.SetDefaultPrinter(impressora)
                log_callback(f"🖨️  Impressora selecionada: {impressora}")
            except Exception as e:
                log_callback(f"⚠️  Não foi possível selecionar impressora '{impressora}': {e}")

        for i, arquivo in enumerate(arquivos):
            if cancelar_flag and cancelar_flag[0]:
                log_callback("⛔ Impressão cancelada pelo usuário.")
                break

            caminho = os.path.join(pasta, arquivo)
            log_callback(f"⏳ [{i+1}/{total}] Imprimindo: {arquivo}")

            try:
                doc = word.Documents.Open(
                    os.path.abspath(caminho),
                    ReadOnly=True,
                    AddToRecentFiles=False
                )

                # Inicia confirmação de avisos em thread paralela
                t = threading.Thread(
                    target=_confirmar_avisos_word,
                    args=(12,),
                    daemon=True
                )
                t.start()

                # Envia para impressão
                doc.PrintOut(
                    Background=False  # Aguarda conclusão antes de continuar
                )

                time.sleep(1)
                doc.Close(SaveChanges=False)
                t.join(timeout=2)

                log_callback(f"✅ Impresso: {arquivo}")
                impressos += 1

            except Exception as e:
                log_callback(f"❌ Erro ao imprimir '{arquivo}': {e}")
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass

    finally:
        try:
            word.Quit()
        except Exception:
            pass

        # Restaura impressora original
        if impressora and impressora_original and impressora != impressora_original:
            try:
                win32print.SetDefaultPrinter(impressora_original)
            except Exception:
                pass

    return impressos
