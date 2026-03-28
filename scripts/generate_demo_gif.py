from __future__ import annotations

from pathlib import Path
import sys

from PIL import Image, ImageGrab

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from validador_glebas_app import App


ARQUIVO_EXEMPLO = ROOT / "TESTE_2_COM ERROS.xls"
SAIDA_GIF = ROOT / "assets" / "readme-demo.gif"


def capturar_janela(app: App) -> Image.Image:
    app.update_idletasks()
    x = app.winfo_rootx()
    y = app.winfo_rooty()
    w = app.winfo_width()
    h = app.winfo_height()
    frame = ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)
    frame.thumbnail((920, 680))
    return frame.convert("P", palette=Image.ADAPTIVE, colors=128)


def gerar_demo() -> None:
    app = App()
    app.geometry("1020x740+140+20")
    app.update()

    frames: list[Image.Image] = []
    durations: list[int] = []

    def registrar_frame(ms: int) -> None:
        frames.append(capturar_janela(app))
        durations.append(ms)

    def selecionar_arquivo() -> None:
        app._arquivo = str(ARQUIVO_EXEMPLO)
        app.zona_drop.set_arquivo(ARQUIVO_EXEMPLO.name)
        app.btn_validar.configure(state="normal")
        app.lbl_status.configure(text=f"📄  {ARQUIVO_EXEMPLO.name}")
        app.update()
        registrar_frame(1100)
        app.after(600, iniciar_validacao)

    def iniciar_validacao() -> None:
        app._iniciar_validacao()
        app.update()
        registrar_frame(900)
        app.after(1600, finalizar)

    def finalizar() -> None:
        app.update()
        registrar_frame(1800)
        SAIDA_GIF.parent.mkdir(parents=True, exist_ok=True)
        frames[0].save(
            SAIDA_GIF,
            save_all=True,
            append_images=frames[1:],
            duration=durations,
            loop=0,
            optimize=True,
            disposal=2,
        )
        app.destroy()

    registrar_frame(1000)
    app.after(700, selecionar_arquivo)
    app.mainloop()


if __name__ == "__main__":
    gerar_demo()
