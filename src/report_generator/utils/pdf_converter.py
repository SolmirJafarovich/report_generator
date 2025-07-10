import platform
import subprocess
import sys
from pathlib import Path


def convert_pptx_to_pdf(input_path: str, output_path: str = None) -> bool:
    """
    Конвертирует PPTX-файл в PDF с использованием PowerPoint (Windows) или LibreOffice (Linux/macOS).

    :param input_path: путь до входного .pptx
    :param output_path: путь до выходного .pdf (если None — используется имя из input)
    :return: True, если успешно, иначе False
    """
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"PPTX not found: {input_path}")

    output_path = Path(output_path or input_path.with_suffix(".pdf")).resolve()
    system = platform.system()

    if system == "Windows":
        try:
            import comtypes.client  # Убедитесь, что установлен через pyproject.toml

            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(
                str(input_path), WithWindow=False
            )
            presentation.SaveAs(str(output_path), FileFormat=32)  # 32 = PDF
            presentation.Close()
            powerpoint.Quit()

            return output_path.exists()

        except ImportError:
            print(
                "[!] Не установлены pywin32 / comtypes. Добавьте их в pyproject.toml с marker'ом `sys_platform == 'win32'`."
            )
            return False
        except Exception as e:
            print(f"[!] PowerPoint conversion failed: {e}")
            return False

    elif system in {"Linux", "Darwin"}:
        try:
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_path.parent),
                str(input_path),
            ]
            subprocess.run(
                cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )

            return output_path.exists()

        except Exception as e:
            print(f"[!] LibreOffice conversion failed: {e}")
            return False

    else:
        print(f"[!] Unsupported platform: {system}")
        return False


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pptx_to_pdf.py <input.pptx> [output.pdf]")
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) >= 3 else None

        success = convert_pptx_to_pdf(input_file, output_file)
        print("Успешно!" if success else "Ошибка.")
