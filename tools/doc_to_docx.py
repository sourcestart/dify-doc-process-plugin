import os
import re
import shutil
import subprocess
import tempfile
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


class DocToDocxTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        input_file = tool_parameters.get("input_file")
        if not input_file:
            yield self.create_text_message("Error: Missing required parameter 'input_file'.")
            return

        extension = (input_file.extension or "").lower().lstrip(".")
        filename = input_file.filename or "document.doc"
        if extension != "doc" and not filename.lower().endswith(".doc"):
            yield self.create_text_message("Error: Invalid file format. Only .doc files are supported.")
            return

        converter = self._find_converter()
        if not converter:
            yield self.create_text_message(
                "Error: LibreOffice is not installed or not found in PATH. Please install LibreOffice and ensure "
                "'soffice' (or 'libreoffice') is executable."
            )
            return

        safe_filename = self._sanitize_filename(filename)
        if not safe_filename.lower().endswith(".doc"):
            safe_filename = f"{safe_filename}.doc"

        with tempfile.TemporaryDirectory(prefix="doc_to_docx_") as temp_dir:
            input_path = os.path.join(temp_dir, safe_filename)
            with open(input_path, "wb") as f:
                f.write(input_file.blob or b"")

            if os.path.getsize(input_path) == 0:
                yield self.create_text_message("Error: Uploaded file is empty.")
                return

            output_filename = f"{os.path.splitext(safe_filename)[0]}.docx"
            output_path = os.path.join(temp_dir, output_filename)

            conversion_error = self._convert(converter, input_path, temp_dir)
            if conversion_error:
                yield self.create_text_message(f"Error: Conversion failed. {conversion_error}")
                return

            if not os.path.exists(output_path):
                yield self.create_text_message("Error: Conversion failed. Output .docx file was not generated.")
                return

            with open(output_path, "rb") as f:
                output_blob = f.read()

            if not output_blob:
                yield self.create_text_message("Error: Conversion failed. Output .docx file is empty.")
                return

            yield self.create_json_message(
                {
                    "status": "success",
                    "source_file": safe_filename,
                    "output_file": output_filename,
                    "converter": os.path.basename(converter),
                }
            )
            yield self.create_blob_message(
                blob=output_blob,
                meta={
                    "filename": output_filename,
                    "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                },
            )

    def _find_converter(self) -> str | None:
        candidates = ["soffice", "libreoffice", "soffice.exe", "libreoffice.exe"]
        for name in candidates:
            path = shutil.which(name)
            if path:
                return path
        return None

    def _convert(self, converter: str, input_path: str, output_dir: str) -> str | None:
        cmd = [
            converter,
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            output_dir,
            input_path,
        ]
        try:
            process = subprocess.run(
                cmd,
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=120,
            )
            if process.returncode != 0:
                stderr = (process.stderr or "").strip()
                stdout = (process.stdout or "").strip()
                details = stderr or stdout or f"command exited with code {process.returncode}"
                return details
            return None
        except subprocess.TimeoutExpired:
            return "conversion timed out after 120 seconds."
        except FileNotFoundError:
            return "converter executable not found."
        except Exception as e:
            return str(e)

    def _sanitize_filename(self, filename: str) -> str:
        sanitized = re.sub(r'[\\/*?:"<>|]', "", filename).strip()
        return sanitized or "document.doc"
