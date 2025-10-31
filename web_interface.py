"""Web interface for process_103 Excel conversion."""
from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Optional

from flask import Flask, render_template_string, request, send_file
from werkzeug.utils import secure_filename

from process_103 import process_excel

UPLOAD_FORM_TEMPLATE = """
<!doctype html>
<html lang=\"ru\">
  <head>
    <meta charset=\"utf-8\">
    <title>Конвертация Excel в формат 103</title>
    <style>
      body { font-family: sans-serif; margin: 2rem auto; max-width: 720px; }
      h1 { margin-bottom: 1.5rem; }
      form { display: grid; grid-template-columns: 1fr 2fr; gap: 0.75rem 1rem; }
      label { font-weight: 600; }
      input[type="file"], input[type="text"] { padding: 0.4rem; }
      .full-width { grid-column: 1 / -1; }
      .messages { color: #b00020; margin-bottom: 1rem; }
      button { padding: 0.6rem 1.2rem; font-size: 1rem; }
    </style>
  </head>
  <body>
    <h1>Конвертация Excel в формат 103</h1>
    {% if messages %}
      <div class="messages">
        {% for message in messages %}
          <div>{{ message }}</div>
        {% endfor %}
      </div>
    {% endif %}
    <form method="post" enctype="multipart/form-data">
      <label for="input_file">Исходный Excel *</label>
      <input id="input_file" name="input_file" type="file" accept=".xlsx,.xls" required>

      <label for="template_file">Шаблон (опционально)</label>
      <input id="template_file" name="template_file" type="file" accept=".xlsx,.xls">

      <label for="inn">ИНН отправителя</label>
      <input id="inn" name="inn" type="text" placeholder="например, 7409002190">

      <label for="region">Регион (oblo)</label>
      <input id="region" name="region" type="text" value="Челябинская область">

      <div class="full-width">
        <button type="submit">Конвертировать и скачать</button>
      </div>
    </form>
    <p class="full-width">* Обязательно укажите либо ИНН, либо загрузите шаблон с колонкой <code>inni</code>.</p>
  </body>
</html>
"""


def create_app(secret_key: Optional[str] = None) -> Flask:
    """Create and configure the Flask application."""

    app = Flask(__name__)
    app.secret_key = secret_key or os.environ.get("PROCESS_103_WEB_SECRET", "dev-secret-key")
    app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB uploads

    @app.route("/", methods=["GET", "POST"])
    def index():
        if request.method == "GET":
            return render_template_string(UPLOAD_FORM_TEMPLATE, messages=None)

        messages = []
        input_file = request.files.get("input_file")
        template_file = request.files.get("template_file")
        inn_value = request.form.get("inn") or None
        region = request.form.get("region") or "Челябинская область"

        if not input_file or not input_file.filename:
            messages.append("Загрузите исходный Excel-файл.")

        if messages:
            return render_template_string(UPLOAD_FORM_TEMPLATE, messages=messages)

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            input_path = tmp_path / (secure_filename(input_file.filename) or "input.xlsx")
            input_file.save(input_path)

            template_path: Optional[Path] = None
            if template_file and template_file.filename:
                template_path = tmp_path / (secure_filename(template_file.filename) or "template.xlsx")
                template_file.save(template_path)

            output_path = tmp_path / "103_result.xlsx"
            try:
                process_excel(
                    str(input_path),
                    str(output_path),
                    inn=inn_value,
                    template_path=str(template_path) if template_path else None,
                    region=region,
                )
            except Exception as exc:
                messages.append(str(exc))
                return render_template_string(UPLOAD_FORM_TEMPLATE, messages=messages)

            return send_file(
                output_path,
                as_attachment=True,
                download_name="103_result.xlsx",
            )

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
