from pathlib import Path

from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = Path("uploads")
ALLOWED_EXTENSION = ".xlsx"


@app.route("/", methods=["GET", "POST"])
def home():
    message = ""
    message_type = ""

    if request.method == "POST":
        uploaded_file = request.files.get("excel_file")

        if not uploaded_file or uploaded_file.filename == "":
            message = "请选择 .xlsx 文件"
            message_type = "error"
        else:
            filename = secure_filename(uploaded_file.filename)
            file_extension = Path(filename).suffix.lower()

            if file_extension != ALLOWED_EXTENSION:
                message = "请选择 .xlsx 文件"
                message_type = "error"
            else:
                UPLOAD_FOLDER.mkdir(exist_ok=True)
                save_path = UPLOAD_FOLDER / filename
                uploaded_file.save(save_path)
                message = "文件上传成功"
                message_type = "success"

    return render_template("index.html", message=message, message_type=message_type)


if __name__ == "__main__":
    app.run(debug=True)
