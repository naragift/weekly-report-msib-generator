from flask import Flask, request, send_file, render_template
from io import BytesIO
import pandas as pd
import requests


app = Flask(__name__, template_folder=".")


def get_report_excel(email, password):
    login_url = "https://api.kampusmerdeka.kemdikbud.go.id/user/auth/login/mbkm"
    login_payload = {
        "email": email,
        "password": password
    }

    session = requests.Session()

    login_response = session.post(login_url, json=login_payload)
    if login_response.status_code != 200:
        return

    access_token = login_response.json()["data"]["access_token"]

    headers = {
        "authorization": f"Bearer {access_token}"
    }

    activity_url = "https://api.kampusmerdeka.kemdikbud.go.id/mbkm/mahasiswa/activities"
    id_kegiatan_result = requests.get(activity_url, headers=headers)
    response_data = id_kegiatan_result.json().get("data", [])
    response_data = sorted(response_data, key=lambda x: x.get(
        'akhir_kegiatan'), reverse=True)

    if response_data:
        id_kegiatan = response_data[0]["id"]
    else:
        return

    base_url = f"https://api.kampusmerdeka.kemdikbud.go.id/magang/report/perweek/{id_kegiatan}/"
    data_list = []

    for i in range(1, 21):
        url = base_url + str(i)

        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            data = response.json()
            learned_weekly = data["data"]["learned_weekly"]
            data_list.append({"Minggu": i, "Laporan Mingguan": learned_weekly})
        else:
            return

    df = pd.DataFrame(data_list)

    output_buf = BytesIO()
    df.to_excel(output_buf, index=False)

    output_buf.seek(0)

    return output_buf


@app.route("/process", methods=["POST"])
def process():
    email = request.form.get("email")
    password = request.form.get("password")

    if not email or not password:
        return "salah"

    output_buf = get_report_excel(email, password)
    if not output_buf:
        return "gagal"

    return send_file(output_buf, as_attachment=True, download_name="Weekly Report.xlsx", mimetype="application/ms-excel")


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
