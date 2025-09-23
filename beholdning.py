from flask import Flask, render_template, request, flash, Response
import mysql.connector
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import io
import base64
import datetime
import csv

app = Flask(__name__)
app.secret_key = "hemmelig"

db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': 'Wgrgryjr3@',
    'database': 'glaslager'
}

materialer = [
    "kul_sække", "kul_bigbag",
    "jernoxyd", "koboltoxyd", "pyrit",
    "salpeter_bigbag", "zinkselenit", "chromit"
]

maaned_map = {
    "Januar":1,"Februar":2,"Marts":3,"April":4,"Maj":5,"Juni":6,
    "Juli":7,"August":8,"September":9,"Oktober":10,"November":11,"December":12
}

def format_tal(tal):
    try:
        return f"{tal:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return tal

# -------------------- BEHOLDNING --------------------
@app.route("/beholdning", methods=["GET","POST"])
def beholdning():
    start_dato_str = request.form.get("start_dato")
    slut_dato_str = request.form.get("slut_dato")

    # Standarddatoer
    start_dato = datetime.date(1900, 1, 1)
    slut_dato = datetime.date(3000, 1, 1)

    # Parse sikkert fra YYYY-MM-DD
    try:
        if start_dato_str:
            start_dato = datetime.datetime.strptime(start_dato_str, "%Y-%m-%d").date()
        if slut_dato_str:
            slut_dato = datetime.datetime.strptime(slut_dato_str, "%Y-%m-%d").date()
    except ValueError:
        flash("Datoformatet er ikke korrekt. Brug kalenderen til at vælge dato.")

    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM Lager ORDER BY registeret_aar, registeret_maaned, optalt_dato ASC")
    rækker = cursor.fetchall()
    cursor.close()
    conn.close()

    filtrerede = []
    for r in rækker:
        try:
            år = int(r['registeret_aar'])
            måned_num = maaned_map.get(r['registeret_maaned'], 1)
            r_dato = datetime.date(år, måned_num, 1)

            if start_dato <= r_dato <= slut_dato:
                kul_sække = r.get("kul_sække_total_kg") or 0
                kul_bigbag = r.get("kul_bigbag_total_kg") or 0
                r["kul_total_kg"] = kul_sække + kul_bigbag
                r["kul_total_kg_formatted"] = format_tal(r["kul_total_kg"])

                for m in materialer:
                    if m not in ["kul_sække","kul_bigbag"]:
                        key = f"{m}_total_kg"
                        r[f"{m}_total_kg_formatted"] = format_tal(r.get(key) or 0)

                filtrerede.append(r)
        except Exception as e:
            print("Fejl ved række:", r, e)
            continue

    # -------------------- Grafik --------------------
    graf_url = None
    if filtrerede:
        fig, ax = plt.subplots(figsize=(12,6))

        datoer = []
        kul_vals = []
        jern_vals = {m: [] for m in materialer if m not in ["kul_sække","kul_bigbag"]}

        for r in filtrerede:
            år = int(r['registeret_aar'])
            måned_num = int(r['registeret_maaned'])  # brug direkte tal
            dato = datetime.date(år, måned_num, 1)
            datoer.append(dato)

            kul_vals.append(r.get("kul_total_kg") or 0)

            for m in jern_vals.keys():
                jern_vals[m].append(r.get(f"{m}_total_kg") or 0)


        # Plot kul samlet
        ax.plot(datoer, kul_vals, marker='o', label="Kul samlet")

        # Plot de andre materialer
        for m, vals in jern_vals.items():
            ax.plot(datoer, vals, marker='o', label=m.replace("_"," "))

        # Formatér X-aksen
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
        ax.xaxis.set_major_locator(mdates.MonthLocator())

        # Sæt skala baseret på filter
        start = datetime.datetime.strptime(start_dato_str,"%Y-%m-%d").date() if start_dato_str else min(datoer)
        slut = datetime.datetime.strptime(slut_dato_str,"%Y-%m-%d").date() if slut_dato_str else max(datoer)
        ax.set_xlim(start, slut)

        fig.autofmt_xdate()
        ax.set_xlabel("Tid (filter: start - slut)")
        ax.set_ylabel("Kg")
        ax.legend()

        # Gem graf til base64
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        graf_url = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close(fig)


    return render_template(
        "beholdning.html",
        rækker=filtrerede,
        graf_url=graf_url,
        start_dato=start_dato_str,
        slut_dato=slut_dato_str,
        materialer=materialer
    )

# -------------------- BEHOLDNING EXPORT --------------------
@app.route("/beholdning/export", methods=["POST"])
def beholdning_export():
    start_dato_str = request.form.get("start_dato")
    slut_dato_str = request.form.get("slut_dato")

    start_dato = datetime.date(1900,1,1)
    slut_dato = datetime.date(3000,1,1)

    try:
        if start_dato_str:
            start_dato = datetime.datetime.strptime(start_dato_str,"%Y-%m-%d").date()
        if slut_dato_str:
            slut_dato = datetime.datetime.strptime(slut_dato_str,"%Y-%m-%d").date()
    except ValueError:
        pass

    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM Lager ORDER BY registeret_aar, registeret_maaned, optalt_dato ASC")
    rækker = cursor.fetchall()
    cursor.close()
    conn.close()

    filtrerede_rækker = []
    for r in rækker:
        try:
            år = int(r['registeret_aar'])
            måned_num = int(r['registeret_maaned'])
            r_dato = datetime.date(år,måned_num,1)

            if start_dato <= r_dato <= slut_dato:
                filtrerede_rækker.append(r)
        except:
            continue

    def generate():
        output = io.StringIO()
        writer = csv.writer(output)
        header = ["ID","Måned","År","Optalt af","Kul samlet"] + [m.replace("_"," ").capitalize() for m in materialer if m not in ["kul_bigbag","kul_sække"]]
        writer.writerow(header)
        yield output.getvalue()
        output.seek(0)
        output.truncate(0)

        for r in filtrerede_rækker:
            row = [
                r['id'],
                r['registeret_maaned'],
                r['registeret_aar'],
                r.get("optalt_af",""),
                (r.get("kul_sække_total_kg") or 0) + (r.get("kul_bigbag_total_kg") or 0)
            ]
            for m in materialer:
                if m not in ["kul_bigbag","kul_sække"]:
                    row.append(r.get(f"{m}_total_kg") or 0)
            writer.writerow(row)
            yield output.getvalue()
            output.seek(0)
            output.truncate(0)

    return Response(generate(), mimetype="text/csv",
                    headers={"Content-Disposition":"attachment; filename=beholdning.csv"})

# -------------------- RUN APP --------------------
if __name__=="__main__":
    app.run(debug=True)
