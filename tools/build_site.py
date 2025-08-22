import argparse
from pathlib import Path
import pandas as pd
from jinja2 import Template


BASE_TEMPLATE = """
<!doctype html>
<html lang="id">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Report Kegiatan Program - YBM BRILiaN</title>
  <style>
    body { font-family: system-ui, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; color: #0f172a; }
    h1 { font-size: 20px; margin: 0 0 12px 0; }
    .meta { color: #64748b; font-size: 12px; margin-bottom: 16px; }
    table { border-collapse: collapse; width: 100%; font-size: 13px; }
    th, td { border: 1px solid #e2e8f0; padding: 8px 10px; text-align: left; }
    th { background: #f8fafc; position: sticky; top: 0; }
    tr:nth-child(even) td { background: #fafafa; }
    .container { max-width: 1200px; margin: 0 auto; }
  </style>
  <script>
    function filterTable() {
      const q = document.getElementById('q').value.toLowerCase();
      const rows = document.querySelectorAll('tbody tr');
      rows.forEach(r => {
        const txt = r.innerText.toLowerCase();
        r.style.display = txt.includes(q) ? '' : 'none';
      });
    }
  </script>
  </head>
<body>
  <div class="container">
    <h1>Report Kegiatan Program - YBM BRILiaN</h1>
    <div class="meta">Sumber: Excel terbaru via Microsoft Graph | Dibangun otomatis</div>
    <input id="q" type="text" placeholder="Cari..." oninput="filterTable()" style="padding:8px 10px; margin-bottom:10px; width: 320px;" />
    <div style="overflow:auto; max-height: 70vh; border:1px solid #e2e8f0;">
      <table>
        <thead>
          <tr>
          {% for col in columns %}
            <th>{{ col }}</th>
          {% endfor %}
          </tr>
        </thead>
        <tbody>
        {% for row in rows %}
          <tr>
            {% for col in columns %}
              <td>{{ row[col] }}</td>
            {% endfor %}
          </tr>
        {% endfor %}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--sheet', default=None)
    args = parser.parse_args()

    excel_path = Path(args.input)
    if not excel_path.exists():
        raise SystemExit(f'Excel not found: {excel_path}')

    df = pd.read_excel(excel_path, sheet_name=args.sheet)  # type: ignore
    df = df.fillna('')  # type: ignore
    columns = list(df.columns.astype(str))  # type: ignore
    rows = df.to_dict(orient='records')  # type: ignore

    html = Template(BASE_TEMPLATE).render(columns=columns, rows=rows)
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding='utf-8')
    print(f'Wrote {out}')


if __name__ == '__main__':
    main()
