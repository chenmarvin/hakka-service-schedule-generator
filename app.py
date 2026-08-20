import json
import os
import sys
import traceback
from collections import defaultdict

from flask import Flask, render_template, request, jsonify, send_file

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from generate_schedule import (
    quarter_sundays, build_tables, solve_schedule,
    compute_scores, export_xlsx,
)

app = Flask(__name__)
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')


@app.route('/')
def index():
    with open(CONFIG_PATH, encoding='utf-8') as f:
        config = json.load(f)
    return render_template('index.html', config=config)


@app.route('/generate', methods=['POST'])
def generate():
    try:
        cfg = request.get_json(force=True)

        # Ensure all constraint keys exist
        c = cfg.setdefault('constraints', {})
        for key in ('max_per_role', 'available_months', 'hard_pairs',
                    'no_same_day', 'unavailable_dates'):
            c.setdefault(key, [])

        # Persist
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

        year       = cfg['year']
        quarter    = cfg['quarter']
        mc_pool    = cfg['roles']['司會']
        piano_pool = cfg['roles']['司琴']
        proj_pool  = cfg['roles']['投影']
        all_people = list(dict.fromkeys(mc_pool + piano_pool + proj_pool))

        sundays    = quarter_sundays(year, quarter)
        cfg_tables = build_tables(cfg, sundays)
        _, _, _, _, sunday_months, _ = cfg_tables

        schedule = solve_schedule(
            sundays, mc_pool, piano_pool, proj_pool,
            all_people, cfg_tables, sunday_months
        )
        scores = compute_scores(schedule, all_people)

        q_names   = {1: '第一季', 2: '第二季', 3: '第三季', 4: '第四季'}
        xlsx_name = f'恩霖堂{year}{q_names[quarter]}主日服事表_客語堂.xlsx'
        xlsx_path = os.path.join(BASE_DIR, xlsx_name)
        export_xlsx(schedule, scores, cfg, sundays, xlsx_path)

        paired_mcs = {item['mc'] for item in cfg['constraints'].get('hard_pairs', [])}

        rows = []
        for day in schedule:
            rows.append({
                'date':        f"{day['date'].month}/{day['date'].day}",
                'mc':          day['司會'],
                'piano':       day['司琴'],
                'proj':        day['投影'],
                'mc_score':    scores[day['司會']],
                'piano_score': scores[day['司琴']],
                'proj_score':  scores[day['投影']],
                'paired':      day['司會'] in paired_mcs,
            })

        mc_counts    = defaultdict(int)
        piano_counts = defaultdict(int)
        proj_counts  = defaultdict(int)
        for day in schedule:
            mc_counts[day['司會']]    += 1
            piano_counts[day['司琴']] += 1
            proj_counts[day['投影']]  += 1

        vals = list(scores.values())

        return jsonify({
            'ok':           True,
            'rows':         rows,
            'scores':       scores,
            'mc_counts':    dict(mc_counts),
            'piano_counts': dict(piano_counts),
            'proj_counts':  dict(proj_counts),
            'spread':       max(vals) - min(vals),
            'xlsx_name':    xlsx_name,
        })

    except Exception as e:
        return jsonify({'ok': False, 'error': str(e),
                        'trace': traceback.format_exc()})


@app.route('/download/<path:filename>')
def download(filename):
    path = os.path.join(BASE_DIR, filename)
    if not os.path.exists(path):
        return 'File not found', 404
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
