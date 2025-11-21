#!/usr/bin/env python3
"""
シンプルシフト作成アプリ
スタッフが個別に希望を入力し、全員完了で自動最適化
"""

from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from datetime import datetime
from pathlib import Path
import subprocess
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from excel_manager import ExcelManager

app = Flask(__name__)
app.secret_key = 'simple_shift_app_2024'

# 基本設定
BASE_DIR = Path(__file__).parent
excel_mgr = ExcelManager(base_path='output')
SHIFT_TYPES = ['早番', '中番', '遅番', '休み', '有給', '半休']

# 現在の月を取得
def get_current_month():
    now = datetime.now()
    return f"{now.year}年{now.month}月"


@app.route('/')
def index():
    """トップページ"""
    year_month = get_current_month()
    shift_exists = excel_mgr.shift_exists(year_month)

    # 提出状況を取得
    staff_list = []
    all_submitted = False
    optimized = False

    if shift_exists:
        staff_list = excel_mgr.get_staff_list(year_month)
        all_submitted = excel_mgr.check_all_submitted(year_month)
        output_file = Path('output') / year_month / f"{year_month}_最適化シフト_完成版.xlsx"
        optimized = output_file.exists()

    return render_template('index.html',
                         year_month=year_month,
                         shift_exists=shift_exists,
                         staff_list=staff_list,
                         all_submitted=all_submitted,
                         optimized=optimized)


@app.route('/setup', methods=['GET', 'POST'])
def setup():
    """初回設定：スタッフ登録"""
    year_month = get_current_month()

    if request.method == 'POST':
        try:
            # スタッフ名を取得
            staff_names = []
            for key in request.form.keys():
                if key.startswith('staff_name_'):
                    name = request.form[key].strip()
                    if name:
                        staff_names.append(name)

            if len(staff_names) == 0:
                flash('スタッフ名を最低1名入力してください', 'error')
                return redirect(url_for('setup'))

            # 月次シフトを作成
            excel_mgr.create_month_shift(year_month, staff_names)

            flash(f'{year_month}のスタッフ登録が完了しました！', 'success')
            return redirect(url_for('index'))

        except Exception as e:
            flash(f'エラーが発生しました: {str(e)}', 'error')
            return redirect(url_for('setup'))

    # GET: フォーム表示
    return render_template('setup.html', year_month=year_month)


@app.route('/input/<staff_name>')
def input_form(staff_name):
    """希望入力フォーム"""
    year_month = get_current_month()

    if not excel_mgr.shift_exists(year_month):
        flash('まだスタッフ登録されていません。初回設定を行ってください。', 'error')
        return redirect(url_for('index'))

    # 日付情報を取得
    dates = excel_mgr.get_month_dates(year_month)

    # 既存の希望があれば読み込み
    existing_preferences = excel_mgr.load_staff_preferences(year_month, staff_name)

    return render_template('input.html',
                         year_month=year_month,
                         staff_name=staff_name,
                         dates=dates,
                         shift_types=SHIFT_TYPES,
                         existing_preferences=existing_preferences)


@app.route('/submit', methods=['POST'])
def submit():
    """希望を提出"""
    try:
        year_month = get_current_month()
        staff_name = request.form.get('staff_name')

        if not staff_name:
            flash('スタッフ名が不正です', 'error')
            return redirect(url_for('index'))

        # 希望データを取得
        preferences = {}
        for key, value in request.form.items():
            if key.startswith('shift_'):
                day = int(key.split('_')[1])
                preferences[day] = value

        # データを保存
        excel_mgr.save_staff_preferences(year_month, staff_name, preferences)

        flash(f'{staff_name}さんの希望を提出しました！', 'success')

        # 全員提出済みかチェック
        if excel_mgr.check_all_submitted(year_month):
            flash('全員の提出が完了しました！シフトを最適化しています...', 'info')

            # 自動最適化を実行
            result = run_optimizer(year_month)

            if result['success']:
                flash('シフトの最適化が完了しました！', 'success')
                return redirect(url_for('complete'))
            else:
                flash(f'最適化に失敗しました: {result["error"]}', 'error')

        return redirect(url_for('index'))

    except Exception as e:
        flash(f'エラーが発生しました: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/complete')
def complete():
    """完了画面"""
    year_month = get_current_month()
    output_file = Path('output') / year_month / f"{year_month}_最適化シフト_完成版.xlsx"

    if not output_file.exists():
        flash('最適化されたシフトファイルが見つかりません', 'error')
        return redirect(url_for('index'))

    return render_template('complete.html', year_month=year_month)


@app.route('/manage', methods=['GET', 'POST'])
def manage_staff():
    """スタッフ管理画面"""
    year_month = get_current_month()

    if not excel_mgr.shift_exists(year_month):
        flash('まだスタッフ登録されていません', 'error')
        return redirect(url_for('setup'))

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            # スタッフ追加
            staff_name = request.form.get('staff_name', '').strip()
            if staff_name:
                if excel_mgr.add_staff(year_month, staff_name):
                    flash(f'{staff_name}を追加しました', 'success')
                else:
                    flash(f'{staff_name}は既に登録されています', 'error')
            else:
                flash('スタッフ名を入力してください', 'error')

        elif action == 'remove':
            # スタッフ削除
            staff_name = request.form.get('staff_name', '').strip()
            if staff_name:
                if excel_mgr.remove_staff(year_month, staff_name):
                    flash(f'{staff_name}を削除しました', 'success')
                else:
                    flash(f'{staff_name}が見つかりません', 'error')

        return redirect(url_for('manage_staff'))

    # GET: スタッフ一覧を表示
    staff_list = excel_mgr.get_staff_list(year_month)
    return render_template('staff_manage.html',
                         year_month=year_month,
                         staff_list=staff_list)


@app.route('/download')
def download():
    """最適化されたシフトをダウンロード"""
    year_month = get_current_month()
    output_file = Path('output') / year_month / f"{year_month}_最適化シフト_完成版.xlsx"

    if not output_file.exists():
        flash('ファイルが見つかりません', 'error')
        return redirect(url_for('index'))

    return send_file(
        output_file,
        as_attachment=True,
        download_name=f"{year_month}_最適化シフト.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def run_optimizer(year_month):
    """最適化を実行"""
    try:
        # 全員の希望を統合してExcelファイルを作成
        all_preferences = excel_mgr.get_all_preferences(year_month)
        staff_list = excel_mgr.get_staff_list(year_month)
        staff_names = [s['name'] for s in staff_list]

        # 入力ファイルを作成
        month_folder = Path('output') / year_month
        input_file = month_folder / f"{year_month}.xlsx"
        create_input_excel(input_file, staff_names, year_month, all_preferences)

        # 最適化スクリプトを実行
        optimizer_path = BASE_DIR / 'Shift_optimizer.py'
        python_exe = sys.executable  # 仮想環境のPythonを使用

        result = subprocess.run(
            [python_exe, str(optimizer_path), year_month],
            capture_output=True,
            cwd=str(BASE_DIR)
        )

        # バイトデータをデコード（エラーを無視）
        stdout = result.stdout.decode('cp932', errors='ignore') if result.stdout else ''
        stderr = result.stderr.decode('cp932', errors='ignore') if result.stderr else ''

        print("=== Optimizer Output ===")
        print(stdout)
        if stderr:
            print("=== Optimizer Errors ===")
            print(stderr)

        if result.returncode == 0:
            return {'success': True}
        else:
            return {'success': False, 'error': stderr or stdout}

    except Exception as e:
        return {'success': False, 'error': str(e)}


def create_input_excel(file_path, staff_names, year_month, all_preferences):
    """統合Excelファイルを作成"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "希望シフト"

    # スタイル設定
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')

    # 日付情報を取得
    dates = excel_mgr.get_month_dates(year_month)

    # ヘッダー行
    ws.cell(row=1, column=1, value='氏名').font = header_font
    ws.cell(row=1, column=1).border = thin_border
    ws.cell(row=1, column=1).alignment = center_align

    for idx, date_info in enumerate(dates, start=2):
        cell = ws.cell(row=1, column=idx, value=date_info['date'])
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_align
        cell.number_format = 'M/D'

    # スタッフデータ
    for staff_idx, staff_name in enumerate(staff_names, start=2):
        cell = ws.cell(row=staff_idx, column=1, value=staff_name)
        cell.border = thin_border
        cell.alignment = center_align

        preferences = all_preferences.get(staff_name, {})
        for date_idx, date_info in enumerate(dates, start=2):
            day = date_info['day']
            shift_value = preferences.get(day, 'どちらでも')

            cell = ws.cell(row=staff_idx, column=date_idx, value=shift_value)
            cell.border = thin_border
            cell.alignment = center_align

    # 列幅調整
    ws.column_dimensions['A'].width = 15
    for col in range(2, len(dates) + 2):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 10

    wb.save(file_path)
    print(f"入力ファイルを作成: {file_path}")


if __name__ == '__main__':
    print("=" * 60)
    print("シンプルシフト作成アプリ")
    print("=" * 60)
    print(f"作業ディレクトリ: {BASE_DIR}")
    print(f"出力ディレクトリ: output/")
    print("ブラウザで http://localhost:5000 にアクセス")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)
