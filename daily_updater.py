from datetime import datetime, timedelta
import json
import os
import re
import subprocess
import time
import urllib.parse
import chinese_calendar as chinesecalendar
import ddddocr
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
from playwright.sync_api import sync_playwright


def log_print(msg):
  """带具体时间显示的控制台输出函数"""
  current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
  print(f'[{current_time}] {msg}', flush=True)


def get_target_date_range_trade_days(
    start_date_str='2023-01-01', end_date_str=None
):
  """获取从 2023-01-01 到今天的指定日期范围内的所有实际交易日（自动排除周末及法定节假日）"""
  if not end_date_str:
    end_date_str = datetime.now().strftime('%Y-%m-%d')

  start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
  end_dt = datetime.strptime(end_date_str, '%Y-%m-%d')

  curr = start_dt
  trade_days = []
  while curr <= end_dt:
    if chinesecalendar.is_workday(curr) and curr.weekday() < 5:
      trade_days.append(curr.strftime('%Y-%m-%d'))
    curr += timedelta(days=1)

  return trade_days


def get_filtered_varieties():
  """硬编码期货品种名单，并应用剔除规则"""
  raw_list = [
      '1000中证/IM',
      '20号胶/NR',
      '300沪深/IF',
      '500中证/IC',
      '50上证/IH',
      'PP/PP',
      'PTA/TA',
      'PVC/V',
      '三十债/TL',
      '不锈钢/SS',
      '丙烯/PL',
      '乙二醇/EG',
      '二债/TS',
      '五债/TF',
      '低硫油/LU',
      '动力煤/ZC',
      '十债/T',
      '原木/LG',
      '原油/SC',
      '双胶纸/OP',
      '合成橡胶/BR',
      '商品指数/CI',
      '国际铜/BC',
      '塑料/L',
      '多晶硅/PS',
      '天然橡胶/RU',
      '对二甲苯/PX',
      '尿素/UR',
      '工业硅/SI',
      '强麦/WH',
      '早籼稻/RI',
      '晚籼稻/LR',
      '普麦/PM',
      '棉纱/CY',
      '棉花/CF',
      '棕榈油/P',
      '氧化铝/AO',
      '沥青/BU',
      '沪金/AU',
      '沪铅/PB',
      '沪铜/CU',
      '沪铝/AL',
      '沪银/AG',
      '沪锌/ZN',
      '沪锡/SN',
      '沪镍/NI',
      '油菜籽/RS',
      '液化气/PG',
      '淀粉/CS',
      '烧碱/SH',
      '热卷/HC',
      '焦炭/J',
      '焦煤/JM',
      '燃料油/FU',
      '玉米/C',
      '玻璃/FG',
      '瓶片/PR',
      '生猪/LH',
      '甲醇/MA',
      '白糖/SR',
      '短纤/PF',
      '硅铁/SF',
      '碳酸锂/LC',
      '粳稻/JR',
      '粳米/RR',
      '红枣/CJ',
      '纤维板/FB',
      '纯碱/SA',
      '纯苯/BZ',
      '纸浆/SP',
      '线材/WR',
      '胶合板/BB',
      '花生/PK',
      '苯乙烯/EB',
      '苹果/AP',
      '菜油/OI',
      '菜粕/RM',
      '螺纹钢/RB',
      '豆一/A',
      '豆二/B',
      '豆油/Y',
      '豆粕/M',
      '钯/PD',
      '铁矿石/I',
      '铂/PT',
      '铝合金/AD',
      '锰硅/SM',
      '集运欧线/EC',
      '鸡蛋/JD',
  ]

  excluded_codes = {
      'RI',
      'LR',
      'JR',
      'WH',
      'PM',
      'BB',
      'FB',
      'RS',
      'PT',
      'PD',
      'RR',
      'CI',
      'ZC',
      'SC',
      'LU',
      'WR',
      'ec',
      'EC',
      'LK',
  }
  excluded_names = {'原木', '原油', '线材', '集运欧线'}

  filtered_list = []
  for v in raw_list:
    variety_cn = v.split('/')[0].strip() if '/' in v else v.strip()
    code = v.split('/')[1].strip() if '/' in v else ''
    if code not in excluded_codes and variety_cn not in excluded_names:
      filtered_list.append(v)

  return sorted(filtered_list)


def parse_net_position(val_str):
  if not val_str or not isinstance(val_str, str):
    return 0
  clean_str = re.sub(r'[^\d.-]', '', val_str.replace(',', ''))
  try:
    num = float(clean_str)
  except ValueError:
    num = 0.0
  if '空' in val_str:
    return -abs(num)
  elif '多' in val_str:
    return abs(num)
  return num


def parse_increase_value(val_str):
  if not val_str or not isinstance(val_str, str):
    return 0.0
  clean_str = re.sub(r'[^\d.-]', '', val_str.replace(',', ''))
  try:
    return float(clean_str)
  except ValueError:
    return 0.0


def check_company_match(seat_name, target_companies):
  for company in target_companies:
    if company in seat_name:
      return True
  return False


def safe_goto(page, url, timeout=60000, wait_until='domcontentloaded'):
  for attempt in range(1, 4):
    try:
      page.goto(url, wait_until=wait_until, timeout=timeout)
      page.wait_for_timeout(1000)
      return True
    except Exception as e:
      log_print(f'⚠️ 打开页面超时或异常 ({e})，正在重试...')
      page.wait_for_timeout(1500)
  return False


def auto_login_and_close_ad(page, username, password):
  context = page.context

  # 1. 优先尝试使用 Cookie 注入免登录
  cookies_str = os.environ.get('COOKIES_JSON')
  if cookies_str:
    try:
      cookies = json.loads(cookies_str)
      context.add_cookies(cookies)
      log_print('🍪 成功加载环境变量中的 Cookie，尝试免登录访问...')
      safe_goto(page, 'https://jiaoyikecha.com/#/user/home')
      page.wait_for_timeout(4000)
      if 'login' not in page.url:
        log_print('🎉 Cookie 免登录成功！')
        try:
          close_btn = page.locator(
              '.el-dialog__headerbtn, .el-dialog__close, div[role="dialog"]'
              " button, [class*='close']"
          ).last
          if close_btn.is_visible(timeout=3000):
            close_btn.click(force=True)
            log_print('✅ 成功关闭广告弹窗！')
        except Exception:
          pass
        return True
      else:
        log_print(
            '⚠️ Cookie 已失效或不匹配，自动切回账号密码打码登录兜底流程...'
        )
    except Exception as e:
      log_print(f'⚠️ 解析 Cookie 异常: {e}')

  # 2. 兜底逻辑：原账号密码打码登录
  log_print('正在打开登录页面...')
  if not safe_goto(page, 'https://jiaoyikecha.com/#/user/login'):
    return False

  page.wait_for_timeout(3000)
  ocr = ddddocr.DdddOcr(show_ad=False)
  max_retries = 10

  for attempt in range(1, max_retries + 1):
    try:
      log_print(f'\n--- 第 {attempt} 次尝试登录 ---')

      user_locator = page.locator(
          "input[placeholder*='账号'], input[placeholder*='手机号'], input[type='text']"
      ).first
      user_locator.wait_for(state='visible', timeout=10000)

      pwd_locator = page.locator(
          "input[placeholder*='密码'], input[type='password']"
      ).first

      user_locator.click()
      user_locator.fill('')
      user_locator.type(username, delay=40)
      user_locator.dispatch_event('input')
      page.wait_for_timeout(400)

      pwd_locator.click()
      pwd_locator.fill('')
      pwd_locator.type(password, delay=40)
      pwd_locator.dispatch_event('input')
      page.wait_for_timeout(600)

      captcha_locator = page.locator(
          "img[src*='captcha'], img[class*='code'], .el-input + img, canvas,"
          ' form img'
      ).first
      captcha_locator.wait_for(state='visible', timeout=8000)
      page.wait_for_timeout(1000)

      captcha_bytes = captcha_locator.screenshot()
      captcha_code = ocr.classification(captcha_bytes)
      log_print(f'-> 【ddddocr】识别结果: {captcha_code}')

      code_locator = page.locator(
          "input[placeholder*='图形验证码'], input[placeholder*='验证码']"
      ).first
      code_locator.click()
      code_locator.fill('')
      code_locator.type(captcha_code, delay=40)
      code_locator.dispatch_event('input')
      page.wait_for_timeout(1000)

      login_btn = page.locator(
          ".el-button--primary, button:has-text('登录')"
      ).first
      if login_btn.count() > 0:
        box = login_btn.bounding_box()
        if box:
          page.mouse.click(
              box['x'] + box['width'] / 2, box['y'] + box['height'] / 2, force=True
          )
        else:
          login_btn.click(force=True)
      else:
        page.locator('button').last.click(force=True)

      log_print('⏳ 正在等待服务器响应登录与页面跳转（12秒缓冲）...')
      page.wait_for_timeout(12000)

      current_url = page.url
      log_print(f'-> 当前页面 URL: {current_url}')

      if 'login' not in current_url:
        log_print('🎉 恭喜！成功登录系统！')
        page.wait_for_timeout(3000)
        try:
          close_btn = page.locator(
              '.el-dialog__headerbtn, .el-dialog__close, div[role="dialog"]'
              " button, [class*='close']"
          ).last
          if close_btn.is_visible(timeout=4000):
            close_btn.click(force=True)
            page.wait_for_timeout(1000)
            log_print('✅ 成功关闭广告弹窗！')
        except Exception:
          pass
        return True
      else:
        log_print('⚠️ 登录未成功（仍停留在登录页），正在安全刷新验证码重试...')
        try:
          captcha_locator.click()
        except Exception:
          pass
        page.wait_for_timeout(3000)
    except Exception as e:
      log_print(f'❌ 登录尝试过程中发生异常: {e}')
      page.wait_for_timeout(3000)

  return False


def crawl_single_variety(page, date_str, raw_variety):
  v_name = (
      raw_variety.split('/')[0] if '/' in raw_variety else raw_variety
  ).strip()
  long_rows, short_rows = [], []

  for _ in range(3):
    try:
      encoded_variety = urllib.parse.quote(v_name)
      target_url = f'https://jiaoyikecha.com/#/variety/position/variety={encoded_variety}/code=all/date={date_str}'
      safe_goto(page, target_url)

      try:
        close_btn = page.locator(
            '.el-message-box__close, .el-dialog__headerbtn'
        ).first
        if close_btn.is_visible(timeout=800):
          close_btn.click(force=True)
      except Exception:
        pass

      start_time = time.time()
      while time.time() - start_time < 6:
        page.wait_for_timeout(400)
        is_rendered = page.evaluate("""
                    () => {
                        const bodyText = document.body.innerText || '';
                        const tables = document.querySelectorAll('table');
                        let hasTableWithData = false;
                        for (let t of tables) {
                            if (t.querySelectorAll('tr').length > 3) {
                                hasTableWithData = true;
                                break;
                            }
                        }
                        return (bodyText.includes('多头') || bodyText.includes('空头')) && hasTableWithData;
                    }
                """)
        if is_rendered:
          break

      table_data = page.evaluate("""
                () => {
                    let longData = [], shortData = [];
                    const tables = document.querySelectorAll('table');
                    tables.forEach((table) => {
                        let parent = table.parentElement, cardText = "";
                        for (let i = 0; i < 3; i++) {
                            if (parent) { cardText += " " + parent.innerText; parent = parent.parentElement; }
                        }
                        const rows = Array.from(table.querySelectorAll('tr')).map(tr => 
                            Array.from(tr.querySelectorAll('th, td')).map(td => td.innerText.trim())
                        ).filter(cells => cells.length > 0);
                        if ((cardText.includes('多头') || cardText.includes('多头持仓')) && longData.length === 0 && rows.length > 2) longData = rows;
                        else if ((cardText.includes('空头') || cardText.includes('空头持仓')) && shortData.length === 0 && rows.length > 2) shortData = rows;
                    });
                    if (longData.length === 0 && tables.length > 0) longData = Array.from(tables[0].querySelectorAll('tr')).map(tr => Array.from(tr.querySelectorAll('th, td')).map(td => td.innerText.trim())).filter(c => c.length > 0);
                    if (shortData.length === 0 && tables.length > 1) shortData = Array.from(tables[1].querySelectorAll('tr')).map(tr => Array.from(tr.querySelectorAll('th, td')).map(td => td.innerText.trim())).filter(c => c.length > 0);
                    return { long_table: longData, short_table: shortData };
                }
            """)
      long_rows, short_rows = table_data.get('long_table', []), table_data.get(
          'short_table', []
      )
      if len(long_rows) > 2 or len(short_rows) > 2:
        break
    except Exception:
      page.wait_for_timeout(800)

  def clean_rows(rows, is_long=True):
    lst = []
    for r in rows:
      if len(r) >= 2 and r[0] not in ['席位', '持仓', '合计', '']:
        net_val = r[3] if len(r) > 3 else ''
        lst.append({
            '席位/期货公司': r[0],
            '持仓量': r[1] if len(r) > 1 else '',
            '增减量': r[2] if len(r) > 2 else '',
            '净持仓': net_val,
            'num_val': parse_net_position(net_val),
        })
    df = pd.DataFrame(lst)
    if not df.empty:
      df = df.sort_values(by='num_val', ascending=is_long).reset_index(drop=True)
    return df

  df_long = clean_rows(long_rows, True)
  df_short = clean_rows(short_rows, False)
  return raw_variety, {'long': df_long, 'short': df_short}


def get_longhubang_summary_str(v_data):
  """格式化生成控制台输出的龙虎榜简要信息字符串"""
  df_l, df_s = v_data['long'], v_data['short']
  summary_parts = []

  if not df_l.empty:
    l_items = []
    for idx, row in df_l.head(2).iterrows():
      seat = row['席位/期货公司']
      pos = row['持仓量']
      chg = row['增减量']
      l_items.append(f'{seat}(持:{pos}/增:{chg})')
    summary_parts.append('多头Top2: ' + ' | '.join(l_items))
  else:
    summary_parts.append('多头Top2: 无数据')

  if not df_s.empty:
    s_items = []
    for idx, row in df_s.head(2).iterrows():
      seat = row['席位/期货公司']
      pos = row['持仓量']
      chg = row['增减量']
      s_items.append(f'{seat}(持:{pos}/增:{chg})')
    summary_parts.append('空头Top2: ' + ' | '.join(s_items))
  else:
    summary_parts.append('空头Top2: 无数据')

  return ' [ ' + ' ; '.join(summary_parts) + ' ]'


def export_excel_for_date(parsed_variety_data, date_str, output_folder):
  output_file = os.path.join(
      output_folder, f'全品种期货持仓龙虎榜汇总_{date_str}.xlsx'
  )
  group1 = ['永安期货', '国泰君安', '东证期货', '高盛期货']
  group2 = ['方正中期', '东方财富', '平安期货', '徽商期货']
  summary = []

  valid_variety_data = {}
  for v_name, v_data in parsed_variety_data.items():
    if not v_data['long'].empty or not v_data['short'].empty:
      valid_variety_data[v_name] = v_data

  if not valid_variety_data:
    log_print(f'⚠️ 日期 {date_str} 所有品种均未抓取到有效数据，跳过导出！')
    return False

  try:
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
      for v_name, v_data in valid_variety_data.items():
        df_l, df_s = v_data['long'], v_data['short']

        t10_l = df_l.head(10)['席位/期货公司'].tolist() if not df_l.empty else []
        t10_s = df_s.head(10)['席位/期货公司'].tolist() if not df_s.empty else []

        summary.append({
            '日期': datetime.strptime(date_str, '%Y-%m-%d').date(),
            '品种': v_name,
            '主力合约': '全合约',
            '组1多头Top10数': sum(1 for n in t10_l if check_company_match(n, group1)),
            '组1空头Top10数': sum(1 for n in t10_s if check_company_match(n, group1)),
            '组2多头Top10数': sum(1 for n in t10_l if check_company_match(n, group2)),
            '组2空头Top10数': sum(1 for n in t10_s if check_company_match(n, group2)),
        })

        safe_sheet_name = (
            re.sub(r'[\\/*?:[\]/]', '_', v_name)[:31]
            if v_name != '统计汇总'
            else '汇总'
        )

        if not df_l.empty:
          df_l.drop(columns=['num_val']).to_excel(
              writer,
              sheet_name=safe_sheet_name,
              startcol=0,
              startrow=1,
              index=False,
          )
        if not df_s.empty:
          df_s.drop(columns=['num_val']).to_excel(
              writer,
              sheet_name=safe_sheet_name,
              startcol=5,
              startrow=1,
              index=False,
          )

      if summary:
        pd.DataFrame(summary).to_excel(writer, sheet_name='统计汇总', index=False)
        wb = writer.book
        if '统计汇总' in wb.sheetnames:
          wb._sheets.insert(0, wb._sheets.pop(wb.sheetnames.index('统计汇总')))

    wb = openpyxl.load_workbook(output_file)
    header_fill = PatternFill(
        start_color='1F4E78', end_color='1F4E78', fill_type='solid'
    )
    header_font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9'),
    )

    for ws in wb.worksheets:
      for row_idx, row in enumerate(ws.iter_rows(), start=1):
        is_header = (ws.title == '统计汇总' and row_idx == 1) or (
            ws.title != '统计汇总' and row_idx == 2
        )
        for col_idx, cell in enumerate(row, start=1):
          cell.alignment = Alignment(
              horizontal='center', vertical='center', wrap_text=True
          )
          cell.border = thin_border
          if ws.title == '统计汇总' and row_idx > 1 and col_idx == 1:
            cell.number_format = 'yyyy-mm-dd'
          if is_header:
            cell.fill = header_fill
            cell.font = header_font
      for col in ws.columns:
        max_len = max(
            (
                sum(2 if ord(c) > 127 else 1 for c in str(cell.value or ''))
                for cell in col
            ),
            default=0,
        )
        ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 12)

    wb.save(output_file)
    log_print(f'✅ 数据文件已生成并美化: {output_file}')
    return True
  except Exception as e:
    log_print(f'❌ 导出 Excel 异常: {e}')
    return False


def main():
  log_print('=' * 60)
  log_print(
      '🚀 启动期货龙虎榜智能分时接力采集程序 (高稳定串行单线程版 + Cookie免登录)'
  )
  log_print('=' * 60)

  script_start_time = time.time()
  MAX_RUN_SECONDS = 5.5 * 3600

  target_dates = get_target_date_range_trade_days('2023-01-01')
  varieties = get_filtered_varieties()

  output_folder = 'futures_longhubang_data'
  os.makedirs(output_folder, exist_ok=True)

  pending_dates = []
  for d in target_dates:
    f_path = os.path.join(output_folder, f'全品种期货持仓龙虎榜汇总_{d}.xlsx')
    if not (os.path.exists(f_path) and os.path.getsize(f_path) > 1024):
      pending_dates.append(d)

  if not pending_dates:
    log_print('✨ 所有历史及今日交易日数据均已完整，无需重复抓取！')
    return

  log_print(
      f'📌 目标区间总交易日: {len(target_dates)} 天，当前待补全缺失交易日:'
      f' {len(pending_dates)} 天'
  )

  username = os.environ.get('USER_PHONE', '18392272125')
  password = os.environ.get('USER_PWD', 'lmn123')

  with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=[
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--ignore-certificate-errors',
        ],
    )
    context = browser.new_context(ignore_https_errors=True)
    page = context.new_page()
    page.set_default_timeout(60000)

    if not auto_login_and_close_ad(page, username, password):
      log_print('❌ 初始登录失败，退出程序！')
      browser.close()
      return

    for idx, date_str in enumerate(pending_dates, 1):
      elapsed_time = time.time() - script_start_time
      if elapsed_time > MAX_RUN_SECONDS:
        log_print(
            '⏰ 【限时保护触发】本次运行已满 5.5 小时，脚本已安全保存当前进度并退出。'
            '明天下午 16:46 将自动接力继续运行！'
        )
        break

      log_print(
          f"\n{'='*25} 正在处理待补全交易日 [{idx}/{len(pending_dates)}]:"
          f" {date_str} (已耗时: {elapsed_time/3600:.2f}小时) {'='*25}"
      )
      parsed_variety_data = {}
      total_v = len(varieties)

      for i, v in enumerate(varieties, 1):
        try:
          v_name, v_data = crawl_single_variety(page, date_str, v)
          if v_name:
            parsed_variety_data[v_name] = v_data
            summary_info = get_longhubang_summary_str(v_data)
            log_print(
                f'   📊 [{date_str}] 品种 [{i}/{total_v}] {v_name}'
                f' 抓取成功{summary_info}'
            )
        except Exception as e:
          log_print(f'❌ 品种 {v} 抓取异常: {e}')

      export_excel_for_date(parsed_variety_data, date_str, output_folder)

    browser.close()

  try:
    subprocess.run(
        ['git', 'config', '--global', 'user.name', 'GitHub Action Bot'],
        check=True,
    )
    subprocess.run(
        ['git', 'config', '--global', 'user.email', 'action@github.com'],
        check=True,
    )
    subprocess.run(['git', 'add', output_folder], check=True)
    subprocess.run(
        ['git', 'commit', -m, 'Auto-sync batch futures data update'],
        check=True,
    )
    subprocess.run(['git', 'push'], check=True)
    log_print('🚀 当前批次数据已成功自动同步推送到云端 GitHub 仓库！')
  except Exception as e:
    log_print(f'⚠️ Git 推送提示（若无变更可忽略）: {e}')


if __name__ == '__main__':
  main()
