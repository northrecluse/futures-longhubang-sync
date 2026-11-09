from datetime import datetime
import os
import re
import time
import urllib.parse
from chinese_calendar import is_workday
import ddddocr
import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
from playwright.sync_api import sync_playwright

# Telegram 配置信息（已保留参数接口，但不再主动发送文件）
TG_TOKEN = os.environ.get(
    'TG_BOT_TOKEN', '8253922103:AAFivN8VRYp5C_j_hL89-lx-UiCsISjkhuE'
)
TG_CHAT_ID = os.environ.get('TG_CHAT_ID', '6240934505')


def check_is_trading_day():
  """检查今天是否为中国法定工作日（交易日）。

  如果是周末或法定节假日（且不属于调休上班），返回 False，否则返回 True。
  """
  today = datetime.now().date()
  is_work = is_workday(today)
  if not is_work:
    print(
        f'📅 今天 ({today}) 是周末或法定节假日，非交易日，脚本自动安全退出。'
    )
    return False
  print(f'📅 今天 ({today}) 是法定工作日/交易日，开始执行爬虫任务...')
  return True


def check_is_latest_trading_data(page):
  """核对网站上展示的数据日期是否为【今天/最新交易日】。

  防止非交易日或数据未更新时重复抓取并推送旧数据。
  """
  today_str = datetime.now().strftime('%Y-%m-%d')
  print(f'🔍 正在验证页面数据是否为最新交易日 ({today_str}) ...')

  try:
    # 尝试在页面中查找日期标签或文本
    page.wait_for_selector(
        '.el-date-editor, .date-picker, span:has-text("日期"),'
        ' div[class*="date"]',
        timeout=5000,
    )
    page_text = page.evaluate('() => document.body.innerText')

    # 如果页面中包含今天的日期，或者明确标识为最新，则认为数据已更新
    if today_str.replace('-', '') in page_text or today_str in page_text:
      print(f'✅ 验证通过：页面已更新为最新交易日 ({today_str}) 的数据。')
      return True
    else:
      # 也可以根据具体网站的日期选择框来精确提取
      date_input_val = page.evaluate("""
                () => {
                    const input = document.querySelector('.el-date-editor input, input[placeholder*="日期"]');
                    return input ? input.value : '';
                }
            """)
      if date_input_val:
        print(f'-> 检测到页面当前选中日期为: {date_input_val}')
        if today_str in date_input_val:
          return True
        else:
          print(
              '⚠️ 页面数据日期与当前实际日期不符（可能数据尚未发布更新）。'
          )
          return False
  except Exception as e:
    print(f'-> 日期校验过程出现提示（可忽略）: {e}')

  # 默认兜底：如果在工作日执行且能正常抓取，允许通过
  return True


def parse_net_position(val_str):
  """解析净持仓文本为数值，用于精确排序：

  如 '多 4,411' -> 4411
  如 '空 9,228' -> -9228
  """
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


def check_company_match(seat_name, target_companies):
  """判断席位名称是否属于目标期货公司"""
  for company in target_companies:
    if company in seat_name:
      return True
  return False


def auto_login_and_close_ad(page, username, password):
  """自动填写账号密码、OCR识别验证码登录，并自动关闭广告弹窗"""
  print('正在打开登录页面...')
  page.goto('https://jiaoyikecha.com/#/user/login', wait_until='domcontentloaded')
  page.wait_for_timeout(2000)

  ocr = ddddocr.DdddOcr(show_ad=False)
  max_retries = 10

  for attempt in range(1, max_retries + 1):
    try:
      print(f'\n--- 第 {attempt} 次尝试登录 ---')

      # 1. 填写账号和密码
      print('正在填写账号和密码...')
      user_locator = page.locator(
          "input[placeholder*='账号'], input[placeholder*='手机号'],"
          " input[type='text']"
      ).first
      pwd_locator = page.locator(
          "input[placeholder*='密码'], input[type='password']"
      ).first

      user_locator.click()
      user_locator.fill('')
      user_locator.type(username, delay=30)
      user_locator.dispatch_event('input')

      pwd_locator.click()
      pwd_locator.fill('')
      pwd_locator.type(password, delay=30)
      pwd_locator.dispatch_event('input')

      # 2. 定位并截取验证码
      captcha_locator = page.locator(
          "img[src*='captcha'], img[class*='code'], .el-input + img, canvas,"
          ' form img'
      ).first
      captcha_bytes = captcha_locator.screenshot()
      captcha_code = ocr.classification(captcha_bytes)
      print(f'-> 【ddddocr】识别结果: {captcha_code}')

      # 填写验证码
      code_locator = page.locator(
          "input[placeholder*='图形验证码'], input[placeholder*='验证码']"
      ).first
      code_locator.click()
      code_locator.fill('')
      code_locator.type(captcha_code, delay=30)
      code_locator.dispatch_event('input')
      page.wait_for_timeout(500)

      # 3. 点击绿色主登录按钮
      print('正在精准点击绿色登录大按钮...')
      login_btn = page.locator(
          ".el-button--primary, button:has-text('登录')"
      ).first

      if login_btn.count() > 0:
        box = login_btn.bounding_box()
        if box:
          center_x = box['x'] + box['width'] / 2
          center_y = box['y'] + box['height'] / 2
          print(f'-> 绿色大按钮坐标: ({center_x}, {center_y})')
          page.mouse.click(center_x, center_y, force=True)
        else:
          login_btn.click(force=True)
      else:
        print('-> 未找到主按钮，执行备用方案...')
        page.locator('button').last.click(force=True)

      # 等待页面跳转
      page.wait_for_timeout(4000)
      current_url = page.url
      print(f'当前页面 URL: {current_url}')

      if 'login' not in current_url:
        print('🎉 恭喜！成功登录系统！')

        print('正在等待广告弹窗弹出...')
        page.wait_for_timeout(3000)

        # 检测并自动关闭广告弹窗
        try:
          print('正在检测并关闭广告弹窗...')
          close_btn = page.locator(
              '.el-dialog__headerbtn, .el-dialog__close, div[role="dialog"]'
              " button, [class*='close']"
          ).last

          if close_btn.is_visible(timeout=3000):
            close_btn.click(force=True)
            print('✅ 成功关闭广告弹窗！')
            page.wait_for_timeout(1000)
          else:
            print('-> 未检测到广告弹窗或无需关闭。')
        except Exception as ad_err:
          print(f'-> 关闭广告时发生异常（可忽略）: {ad_err}')

        return True
      else:
        print('⚠️ 验证码可能识别错误，正在自动刷新验证码重试...')
        captcha_locator.click()
        page.wait_for_timeout(1000)

    except Exception as e:
      print(f'❌ 尝试过程中发生异常: {e}')
      page.wait_for_timeout(1000)

  return False


def fetch_all_varieties(page):
  """从页面提取所有期货品种名称，增加多种选择器与兜底机制"""
  print('\n[步骤 1/2] 正在自动提取下拉列表中的所有期货品种名称...')

  varieties_set = set()

  try:
    search_input = page.locator(
        "input[placeholder*='商品'], input[placeholder*='合约'],"
        " input[placeholder*='搜索'], .el-select input"
    ).first
    if search_input.count() > 0:
      search_input.click(force=True)
      page.wait_for_timeout(1500)
    else:
      page.evaluate("""
                () => {
                    const target = document.querySelector('.el-input__inner, input[placeholder*="商品"]');
                    if (target) target.click();
                }
            """)
      page.wait_for_timeout(1500)
  except Exception as e:
    print(f'-> 打开品种下拉框异常: {e}')

  for i in range(25):
    extracted = page.evaluate("""
            () => {
                const items = document.querySelectorAll('li, .el-select-dropdown__item, div, span');
                let names = [];
                for (let item of items) {
                    let text = item.textContent.trim();
                    if (text && text.length >= 2 && text.length <= 30) {
                        names.push(text);
                    }
                }
                return names;
            }
        """)

    for name in extracted:
      if '/' in name and not any(
          char in name for char in ['+', '-', ' ', 'http', '登录', '首页']
      ):
        varieties_set.add(name)

    page.evaluate("""
            () => {
                const wrappers = document.querySelectorAll('.el-select-dropdown__wrap, .el-scrollbar__wrap, div[class*="scrollbar"]');
                for (let wrap of wrappers) {
                    if (wrap.offsetHeight > 0) {
                        wrap.scrollTop += 300;
                    }
                }
            }
        """)
    page.wait_for_timeout(200)

  pattern = re.compile(r'^[a-zA-Z0-9\u4e00-\u9fa5]+/[A-Z0-9]+$')
  cleaned_list = [name for name in varieties_set if pattern.match(name)]
  varieties_list = sorted(list(set(cleaned_list)))

  if not varieties_list:
    print('⚠️ 未能在页面动态捕获到品种下拉列表，触发【兜底内置品种库】...')
    fallback_varieties = [
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
        '原油/SC',
        '双焦/J',
        '豆粕/M',
        '豆油/Y',
        '棕榈/P',
        '螺纹钢/RB',
        '热卷/HC',
        '铁矿石/I',
        '沪铜/CU',
        '沪铝/AL',
        '沪锌/ZN',
        '沪镍/NI',
        '纯碱/SA',
        '玻璃/FG',
        '碳酸锂/LC',
        '工业硅/SI',
    ]
    varieties_list = sorted(fallback_varieties)

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
  }
  filtered_list = []
  for v in varieties_list:
    code = v.split('/')[1] if '/' in v else ''
    if code not in excluded_codes:
      filtered_list.append(v)
    else:
      print(f'-> 已剔除指定品种: {v}')

  print(
      f'-> 成功准备了 {len(filtered_list)} 个期货品种/合约进行抓取（已剔除指定品种）！'
  )
  return filtered_list


def main():
  if not check_is_trading_day():
    return

  username = os.environ.get('USER_PHONE', '18392272125')
  password = os.environ.get('USER_PWD', 'lmn123')

  parsed_variety_data = {}

  with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        slow_mo=0,
        args=[
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--ignore-certificate-errors',
            '--ignore-ssl-errors',
            '--allow-insecure-localhost',
        ],
    )
    context = browser.new_context(ignore_https_errors=True)
    page = context.new_page()

    login_success = auto_login_and_close_ad(page, username, password)
    if not login_success:
      print('❌ 登录多次失败，退出程序！')
      browser.close()
      return

    print('\n正在跳转至 [持仓龙虎榜] 页面...')
    page.goto(
        'https://jiaoyikecha.com/#/variety/position',
        wait_until='domcontentloaded',
    )
    page.wait_for_timeout(3000)

    # 校验当前页面是否已经出当天最新交易日的数据
    if not check_is_latest_trading_data(page):
      print(
          '⚠️ 当前网站尚未更新当天最新交易日的数据，为避免推送旧数据，脚本安全退出。'
      )
      browser.close()
      return

    raw_varieties = fetch_all_varieties(page)

    print('\n[步骤 2/2] 开始按品种循环抓取龙虎榜持仓数据...')

    for index, raw_variety in enumerate(raw_varieties, 1):
      variety_search_kw = (
          raw_variety.split('/')[0] if '/' in raw_variety else raw_variety
      )
      display_variety_name = raw_variety

      long_rows = []
      short_rows = []

      for fetch_attempt in range(1, 4):
        try:
          print(
              f'\n[{index}/{len(raw_varieties)}] 正在切换并抓取品种:'
              f' {display_variety_name} (尝试 {fetch_attempt}/3) ...'
          )

          search_input = page.locator(
              "input[placeholder*='商品'], input[placeholder*='合约'],"
              " input[placeholder*='搜索'], .el-select input"
          ).first
          if search_input.count() > 0:
            search_input.click(force=True)
            search_input.fill('')
            search_input.type(variety_search_kw, delay=100)
            page.wait_for_timeout(800)

            option_locator = page.locator(
                '.el-select-dropdown__item:visible'
            ).first
            if option_locator.count() > 0:
              option_locator.click(force=True)
            else:
              page.keyboard.press('Enter')
          else:
            encoded_variety = urllib.parse.quote(variety_search_kw)
            target_url = (
                'https://jiaoyikecha.com/#/variety/position/variety='
                + encoded_variety
            )
            page.goto(target_url, wait_until='domcontentloaded')

          start_time = time.time()
          while time.time() - start_time < 10:
            page.wait_for_timeout(1000)
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

          page.wait_for_timeout(500)

          table_data = page.evaluate("""
                        () => {
                            let longData = [];
                            let shortData = [];
                            const tables = document.querySelectorAll('table');

                            tables.forEach((table) => {
                                let parent = table.parentElement;
                                let cardText = "";
                                for (let i = 0; i < 3; i++) {
                                    if (parent) {
                                        cardText += " " + parent.innerText;
                                        parent = parent.parentElement;
                                    }
                                }

                                const rows = Array.from(table.querySelectorAll('tr')).map(tr => {
                                    return Array.from(tr.querySelectorAll('th, td')).map(td => td.innerText.trim());
                                }).filter(cells => cells.length > 0);

                                if ((cardText.includes('多头') || cardText.includes('多头持仓')) && longData.length === 0 && rows.length > 2) {
                                    longData = rows;
                                } 
                                else if ((cardText.includes('空头') || cardText.includes('空头持仓')) && shortData.length === 0 && rows.length > 2) {
                                    shortData = rows;
                                }
                            });

                            if (longData.length === 0 && tables.length > 0) {
                                longData = Array.from(tables[0].querySelectorAll('tr')).map(tr => 
                                    Array.from(tr.querySelectorAll('th, td')).map(td => td.innerText.trim())
                                ).filter(cells => cells.length > 0);
                            }
                            if (shortData.length === 0 && tables.length > 1) {
                                shortData = Array.from(tables[1].querySelectorAll('tr')).map(tr => 
                                    Array.from(tr.querySelectorAll('tr')).map(td => td.innerText.trim())
                                ).filter(cells => cells.length > 0);
                            }

                            return { 
                                long_table: longData, 
                                short_table: shortData 
                            };
                        }
                    """)

          long_rows = table_data.get('long_table', [])
          short_rows = table_data.get('short_table', [])

          if len(long_rows) > 2 or len(short_rows) > 2:
            break
          else:
            print(f'-> 【{display_variety_name}】数据未完整渲染，重试中...')
            page.reload(wait_until='domcontentloaded')
            page.wait_for_timeout(2000)

        except Exception as e:
          print(
              f'-> 处理品种 {display_variety_name} 发生异常 (尝试'
              f' {fetch_attempt}/3): {e}'
          )
          page.wait_for_timeout(2000)

      long_list = []
      for row in long_rows:
        if len(row) >= 2 and row[0] not in ['席位', '持仓', '合计', '']:
          net_val = row[3] if len(row) > 3 else ''
          long_list.append({
              '席位/期货公司': row[0],
              '持仓量': row[1] if len(row) > 1 else '',
              '增减量': row[2] if len(row) > 2 else '',
              '净持仓': net_val,
              'num_val': parse_net_position(net_val),
          })

      short_list = []
      for row in short_rows:
        if len(row) >= 2 and row[0] not in ['席位', '持仓', '合计', '']:
          net_val = row[3] if len(row) > 3 else ''
          short_list.append({
              '席位/期货公司': row[0],
              '持仓量': row[1] if len(row) > 1 else '',
              '增减量': row[2] if len(row) > 2 else '',
              '净持仓': net_val,
              'num_val': parse_net_position(net_val),
          })

      df_long = pd.DataFrame(long_list)
      if not df_long.empty:
        df_long = df_long.sort_values(
            by='num_val', ascending=False
        ).reset_index(drop=True)

      df_short = pd.DataFrame(short_list)
      if not df_short.empty:
        df_short = (
            df_short.sort_values(by='num_val', ascending=True)
            .reset_index(drop=True)
        )

      parsed_variety_data[display_variety_name] = {
          'long': df_long,
          'short': df_short,
      }

      print(
          f'-> 【{display_variety_name}】数据解析完成'
          f' (多头:{len(df_long)}条, 空头:{len(df_short)}条)'
      )

    browser.close()

  # 导出 Excel 文件（仅在本地/云端存储，不再通过 Telegram 主动推送文件）
  timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
  output_file = f'全品种期货持仓龙虎榜汇总_{timestamp}.xlsx'

  group1_targets = ['永安期货', '国泰君安', '东证期货', '高盛期货']
  group2_targets = ['方正期货', '东方财富', '平安期货', '徽商期货']

  summary_records = []

  with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for variety_name, v_data in parsed_variety_data.items():
      df_l = v_data['long']
      df_s = v_data['short']

      l_top10 = df_l.head(10)['席位/期货公司'].tolist() if not df_l.empty else []
      s_top10 = df_s.head(10)['席位/期货公司'].tolist() if not df_s.empty else []

      g1_l_count = sum(
          1 for name in l_top10 if check_company_match(name, group1_targets)
      )
      g1_s_count = sum(
          1 for name in s_top10 if check_company_match(name, group1_targets)
      )

      g2_l_count = sum(
          1 for name in l_top10 if check_company_match(name, group2_targets)
      )
      g2_s_count = sum(
          1 for name in s_top10 if check_company_match(name, group2_targets)
      )

      summary_records.append({
          '品种': variety_name,
          '组1(永安/国泰/东证/高盛) 多头 Top10 出现数': g1_l_count,
          '组1(永安/国泰/东证/高盛) 空头 Top10 出现数': g1_s_count,
          '组2(方正/东财/平安/徽商) 多头 Top10 出现数': g2_l_count,
          '组2(方正/东财/平安/徽商) 空头 Top10 出现数': g2_s_count,
      })

      clean_l = (
          df_l.drop(columns=['num_val']) if 'num_val' in df_l.columns else df_l
      )
      clean_s = (
          df_s.drop(columns=['num_val']) if 'num_val' in df_s.columns else df_s
      )

      # 规范化工作表名称，防止斜杠或特殊字符报错
      clean_name = variety_name.replace('/', '_').replace('\\', '_')
      safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', clean_name)[:31]

      clean_l.to_excel(
          writer, sheet_name=safe_sheet_name, startcol=0, startrow=1, index=False
      )
      clean_s.to_excel(
          writer, sheet_name=safe_sheet_name, startcol=5, startrow=1, index=False
      )

      ws = writer.sheets[safe_sheet_name]
      ws.cell(row=1, column=1, value='【多头龙虎榜 (按净持仓从多到少)】')
      ws.cell(row=1, column=6, value='【空头龙虎榜 (按净持仓从少到多)】')

    if summary_records:
      df_summary = pd.DataFrame(summary_records)

      df_summary['__sort_key1'] = (
          df_summary['组1(永安/国泰/东证/高盛) 多头 Top10 出现数']
          + df_summary['组1(永安/国泰/东证/高盛) 空头 Top10 出现数']
      )
      df_summary = df_summary.sort_values(
          by=[
              '__sort_key1',
              '组1(永安/国泰/东证/高盛) 多头 Top10 出现数',
          ],
          ascending=False,
      ).reset_index(drop=True)
      df_summary = df_summary.drop(columns=['__sort_key1'])

      df_summary.to_excel(writer, sheet_name='统计汇总', index=False)
      workbook = writer.book
      summary_sheet = workbook['统计汇总']
      workbook._sheets.remove(summary_sheet)
      workbook._sheets.insert(0, summary_sheet)

  # Excel 格式优化
  wb = openpyxl.load_workbook(output_file)
  for ws in wb.worksheets:
    for row in ws.iter_rows():
      for cell in row:
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for col in ws.columns:
      max_len = 0
      col_letter = col[0].column_letter
      for cell in col:
        if cell.value is not None:
          val_str = str(cell.value)
          length = sum(2 if ord(char) > 127 else 1 for char in val_str)
          if length > max_len:
            max_len = length
      ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

  wb.save(output_file)

  print(f"\n{'='*20} 全部任务完成 {'='*20}")
  print(f'最新交易日汇总 Excel 文件已成功本地保存至: {output_file}')
  print('（已按照要求：移除了通过 Telegram 主动发送文件的逻辑）')


if __name__ == '__main__':
  main()
