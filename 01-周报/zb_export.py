import time
import os
import json
import datetime
import openpyxl
from openpyxl import Workbook
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.edge.options import Options

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "zb_export_config.json"
SMARTBI_BASE_URL = "https://bi.61info.cn/smartbi"
DOWNLOAD_DIR = str(BASE_DIR)
SUMMARY_OUTPUT_PATH = BASE_DIR / "周报结论汇总.xlsx"
TEAM_ORDER = ["海外团队", "欧美澳", "港澳", "台湾"]
PAGE_READY_TIMEOUT_SECONDS = 120
PAGE_READY_POLL_SECONDS = 2
EXPORT_DIALOG_TIMEOUT_SECONDS = 30
EXPORT_DIALOG_POLL_SECONDS = 1
POST_EXCEL_SELECT_WAIT_SECONDS = 2
DOWNLOAD_READY_TIMEOUT_SECONDS = 120
DOWNLOAD_POLL_SECONDS = 1
FILE_STABLE_CHECK_ROUNDS = 2
REPORT_RULES = [
    {"label": "1）一单结课续费率", "extract_mode": "fixed_rows_i"},
    {"label": "2）统合结课续费率", "extract_mode": "scan_rows_m"},
    {"label": "3）一续升舱续费率", "extract_mode": "summary_block_h"},
    {"label": "4）统合早鸟续费率", "extract_mode": "summary_block_h"},
]


def load_auth_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"未找到配置文件: {CONFIG_PATH}")

    try:
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"配置文件 JSON 解析失败: {CONFIG_PATH}") from exc

    username = str(config.get("username", "")).strip()
    password = str(config.get("password", "")).strip()
    report_ids = config.get("report_ids", [])
    if not username or not password:
        raise ValueError(f"配置文件缺少 username 或 password: {CONFIG_PATH}")
    if not isinstance(report_ids, list) or len(report_ids) < len(REPORT_RULES):
        raise ValueError(f"配置文件缺少至少 {len(REPORT_RULES)} 个 report_ids: {CONFIG_PATH}")

    normalized_report_ids = [str(item).strip() for item in report_ids if str(item).strip()]
    if len(normalized_report_ids) < len(REPORT_RULES):
        raise ValueError(f"配置文件中的 report_ids 至少需要 {len(REPORT_RULES)} 个有效报表ID: {CONFIG_PATH}")

    return username, password, normalized_report_ids


def build_report_url(report_id):
    return f"{SMARTBI_BASE_URL}/vision/openresource.jsp?resid={report_id}"

def get_timestamp_filename(original_name):
    """在文件名后添加年月日时间戳"""
    # 获取当前日期
    today = datetime.datetime.now()
    date_str = today.strftime("%Y%m%d")  # 格式：20260505
    
    # 分离文件名和扩展名
    name, ext = os.path.splitext(original_name)
    
    # 构建新文件名：原文件名 + 日期 + 扩展名
    new_name = f"{name}--{date_str}{ext}"
    return new_name


def format_rate(value):
    if value is None:
        return "%"
    if isinstance(value, (int, float)):
        return f"{value * 100:.2f}%"
    return str(value)


def build_conclusion_line(label, team_rates):
    return (
        f"{label}：整体 {team_rates.get('海外团队', '%')}（MTD目标 %，月目标 %），"
        f"欧美澳 {team_rates.get('欧美澳', '%')}（MTD目标 %，月目标 %），"
        f"港澳 {team_rates.get('港澳', '%')}（MTD目标 %，月目标 %），"
        f"台湾 {team_rates.get('台湾', '%')}（MTD目标 %，月目标 %）；"
    )


def write_summary_workbook(conclusion_lines):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    default_font = openpyxl.styles.Font(name='微软雅黑', size=11)

    for row_index, line in enumerate(conclusion_lines, start=1):
        cell = sheet.cell(row=row_index, column=1, value=line)
        cell.font = default_font

    sheet.column_dimensions["A"].width = 120
    workbook.save(SUMMARY_OUTPUT_PATH)


def extract_team_rates_fixed_rows(sheet1):
    team_rows = {
        "海外团队": 6,
        "港澳": 7,
        "台湾": 8,
        "欧美澳": 9,
    }
    team_rates = {team: "%" for team in TEAM_ORDER}
    print("- 数据提取详情:")
    for team_name, row_num in team_rows.items():
        if row_num <= sheet1.max_row:
            cell_b = sheet1.cell(row=row_num, column=2).value
            cell_i = sheet1.cell(row=row_num, column=9).value
            print(f"  行{row_num}: B{row_num}='{cell_b}', I{row_num}='{cell_i}'")
            team_rates[team_name] = format_rate(cell_i)
    return team_rates


def extract_team_rates_scan_rows(sheet1):
    team_rates = {}
    print("- 数据提取详情:")
    for row in range(6, min(30, sheet1.max_row + 1)):
        cell_b = sheet1.cell(row=row, column=2).value
        cell_m = sheet1.cell(row=row, column=13).value
        if cell_b:
            cell_b_str = str(cell_b).strip()
            for team in TEAM_ORDER:
                if team in cell_b_str and team not in team_rates:
                    rate_str = format_rate(cell_m)
                    team_rates[team] = rate_str
                    print(f"  行{row}: B{row}='{cell_b}', M{row}='{cell_m}' → {team}: {rate_str}")
    return {team: team_rates.get(team, "%") for team in TEAM_ORDER}


def extract_team_rates_summary_block(sheet1):
    team_rows = {
        "海外团队": 5,
        "港澳": 6,
        "台湾": 7,
        "欧美澳": 8,
    }
    team_rates = {team: "%" for team in TEAM_ORDER}
    target_column = None
    target_headers = {"当前续费率", "目前续费率"}
    for column_index in range(1, sheet1.max_column + 1):
        header_value = sheet1.cell(row=4, column=column_index).value
        header_text = str(header_value).strip() if header_value is not None else ""
        if header_text in target_headers:
            target_column = column_index
            break

    if target_column is None:
        raise ValueError("未在汇总板块找到【当前续费率/目前续费率】字段。")

    print("- 数据提取详情:")
    for team_name, row_num in team_rows.items():
        if row_num <= sheet1.max_row:
            cell_b = sheet1.cell(row=row_num, column=2).value
            cell_value = sheet1.cell(row=row_num, column=target_column).value
            print(f"  行{row_num}: B{row_num}='{cell_b}', col{target_column}='{cell_value}'")
            team_rates[team_name] = format_rate(cell_value)
    return team_rates


def analyze_report_and_build_conclusion(file_path, report_rule):
    """分析报表数据并生成结论文本"""
    try:
        workbook_data = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        sheet1 = workbook_data.active

        print(f"\n📋 报表结构分析：{report_rule['label']}")
        print(f"- 表格行数: {sheet1.max_row}")
        print(f"- 表格列数: {sheet1.max_column}")

        if report_rule["extract_mode"] == "fixed_rows_i":
            team_rates = extract_team_rates_fixed_rows(sheet1)
        elif report_rule["extract_mode"] == "summary_block_h":
            team_rates = extract_team_rates_summary_block(sheet1)
        else:
            team_rates = extract_team_rates_scan_rows(sheet1)

        workbook_data.close()

        conclusion_line = build_conclusion_line(report_rule["label"], team_rates)
        print("✓ 已生成报表结论")

        print("\n📊 分析结论预览：")
        print(f"  {conclusion_line}")

        print("\n📈 数据提取结果：")
        for team in TEAM_ORDER:
            print(f"  - {team}续费率: {team_rates.get(team, '%')}")

        return conclusion_line

    except Exception as e:
        print(f"⚠ 分析报表失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def wait_for_report_ready(driver, step_label):
    print(f"\n{step_label}: 等待页面加载")
    max_checks = max(1, PAGE_READY_TIMEOUT_SECONDS // PAGE_READY_POLL_SECONDS)
    for i in range(max_checks):
        time.sleep(PAGE_READY_POLL_SECONDS)
        try:
            page_source = driver.page_source
            if "btnExport" in page_source or "导出" in page_source:
                print(f"  ✓ 页面加载完成 (第{i+1}次检查)")
                return
        except Exception as e:
            print(f"  检查失败: {e}")
        print(f"  等待中... ({i+1}/{max_checks})")
    raise TimeoutError(f"{step_label} 超时，页面未出现导出入口。")


def click_export_controls(driver, export_script, excel_script, online_export_script, step_base):
    print(f"\n步骤{step_base}: 点击导出按钮")
    result = driver.execute_script(export_script)
    print(f"  {result}")

    print("  等待导出弹窗加载...")
    max_checks = max(1, EXPORT_DIALOG_TIMEOUT_SECONDS // EXPORT_DIALOG_POLL_SECONDS)
    for i in range(max_checks):
        time.sleep(EXPORT_DIALOG_POLL_SECONDS)
        hasExcel = driver.execute_script("return document.getElementById('EXCEL2007') !== null;")
        if hasExcel:
            print(f"  ✓ 导出弹窗已加载 (第{i+1}次检查)")
            break
        print(f"  等待弹窗... ({i+1}/{max_checks})")
    else:
        raise TimeoutError("导出弹窗加载超时，未检测到 Excel 选项。")

    print(f"\n步骤{step_base + 1}: 选择Excel")
    result = driver.execute_script(excel_script)
    print(f"  {result}")
    time.sleep(POST_EXCEL_SELECT_WAIT_SECONDS)

    print(f"\n步骤{step_base + 2}: 点击在线导出")
    result = driver.execute_script(online_export_script)
    print(f"  {result}")


def list_download_files():
    if not os.path.exists(DOWNLOAD_DIR):
        return {}

    files = {}
    for file in os.listdir(DOWNLOAD_DIR):
        file_path = os.path.join(DOWNLOAD_DIR, file)
        if not os.path.isfile(file_path):
            continue
        if file.startswith("~$") or file.endswith((".crdownload", ".tmp", ".part")):
            continue
        files[file_path] = {
            "name": file,
            "ctime": os.path.getctime(file_path),
            "size": os.path.getsize(file_path),
        }
    return files


def wait_for_new_download(previous_files):
    print("\n检查下载的文件...")
    deadline = time.time() + DOWNLOAD_READY_TIMEOUT_SECONDS
    stable_counter = 0
    candidate_path = None
    candidate_size = None

    while time.time() < deadline:
        current_files = list_download_files()
        new_paths = [path for path in current_files if path not in previous_files]
        if new_paths:
            newest_path = max(new_paths, key=lambda path: current_files[path]["ctime"])
            newest_size = current_files[newest_path]["size"]
            if newest_path == candidate_path and newest_size == candidate_size:
                stable_counter += 1
            else:
                candidate_path = newest_path
                candidate_size = newest_size
                stable_counter = 1

            if stable_counter >= FILE_STABLE_CHECK_ROUNDS:
                meta = current_files[newest_path]
                print(f"✓ 最新文件: {meta['name']} (创建时间: {datetime.datetime.fromtimestamp(meta['ctime'])})")
                return newest_path
        time.sleep(DOWNLOAD_POLL_SECONDS)

    print("⚠ 未在预期时间内检测到稳定的新下载文件")
    return None


def rename_latest_download(download_path):
    if not download_path:
        return None
    latest_file = os.path.basename(download_path)
    file_size = os.path.getsize(download_path)
    new_filename = get_timestamp_filename(latest_file)
    new_filepath = os.path.join(DOWNLOAD_DIR, new_filename)

    try:
        os.rename(download_path, new_filepath)
        print(f"✓ 重命名: {latest_file} → {new_filename} ({file_size} 字节)")
        print(f"✅ 文件已保存为: {new_filename}")
        return new_filepath
    except Exception as e:
        print(f"⚠ 重命名失败: {latest_file} → {e}")
        return None

def js_export():
    print("使用JavaScript注入的导出脚本")
    username, password, report_ids = load_auth_config()
    conclusion_lines = []
    
    # 创建下载目录
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)
        print(f"✓ 创建下载目录: {DOWNLOAD_DIR}")
    
    edge_options = Options()
    edge_options.add_argument("--start-maximized")
    
    # 设置下载路径
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    edge_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Edge(options=edge_options)
    
    try:
        first_report_id = report_ids[0]
        print(f"步骤1: 打开页面 | report_id={first_report_id}")
        driver.get(build_report_url(first_report_id))
        time.sleep(5)
        
        # 步骤2: 使用JavaScript登录
        print("\n步骤2: 使用JavaScript登录")
        login_script = f"""
            var username = document.querySelector('input[type="text"]');
            var password = document.querySelector('input[type="password"]');
            var loginBtn = document.querySelector('input.item-submit');
            
            if(username && password && loginBtn) {{
                username.value = {json.dumps(username)};
                password.value = {json.dumps(password)};
                loginBtn.click();
                return '登录成功';
            }} else {{
                return '登录元素未找到';
            }}
        """
        result = driver.execute_script(login_script)
        print(f"  {result}")
        
        wait_for_report_ready(driver, "步骤3")
        
        export_script = """
            var exportBtn = document.querySelector('input.btnExport');
            if(!exportBtn) exportBtn = document.querySelector('input[title="导出"]');
            if(!exportBtn) exportBtn = document.querySelector('input[value="导出"]');
            if(!exportBtn) exportBtn = document.querySelector('.queryview-toolbar-button');
            
            if(exportBtn) {
                exportBtn.click();
                console.log('导出按钮已点击');
                return '导出按钮已点击';
            } else {
                // 调试信息
                console.log('导出按钮未找到，页面上的按钮元素:', document.querySelectorAll('input, button'));
                return '导出按钮未找到';
            }
        """
        excel_script = """
            var foundExcel = null;
            var methodUsed = '';
            
            // 方式1: 优先使用ID定位（最可靠）
            var excelById = document.getElementById('EXCEL2007');
            if(excelById) {
                foundExcel = excelById;
                methodUsed = 'ID';
            }
            
            // 方式2: 使用caption属性定位
            if(!foundExcel) {
                var excelByCaption = document.querySelector('[caption="Excel"]');
                if(excelByCaption) {
                    foundExcel = excelByCaption;
                    methodUsed = 'caption属性';
                }
            }
            
            // 方式3: 使用class组合+文本定位
            if(!foundExcel) {
                var excelByClass = document.querySelector('div.dropdown-box-span-select.dropdown-box-span-higher');
                if(excelByClass && excelByClass.textContent && excelByClass.textContent.indexOf('Excel') !== -1) {
                    foundExcel = excelByClass;
                    methodUsed = 'class组合';
                }
            }
            
            // 方式4: 使用onmousedown事件特征定位
            if(!foundExcel) {
                var allDivs = document.querySelectorAll('div[onmousedown*="doItemClick"]');
                for(var i=0; i<allDivs.length; i++) {
                    if(allDivs[i].id && allDivs[i].id.indexOf('EXCEL') !== -1) {
                        foundExcel = allDivs[i];
                        methodUsed = 'onmousedown事件';
                        break;
                    }
                }
            }
            
            if(foundExcel) {
                // 调试信息：检查框架状态
                console.log('=== Excel按钮点击调试信息 ===');
                console.log('Excel按钮元素:', foundExcel);
                console.log('按钮ID:', foundExcel.id);
                console.log('按钮caption:', foundExcel.getAttribute('caption'));
                console.log('onmousedown函数:', typeof foundExcel.onmousedown);
                console.log('onmouseover函数:', typeof foundExcel.onmouseover);
                console.log('getOwner函数是否存在:', typeof getOwner === 'function');
                console.log('jsloader是否存在:', typeof jsloader !== 'undefined');
                if(window.jsloader) {
                    console.log('jsloader.resolve函数:', typeof jsloader.resolve);
                }
                
                var resultMessage = '';
                
                // 方法1: 优先使用框架方法
                try {
                    // 检查getOwner框架
                    if(typeof getOwner === 'function') {
                        var owner = getOwner(foundExcel);
                        if(owner && typeof owner.doItemClick === 'function') {
                            owner.doItemClick(foundExcel);
                            resultMessage = 'Excel已选择 (通过getOwner框架)';
                            console.log('✓ 通过getOwner框架成功');
                        }
                    }
                } catch(e) {
                    console.log('getOwner框架失败:', e.message);
                }
                
                // 方法2: 使用jsloader框架
                if(!resultMessage) {
                    try {
                        if(window.jsloader && typeof jsloader.resolve === 'function') {
                            var util = jsloader.resolve('freequery.common.util');
                            if(util && typeof util.getOwner === 'function') {
                                var owner = util.getOwner(foundExcel);
                                if(owner && typeof owner.doItemClick === 'function') {
                                    owner.doItemClick(foundExcel);
                                    resultMessage = 'Excel已选择 (通过jsloader框架)';
                                    console.log('✓ 通过jsloader框架成功');
                                }
                            }
                        }
                    } catch(e) {
                        console.log('jsloader框架失败:', e.message);
                    }
                }
                
                // 方法3: 触发原生事件
                if(!resultMessage) {
                    try {
                        // 先触发mouseover
                        if(typeof foundExcel.onmouseover === 'function') {
                            foundExcel.onmouseover();
                            console.log('✓ 触发onmouseover');
                        }
                        
                        // 再触发mousedown（主要事件）
                        if(typeof foundExcel.onmousedown === 'function') {
                            foundExcel.onmousedown();
                            resultMessage = 'Excel已选择 (通过onmousedown事件)';
                            console.log('✓ 通过onmousedown事件成功');
                        } else {
                            // 如果onmousedown不存在，直接click
                            foundExcel.click();
                            resultMessage = 'Excel已选择 (通过click事件)';
                            console.log('✓ 通过click事件成功');
                        }
                    } catch(e) {
                        console.log('原生事件失败:', e.message);
                    }
                }
                
                // 方法4: 模拟鼠标事件作为最终兜底
                if(!resultMessage) {
                    try {
                        // 创建完整的鼠标事件序列
                        var mouseoverEvent = new MouseEvent('mouseover', {
                            bubbles: true,
                            cancelable: true,
                            view: window
                        });
                        foundExcel.dispatchEvent(mouseoverEvent);
                        
                        var mousedownEvent = new MouseEvent('mousedown', {
                            bubbles: true,
                            cancelable: true,
                            view: window,
                            button: 0
                        });
                        foundExcel.dispatchEvent(mousedownEvent);
                        
                        var mouseupEvent = new MouseEvent('mouseup', {
                            bubbles: true,
                            cancelable: true,
                            view: window,
                            button: 0
                        });
                        foundExcel.dispatchEvent(mouseupEvent);
                        
                        var clickEvent = new MouseEvent('click', {
                            bubbles: true,
                            cancelable: true,
                            view: window
                        });
                        foundExcel.dispatchEvent(clickEvent);
                        
                        resultMessage = 'Excel已选择 (通过模拟事件)';
                        console.log('✓ 通过模拟事件成功');
                    } catch(e) {
                        console.log('模拟事件失败:', e.message);
                        resultMessage = '所有方法都失败: ' + e.message;
                    }
                }
                
                return resultMessage + ' (通过' + methodUsed + '定位)';
            } else {
                return 'Excel未找到';
            }
        """
        online_export_script = """
            var foundOnline = null;
            
            // 方式1: 使用class和value精确匹配
            var onlineByClass = document.querySelector('input._btnOK.barbtn.btn-default[value="在线导出"]');
            if(onlineByClass) {
                foundOnline = onlineByClass;
                console.log('✓ 通过class和value找到在线导出按钮');
            }
            
            // 方式2: 使用qtp属性定位
            if(!foundOnline) {
                var onlineByQtp = document.querySelector('[qtp="baseDialogEx_btnOK"]');
                if(onlineByQtp) {
                    foundOnline = onlineByQtp;
                    console.log('✓ 通过qtp属性找到在线导出按钮');
                }
            }
            
            // 方式3: 使用class组合定位
            if(!foundOnline) {
                var onlineByClassOnly = document.querySelector('input._btnOK.barbtn.btn-default');
                if(onlineByClassOnly && onlineByClassOnly.value === '在线导出') {
                    foundOnline = onlineByClassOnly;
                    console.log('✓ 通过class组合找到在线导出按钮');
                }
            }
            
            // 方式4: 文本匹配作为兜底
            if(!foundOnline) {
                var allInputs = document.querySelectorAll('input[type="button"]');
                for(var i=0; i<allInputs.length; i++) {
                    if(allInputs[i].value && allInputs[i].value.trim() === '在线导出') {
                        foundOnline = allInputs[i];
                        console.log('✓ 通过文本匹配找到在线导出按钮');
                        break;
                    }
                }
            }
            
            if(foundOnline) {
                // 调试信息
                console.log('在线导出按钮元素:', foundOnline);
                console.log('按钮class:', foundOnline.className);
                console.log('按钮value:', foundOnline.value);
                console.log('按钮qtp:', foundOnline.getAttribute('qtp'));
                
                // 点击按钮
                foundOnline.click();
                return '在线导出已点击';
            } else {
                // 调试信息
                console.log('未找到在线导出按钮，页面上的按钮元素:');
                var allButtons = document.querySelectorAll('input[type="button"], button');
                for(var j=0; j<allButtons.length; j++) {
                    console.log('按钮' + j + ':', allButtons[j].value || allButtons[j].textContent, allButtons[j].className);
                }
                return '在线导出未找到';
            }
        """
        for index, report_rule in enumerate(REPORT_RULES):
            previous_files = list_download_files()
            if index > 0:
                print("\n" + "=" * 50)
                print(f"开始导出第{index + 1}个报表...")
                print("=" * 50)
                report_id = report_ids[index]
                open_step = 7 + (index - 1) * 4
                print(f"\n步骤{open_step}: 打开报表页面 | report_id={report_id}")
                driver.get(build_report_url(report_id))
                wait_for_report_ready(driver, f"步骤{open_step + 1}")
                click_export_controls(driver, export_script, excel_script, online_export_script, open_step + 2)
            else:
                report_id = report_ids[index]
                click_export_controls(driver, export_script, excel_script, online_export_script, 4)

            print(f"\n✅ 第{index + 1}个报表导出完成！")
            download_path = wait_for_new_download(previous_files)
            new_filepath = rename_latest_download(download_path)
            if not new_filepath:
                print("⚠ 未完成文件重命名或未检测到新文件")
                continue

            print("\n开始分析报表数据...")
            conclusion_line = analyze_report_and_build_conclusion(new_filepath, report_rule)
            if conclusion_line:
                conclusion_lines.append(conclusion_line)
                print("✅ 报表分析完成！")
            else:
                print("⚠ 报表分析失败，但文件已保存")

        if conclusion_lines:
            write_summary_workbook(conclusion_lines)
            print(f"\n✅ 周报结论已汇总到: {SUMMARY_OUTPUT_PATH}")
        
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        import traceback
        traceback.print_exc()
    finally:
        time.sleep(10)
        driver.quit()

if __name__ == "__main__":
    js_export()
