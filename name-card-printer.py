import os
import time
import win32com.client as win32
import pandas as pd  # 导入处理Excel的库

# ==========================================
# 1. 系统配置区域
# ==========================================

# Excel 配置
EXCEL_CONFIG = {
    "filename": "名单.xlsx",  # 你的Excel文件名
    "sheet_name": 0,  # 读取第几个Sheet，0表示第一个
    "header": None,  # None表示没有表头(第一行就是名字)，如果第一行是“姓名”这种标题，改为 0
    "column_index": 0  # 读取第几列，0表示A列，1表示B列
}

# 模板配置 (字数 : 配置文件)
TEMPLATE_CONFIG = {
    2: {"file": "2个字.docx", "ph": "模 版"},
    3: {"file": "3个字.docx", "ph": "模板版"},
    4: {"file": "4个字.docx", "ph": "模板模板"},
    5: {"file": "5个字.docx", "ph": "模板模板版"},
    6: {"file": "6个字.docx", "ph": "模板模板模板"},
    7: {"file": "7个字.docx", "ph": "模板模板模板版"},
    # 如果有更多字数，按格式往下面加即可
}


# ==========================================
# 2. 功能函数定义
# ==========================================

def get_names_from_excel(config):
    """读取Excel文件获取名单列表"""
    file_path = config["filename"]
    if not os.path.exists(file_path):
        print(f"❌ 错误：找不到Excel文件 '{file_path}'")
        return []

    try:
        # 读取Excel
        df = pd.read_excel(file_path, sheet_name=config["sheet_name"], header=config["header"])

        # 提取指定列的数据
        # iloc[:, i] 表示取所有行的第i列
        raw_data = df.iloc[:, config["column_index"]]

        # 清洗数据：转为字符串，去除空值(NaN)，去除首尾空格
        name_list = raw_data.dropna().astype(str).str.strip().tolist()

        # 过滤掉可能读取到的表头（如果配置了header=None但实际上有表头，比如读到了'姓名'这个词）
        # 这里做一个简单的过滤，如果名字里包含"姓名"且长度为2，可能需要人工确认，这里简单处理保留

        print(f"📊 成功从 Excel 读取到 {len(name_list)} 个名字。")
        return name_list

    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return []


def word_replace_and_print(word_app, template_path, placeholder, new_name):
    """Word替换并打印核心逻辑"""
    abs_path = os.path.abspath(template_path)
    if not os.path.exists(abs_path):
        print(f"❌ 错误：找不到模板文件 {template_path}")
        return False

    try:
        doc = word_app.Documents.Open(abs_path)
        wdReplaceAll = 2

        # 遍历所有StoryRanges (包含文本框、正文等)
        for story in doc.StoryRanges:
            find_obj = story.Find
            find_obj.Text = placeholder
            find_obj.Replacement.Text = new_name
            find_obj.Execute(Replace=wdReplaceAll)

            while story.NextStoryRange:
                story = story.NextStoryRange
                find_obj = story.Find
                find_obj.Text = placeholder
                find_obj.Replacement.Text = new_name
                find_obj.Execute(Replace=wdReplaceAll)

        print(f"🖨️  正在发送打印任务: {new_name}")
        doc.PrintOut()
        time.sleep(2)  # 缓冲时间
        doc.Close(SaveChanges=False)
        return True
    except Exception as e:
        print(f"❌ 打印处理错误: {e}")
        try:
            doc.Close(SaveChanges=False)
        except:
            pass
        return False


# ==========================================
# 3. 主程序逻辑
# ==========================================
if __name__ == "__main__":

    print("--- 自动化席卡打印系统 (Excel版) ---")

    # 1. 从 Excel 获取名单
    raw_name_list = get_names_from_excel(EXCEL_CONFIG)

    if not raw_name_list:
        print("程序终止：名单为空或读取失败。")
        exit()

    # 2. 自动匹配模板
    all_jobs = []
    print("\n正在匹配模板...")

    for name in raw_name_list:
        # 去除名字内部的所有空格来计算真实字数 (如 "陈 伟" -> 2字)
        clean_name = name.replace(" ", "").replace("　", "")
        name_len = len(clean_name)

        # 查找配置
        config = TEMPLATE_CONFIG.get(name_len)

        if config:
            all_jobs.append({
                'name': name,  # 打印原本的内容（Excel里是啥就是啥）
                'clean_name': clean_name,  # 用于显示的干净名字
                'len': name_len,
                'tpl': config['file'],
                'ph': config['ph']
            })
        else:
            print(f"⚠️  跳过: '{name}' (长度{name_len}字，未配置对应模板)")

    if not all_jobs:
        print("❌ 没有有效的打印任务。")
        exit()

    print(f"✅ 生成 {len(all_jobs)} 个打印任务。")

    # 3. 启动 Word
    print("正在启动 Word...")
    word = win32.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = False

    try:
        # --- 试打环节 ---
        first_job = all_jobs[0]
        print("\n===================================")
        print(f"🧪 试打第1位：{first_job['name']}")
        print(f"   匹配模板：{first_job['tpl']}")
        print("===================================")

        success = word_replace_and_print(word, first_job['tpl'], first_job['ph'], first_job['name'])

        if not success:
            print("❌ 试打失败，程序退出。")
            word.Quit()
            exit()

        # --- 确认环节 ---
        print("\n" + "=" * 50)
        print("请检查打印机输出结果。")
        print("=" * 50)
        user_input = input(">>> 确认无误继续打印剩余名单？(输入 y 继续，其他键退出): ")

        if user_input.lower() == 'y':
            print("\n🚀 开始批量打印剩余名单...")

            remaining_jobs = all_jobs[1:]

            for index, job in enumerate(remaining_jobs):
                print(f"[{index + 1}/{len(remaining_jobs)}] ", end="")
                word_replace_and_print(word, job['tpl'], job['ph'], job['name'])

            print("\n✅ 所有任务已完成！")
        else:
            print("\n🛑 已取消打印。")

    except Exception as e:
        print(f"\n❌ 发生未知错误: {e}")
    finally:
        print("退出 Word。")
        word.Quit()