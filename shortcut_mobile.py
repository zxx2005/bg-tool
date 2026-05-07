import streamlit as st
import random
import json
import os

# ---------------------- 页面基础设置 ----------------------
st.set_page_config(
    page_title="办公快捷键大师",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ---------------------- 初始化收藏错题文件 ----------------------
data_file = "shortcut_mobile.json"
if not os.path.exists(data_file):
    data = {"favorites": [], "errors": []}
    with open(data_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
else:
    with open(data_file, "r", encoding="utf-8") as f:
        data = json.load(f)

# ---------------------- 专属快捷键库（Excel + Word + WPS） ----------------------
shortcuts = [
    # ---------------------- Excel 表格（职场核心） ----------------------
{"key":"Ctrl + S","desc":"快速保存工作簿，防止数据丢失","cate":"Excel"},
{"key":"Ctrl + C","desc":"复制选中单元格内容","cate":"Excel"},
{"key":"Ctrl + V","desc":"粘贴复制内容","cate":"Excel"},
{"key":"Ctrl + X","desc":"剪切选中内容","cate":"Excel"},
{"key":"Ctrl + Z","desc":"撤销上一步操作","cate":"Excel"},
{"key":"Ctrl + Y","desc":"恢复撤销操作","cate":"Excel"},
{"key":"Ctrl + A","desc":"全选整个工作表","cate":"Excel"},
{"key":"Ctrl + F","desc":"查找内容","cate":"Excel"},
{"key":"Ctrl + H","desc":"查找并批量替换","cate":"Excel"},
{"key":"Ctrl + N","desc":"新建空白工作簿","cate":"Excel"},
{"key":"Ctrl + O","desc":"打开Excel文件","cate":"Excel"},
{"key":"Ctrl + P","desc":"调出打印预览界面","cate":"Excel"},
{"key":"Ctrl + W","desc":"关闭当前表格","cate":"Excel"},
{"key":"Ctrl + F4","desc":"关闭Excel窗口","cate":"Excel"},
{"key":"F12","desc":"另存为文件","cate":"Excel"},
{"key":"Ctrl + B","desc":"文字加粗","cate":"Excel"},
{"key":"Ctrl + I","desc":"文字斜体","cate":"Excel"},
{"key":"Ctrl + U","desc":"添加下划线","cate":"Excel"},
{"key":"Ctrl + 1","desc":"打开单元格格式窗口","cate":"Excel"},
{"key":"Ctrl + Shift + $","desc":"转为货币格式","cate":"Excel"},
{"key":"Ctrl + Shift + %","desc":"转为百分比格式","cate":"Excel"},
{"key":"Ctrl + Shift + #","desc":"转为日期格式","cate":"Excel"},
{"key":"Ctrl + Shift + @","desc":"转为时间格式","cate":"Excel"},
{"key":"Ctrl + Shift + !","desc":"转为常规数字格式","cate":"Excel"},
{"key":"Ctrl + Shift + &","desc":"添加外边框","cate":"Excel"},
{"key":"Ctrl + Shift + _","desc":"清除所有边框","cate":"Excel"},
{"key":"Ctrl + ;","desc":"插入系统当前日期","cate":"Excel"},
{"key":"Ctrl + Shift + ;","desc":"插入系统当前时间","cate":"Excel"},
{"key":"Ctrl + D","desc":"向下填充内容","cate":"Excel"},
{"key":"Ctrl + R","desc":"向右填充内容","cate":"Excel"},
{"key":"Ctrl + +","desc":"插入单元格/行/列","cate":"Excel"},
{"key":"Ctrl + -","desc":"删除单元格/行/列","cate":"Excel"},
{"key":"Ctrl + PageUp","desc":"切换上一个工作表","cate":"Excel"},
{"key":"Ctrl + PageDown","desc":"切换下一个工作表","cate":"Excel"},
{"key":"Alt + =","desc":"一键自动求和","cate":"Excel"},
{"key":"F4","desc":"重复上一步任何操作","cate":"Excel"},
{"key":"F2","desc":"编辑单元格内容","cate":"Excel"},
{"key":"F5","desc":"定位单元格位置","cate":"Excel"},
{"key":"F11","desc":"快速生成图表","cate":"Excel"},
{"key":"Shift + F11","desc":"新建空白工作表","cate":"Excel"},
{"key":"Ctrl + Shift + L","desc":"开启或关闭筛选功能","cate":"Excel"},
{"key":"Ctrl + Shift + O","desc":"选中所有批注单元格","cate":"Excel"},
{"key":"Ctrl + Shift + U","desc":"展开或折叠编辑栏","cate":"Excel"},
{"key":"Ctrl + Alt + V","desc":"选择性粘贴高级窗口","cate":"Excel"},
{"key":"Ctrl + Shift + Space","desc":"全选当前表格区域","cate":"Excel"},
{"key":"Shift + Space","desc":"选中整行","cate":"Excel"},
{"key":"Ctrl + Space","desc":"选中整列","cate":"Excel"},
{"key":"Ctrl + 5","desc":"添加删除线","cate":"Excel"},
{"key":"Ctrl + 9","desc":"隐藏选中行","cate":"Excel"},
{"key":"Ctrl + 0","desc":"隐藏选中列","cate":"Excel"},
{"key":"Ctrl + Shift + 9","desc":"取消隐藏行","cate":"Excel"},
{"key":"Ctrl + Shift + 0","desc":"取消隐藏列","cate":"Excel"},
{"key":"Alt + Enter","desc":"单元格内自动换行","cate":"Excel"},
{"key":"Ctrl + `","desc":"公式模式和结果模式切换","cate":"Excel"},
{"key":"Ctrl + ←","desc":"跳到数据最左侧","cate":"Excel"},
{"key":"Ctrl + →","desc":"跳到数据最右侧","cate":"Excel"},
{"key":"Ctrl + ↑","desc":"跳到数据最上方","cate":"Excel"},
{"key":"Ctrl + ↓","desc":"跳到数据最下方","cate":"Excel"},
{"key":"Shift + 方向键","desc":"逐步扩大选区","cate":"Excel"},
{"key":"Tab","desc":"向右切换单元格","cate":"Excel"},
{"key":"Shift + Tab","desc":"向左切换单元格","cate":"Excel"},
{"key":"Enter","desc":"向下移动单元格","cate":"Excel"},
{"key":"Shift + Enter","desc":"向上移动单元格","cate":"Excel"},
{"key":"Esc","desc":"取消当前编辑","cate":"Excel"},
{"key":"Ctrl + Alt + +","desc":"自定义快捷键设置","cate":"Excel"},
{"key":"Alt + F1","desc":"嵌入图表到工作表","cate":"Excel"},
{"key":"Alt + F11","desc":"打开VBA编辑器","cate":"Excel"},
{"key":"Ctrl + F3","desc":"打开名称管理器","cate":"Excel"},
{"key":"Ctrl + Shift + F3","desc":"根据所选内容创建名称","cate":"Excel"},
{"key":"F7","desc":"拼写检查","cate":"Excel"},
{"key":"F8","desc":"进入扩展选取模式","cate":"Excel"},
{"key":"Shift + F8","desc":"添加选区模式","cate":"Excel"},
{"key":"Ctrl + F8","desc":"调整窗口大小","cate":"Excel"},
{"key":"Alt + F8","desc":"运行宏","cate":"Excel"},
{"key":"Ctrl + Shift + =","desc":"快速插入单元格","cate":"Excel"},
{"key":"Ctrl + .","desc":"旋转单元格格式","cate":"Excel"},
{"key":"Ctrl + Alt + →","desc":"向右移动页面","cate":"Excel"},
{"key":"Ctrl + Alt + ←","desc":"向左移动页面","cate":"Excel"},
{"key":"Shift + Alt + →","desc":"分组字段折叠","cate":"Excel"},
{"key":"Shift + Alt + ←","desc":"取消分组折叠","cate":"Excel"},
{"key":"Ctrl + Shift + ;","desc":"录入静态当前时间","cate":"Excel"},
{"key":"Ctrl + Alt + F5","desc":"刷新外部数据","cate":"Excel"},
{"key":"Ctrl + End","desc":"跳到表格最后一个单元格","cate":"Excel"},
{"key":"Ctrl + Home","desc":"跳到A1单元格","cate":"Excel"},
{"key":"Ctrl + Shift + Home","desc":"选中从光标到开头区域","cate":"Excel"},
{"key":"Ctrl + Shift + End","desc":"选中从光标到末尾区域","cate":"Excel"},
{"key":"PageUp","desc":"页面向上翻页","cate":"Excel"},
{"key":"PageDown","desc":"页面向下翻页","cate":"Excel"},
{"key":"Alt + PageUp","desc":"向左翻页","cate":"Excel"},
{"key":"Alt + PageDown","desc":"向右翻页","cate":"Excel"},
{"key":"Ctrl + PageUp","desc":"上个工作表","cate":"Excel"},
{"key":"Ctrl + PageDown","desc":"下个工作表","cate":"Excel"},
{"key":"Ctrl + Shift + PageUp","desc":"选中上一个工作表","cate":"Excel"},
{"key":"Ctrl + Shift + PageDown","desc":"选中下一个工作表","cate":"Excel"},
{"key":"Ctrl + Alt + Del","desc":"打开任务管理界面","cate":"Excel"},
{"key":"Shift + F5","desc":"定位条件选择","cate":"Excel"},
{"key":"Ctrl + F9","desc":"最小化工作簿","cate":"Excel"},
{"key":"Ctrl + F10","desc":"还原工作簿窗口","cate":"Excel"},
{"key":"Ctrl + F6","desc":"切换窗口","cate":"Excel"},

    # ---------------------- Word 文字（文档排版核心） ----------------------
{"key":"Ctrl + S","desc":"保存Word文档","cate":"Word"},
{"key":"Ctrl + C","desc":"复制文字内容","cate":"Word"},
{"key":"Ctrl + V","desc":"粘贴文字内容","cate":"Word"},
{"key":"Ctrl + X","desc":"剪切选中文字","cate":"Word"},
{"key":"Ctrl + Z","desc":"撤销操作","cate":"Word"},
{"key":"Ctrl + Y","desc":"恢复操作","cate":"Word"},
{"key":"Ctrl + A","desc":"全选整篇文档","cate":"Word"},
{"key":"Ctrl + F","desc":"查找文字","cate":"Word"},
{"key":"Ctrl + H","desc":"查找替换文字","cate":"Word"},
{"key":"Ctrl + N","desc":"新建空白文档","cate":"Word"},
{"key":"Ctrl + O","desc":"打开已有文档","cate":"Word"},
{"key":"Ctrl + P","desc":"打印文档","cate":"Word"},
{"key":"Ctrl + W","desc":"关闭当前文档","cate":"Word"},
{"key":"Ctrl + B","desc":"字体加粗","cate":"Word"},
{"key":"Ctrl + I","desc":"字体斜体","cate":"Word"},
{"key":"Ctrl + U","desc":"添加下划线","cate":"Word"},
{"key":"Ctrl + Shift + >","desc":"整体放大字号","cate":"Word"},
{"key":"Ctrl + Shift + <","desc":"整体缩小字号","cate":"Word"},
{"key":"Ctrl + ]","desc":"逐号增大字体","cate":"Word"},
{"key":"Ctrl + [","desc":"逐号缩小字体","cate":"Word"},
{"key":"Ctrl + L","desc":"文字左对齐","cate":"Word"},
{"key":"Ctrl + E","desc":"文字居中对齐","cate":"Word"},
{"key":"Ctrl + R","desc":"文字右对齐","cate":"Word"},
{"key":"Ctrl + J","desc":"两端自动对齐","cate":"Word"},
{"key":"Ctrl + 1","desc":"单倍行距","cate":"Word"},
{"key":"Ctrl + 2","desc":"双倍行距","cate":"Word"},
{"key":"Ctrl + 5","desc":"1.5倍行距","cate":"Word"},
{"key":"Ctrl + 0","desc":"取消段前间距","cate":"Word"},
{"key":"Ctrl + K","desc":"插入超链接","cate":"Word"},
{"key":"Ctrl + Enter","desc":"强制分页符","cate":"Word"},
{"key":"Ctrl + Shift + Enter","desc":"插入分节符","cate":"Word"},
{"key":"Ctrl + Home","desc":"跳到文档最开头","cate":"Word"},
{"key":"Ctrl + End","desc":"跳到文档最末尾","cate":"Word"},
{"key":"Shift + F3","desc":"英文大小写切换","cate":"Word"},
{"key":"Ctrl + Shift + C","desc":"复制文字格式","cate":"Word"},
{"key":"Ctrl + Shift + V","desc":"粘贴文字格式","cate":"Word"},
{"key":"Ctrl + Shift + L","desc":"添加项目符号列表","cate":"Word"},
{"key":"Ctrl + *","desc":"显示隐藏编辑标记","cate":"Word"},
{"key":"Alt + F7","desc":"拼写语法检查","cate":"Word"},
{"key":"Ctrl + Shift + S","desc":"打开样式面板","cate":"Word"},
{"key":"Ctrl + Alt + I","desc":"打印预览视图","cate":"Word"},
{"key":"Ctrl + Alt + O","desc":"大纲视图模式","cate":"Word"},
{"key":"Ctrl + Alt + N","desc":"普通页面视图","cate":"Word"},
{"key":"Alt + Shift + →","desc":"提升标题层级","cate":"Word"},
{"key":"Alt + Shift + ←","desc":"降低标题层级","cate":"Word"},
{"key":"Ctrl + Shift + 1","desc":"应用标题1样式","cate":"Word"},
{"key":"Ctrl + Shift + 2","desc":"应用标题2样式","cate":"Word"},
{"key":"Ctrl + Shift + 3","desc":"应用标题3样式","cate":"Word"},
{"key":"Ctrl + Shift + N","desc":"正文标准样式","cate":"Word"},
{"key":"Ctrl + Delete","desc":"删除光标右侧整个词语","cate":"Word"},
{"key":"Ctrl + Backspace","desc":"删除光标左侧整个词语","cate":"Word"},
{"key":"Ctrl + ←","desc":"光标跳到上个词语","cate":"Word"},
{"key":"Ctrl + →","desc":"光标跳到下个词语","cate":"Word"},
{"key":"Shift + ←","desc":"向左选文字","cate":"Word"},
{"key":"Shift + →","desc":"向右选文字","cate":"Word"},
{"key":"Shift + ↑","desc":"向上选文字","cate":"Word"},
{"key":"Shift + ↓","desc":"向下选文字","cate":"Word"},
{"key":"Ctrl + Shift + ←","desc":"向左快速选词","cate":"Word"},
{"key":"Ctrl + Shift + →","desc":"向右快速选词","cate":"Word"},
{"key":"Enter","desc":"换行分段","cate":"Word"},
{"key":"Shift + Enter","desc":"软换行不产生段落","cate":"Word"},
{"key":"Esc","desc":"取消当前操作","cate":"Word"},
{"key":"Tab","desc":"向右缩进","cate":"Word"},
{"key":"Shift + Tab","desc":"向左缩进","cate":"Word"},
{"key":"Ctrl + Tab","desc":"插入制表符","cate":"Word"},
{"key":"Ctrl + F3","desc":"剪切词条存入图文集","cate":"Word"},
{"key":"Ctrl + Shift + F3","desc":"粘贴图文集内容","cate":"Word"},
{"key":"F7","desc":"拼写检查","cate":"Word"},
{"key":"F12","desc":"文档另存为","cate":"Word"},
{"key":"Ctrl + F12","desc":"打开文档","cate":"Word"},
{"key":"Alt + F4","desc":"关闭整个Word软件","cate":"Word"},
{"key":"Ctrl + F4","desc":"关闭当前文档","cate":"Word"},
{"key":"Ctrl + Shift + F9","desc":"取消域代码","cate":"Word"},
{"key":"Ctrl + Alt + C","desc":"插入版权符号","cate":"Word"},
{"key":"Ctrl + Alt + R","desc":"插入注册商标符号","cate":"Word"},
{"key":"Ctrl + Alt + T","desc":"插入商标符号","cate":"Word"},
{"key":"Ctrl + -","desc":"可选连字符","cate":"Word"},
{"key":"Ctrl + Shift + -","desc":"长破折号","cate":"Word"},
{"key":"Ctrl + Shift + =","desc":"上标文字","cate":"Word"},
{"key":"Ctrl + =","desc":"下标文字","cate":"Word"},
{"key":"Ctrl + .","desc":"英文全角半角切换","cate":"Word"},
{"key":"Ctrl + /","desc":"快速打开帮助搜索","cate":"Word"},
{"key":"Alt + Shift + D","desc":"插入当前日期","cate":"Word"},
{"key":"Alt + Shift + T","desc":"插入当前时间","cate":"Word"},
{"key":"Ctrl + Alt + M","desc":"插入批注","cate":"Word"},
{"key":"Ctrl + Shift + E","desc":"开启修订模式","cate":"Word"},
{"key":"Ctrl + Shift + F","desc":"字体设置窗口","cate":"Word"},
{"key":"Ctrl + Shift + P","desc":"字号设置窗口","cate":"Word"},
{"key":"Ctrl + Alt + Plus","desc":"自定义快捷键","cate":"Word"},
{"key":"PageUp","desc":"向上翻页","cate":"Word"},
{"key":"PageDown","desc":"向下翻页","cate":"Word"},
{"key":"Ctrl + PageUp","desc":"跳到上一页顶部","cate":"Word"},
{"key":"Ctrl + PageDown","desc":"跳到下一页顶部","cate":"Word"},
{"key":"Shift + PageUp","desc":"向上整页选中","cate":"Word"},
{"key":"Shift + PageDown","desc":"向下整页选中","cate":"Word"},
{"key":"Ctrl + Shift + Home","desc":"选中光标到文档开头","cate":"Word"},
{"key":"Ctrl + Shift + End","desc":"选中光标到文档结尾","cate":"Word"},
    # ---------------------- WPS 通用（兼容Word/Excel/PPT） ----------------------
{"key":"Ctrl + S","desc":"通用保存所有文件","cate":"WPS"},
{"key":"Ctrl + C","desc":"通用复制","cate":"WPS"},
{"key":"Ctrl + V","desc":"通用粘贴","cate":"WPS"},
{"key":"Ctrl + X","desc":"通用剪切","cate":"WPS"},
{"key":"Ctrl + Z","desc":"通用撤销","cate":"WPS"},
{"key":"Ctrl + Y","desc":"通用恢复","cate":"WPS"},
{"key":"Ctrl + A","desc":"通用全选","cate":"WPS"},
{"key":"Ctrl + F","desc":"通用查找","cate":"WPS"},
{"key":"Ctrl + H","desc":"通用替换","cate":"WPS"},
{"key":"Ctrl + N","desc":"新建任意文档类型","cate":"WPS"},
{"key":"Ctrl + O","desc":"打开文件","cate":"WPS"},
{"key":"Ctrl + P","desc":"打印文件","cate":"WPS"},
{"key":"Ctrl + W","desc":"关闭当前标签页","cate":"WPS"},
{"key":"Ctrl + Tab","desc":"向右切换文档标签","cate":"WPS"},
{"key":"Ctrl + Shift + Tab","desc":"向左切换文档标签","cate":"WPS"},
{"key":"F12","desc":"另存为文件","cate":"WPS"},
{"key":"Ctrl + F1","desc":"隐藏或显示顶部功能栏","cate":"WPS"},
{"key":"Ctrl + Shift + F12","desc":"直接快速打印","cate":"WPS"},
{"key":"Alt + F4","desc":"彻底退出WPS软件","cate":"WPS"},
{"key":"Ctrl + Shift + M","desc":"插入批注","cate":"WPS"},
{"key":"Ctrl + Shift + C","desc":"通用复制格式","cate":"WPS"},
{"key":"Ctrl + Shift + V","desc":"通用粘贴格式","cate":"WPS"},
{"key":"Ctrl + Shift + S","desc":"打开样式设置面板","cate":"WPS"},
{"key":"Ctrl + Shift + L","desc":"添加项目符号列表","cate":"WPS"},
{"key":"Ctrl + Shift + D","desc":"添加删除线","cate":"WPS"},
{"key":"Ctrl + Shift + >","desc":"放大字体","cate":"WPS"},
{"key":"Ctrl + Shift + <","desc":"缩小字体","cate":"WPS"},
{"key":"Ctrl + E","desc":"居中对齐","cate":"WPS"},
{"key":"Ctrl + L","desc":"左对齐","cate":"WPS"},
{"key":"Ctrl + R","desc":"右对齐","cate":"WPS"},
{"key":"Ctrl + J","desc":"两端对齐","cate":"WPS"},
{"key":"Ctrl + 1","desc":"单倍行距","cate":"WPS"},
{"key":"Ctrl + 2","desc":"双倍行距","cate":"WPS"},
{"key":"Ctrl + 5","desc":"1.5倍行距","cate":"WPS"},
{"key":"Ctrl + Home","desc":"跳到文档开头","cate":"WPS"},
{"key":"Ctrl + End","desc":"跳到文档结尾","cate":"WPS"},
{"key":"Shift + F3","desc":"英文大小写切换","cate":"WPS"},
{"key":"Alt + Shift + →","desc":"升级列表层级","cate":"WPS"},
{"key":"Alt + Shift + ←","desc":"降级列表层级","cate":"WPS"},
{"key":"Alt + Enter","desc":"查看文件属性","cate":"WPS"},
{"key":"Ctrl + Shift + Delete","desc":"永久删除文件不进回收站","cate":"WPS"},
{"key":"Alt + Space + N","desc":"窗口最小化","cate":"WPS"},
{"key":"Alt + Space + X","desc":"窗口最大化","cate":"WPS"},
{"key":"Alt + Space + R","desc":"窗口还原大小","cate":"WPS"},
{"key":"Ctrl + K","desc":"插入超链接","cate":"WPS"},
{"key":"Ctrl + Alt + Z","desc":"多级撤销","cate":"WPS"},
{"key":"Ctrl + Shift + Z","desc":"多级恢复","cate":"WPS"},
{"key":"Alt + F8","desc":"运行宏功能","cate":"WPS"},
{"key":"Ctrl + Alt + I","desc":"打印预览","cate":"WPS"},
{"key":"Ctrl + Alt + O","desc":"大纲视图","cate":"WPS"},
{"key":"Ctrl + Alt + N","desc":"普通视图","cate":"WPS"},
{"key":"Ctrl + B","desc":"字体加粗","cate":"WPS"},
{"key":"Ctrl + I","desc":"字体斜体","cate":"WPS"},
{"key":"Ctrl + U","desc":"字体下划线","cate":"WPS"},
{"key":"Ctrl + Shift + U","desc":"加粗下划线","cate":"WPS"},
{"key":"Ctrl + Shift + B","desc":"快速加粗","cate":"WPS"},
{"key":"Ctrl + Shift + I","desc":"快速斜体","cate":"WPS"},
{"key":"Ctrl + Enter","desc":"分页符","cate":"WPS"},
{"key":"Ctrl + Shift + Enter","desc":"分节符","cate":"WPS"},
{"key":"Ctrl + ←","desc":"光标跳到上个词","cate":"WPS"},
{"key":"Ctrl + →","desc":"光标跳到下个词","cate":"WPS"},
{"key":"Ctrl + Shift + ←","desc":"向左快速选中词语","cate":"WPS"},
{"key":"Ctrl + Shift + →","desc":"向右快速选中词语","cate":"WPS"},
{"key":"Ctrl + Delete","desc":"删除右侧整词","cate":"WPS"},
{"key":"Ctrl + Backspace","desc":"删除左侧整词","cate":"WPS"},
{"key":"Tab","desc":"向右缩进","cate":"WPS"},
{"key":"Shift + Tab","desc":"向左缩进","cate":"WPS"},
{"key":"Ctrl + Tab","desc":"插入制表符空格","cate":"WPS"},
{"key":"Esc","desc":"取消弹窗和操作","cate":"WPS"},
{"key":"Enter","desc":"换行确认","cate":"WPS"},
{"key":"Shift + Enter","desc":"软换行","cate":"WPS"},
{"key":"F7","desc":"拼写检查","cate":"WPS"},
{"key":"Ctrl + F4","desc":"关闭当前页面","cate":"WPS"},
{"key":"Ctrl + F6","desc":"切换窗口","cate":"WPS"},
{"key":"Ctrl + Shift + F6","desc":"反向切换窗口","cate":"WPS"},
{"key":"PageUp","desc":"向上翻页","cate":"WPS"},
{"key":"PageDown","desc":"向下翻页","cate":"WPS"},
{"key":"Ctrl + PageUp","desc":"上一页","cate":"WPS"},
{"key":"Ctrl + PageDown","desc":"下一页","cate":"WPS"},
{"key":"Shift + PageUp","desc":"向上选页","cate":"WPS"},
{"key":"Shift + PageDown","desc":"向下选页","cate":"WPS"},
{"key":"Ctrl + Alt + M","desc":"新建批注","cate":"WPS"},
{"key":"Ctrl + Shift + E","desc":"开启修订模式","cate":"WPS"},
{"key":"Ctrl + =","desc":"下标字体","cate":"WPS"},
{"key":"Ctrl + Shift + =","desc":"上标字体","cate":"WPS"},
{"key":"Ctrl + -","desc":"短横线符号","cate":"WPS"},
{"key":"Ctrl + Shift + -","desc":"长横线符号","cate":"WPS"},
{"key":"Alt + Shift + D","desc":"插入日期","cate":"WPS"},
{"key":"Alt + Shift + T","desc":"插入时间","cate":"WPS"},
{"key":"Ctrl + *","desc":"显示编辑标记","cate":"WPS"},
{"key":"Ctrl + .","desc":"中英文标点切换","cate":"WPS"},
{"key":"Ctrl + /","desc":"帮助搜索","cate":"WPS"},
{"key":"Ctrl + Alt + Del","desc":"系统任务管理器","cate":"WPS"},
{"key":"Ctrl + Shift + Esc","desc":"直接打开任务管理器","cate":"WPS"},
{"key":"Ctrl + Alt + Plus","desc":"自定义软件快捷键","cate":"WPS"},   
]

# ---------------------- 三软件标签页切换 ----------------------
tab_excel, tab_word, tab_wps = st.tabs(["📊 Excel", "📄 Word", "📎 WPS 通用"])

# 为每个标签页准备快捷键数据
excel_shortcuts = [x for x in shortcuts if x["cate"] == "Excel"]
word_shortcuts = [x for x in shortcuts if x["cate"] == "Word"]
wps_shortcuts = [x for x in shortcuts if x["cate"] == "WPS"]

# 顺序学习版：按列表顺序一个一个往下刷
def show_shortcut_module(shortcut_list, category_name):
    # 初始化索引，解决 KeyError 问题
    if f"idx_{category_name}" not in st.session_state:
        st.session_state[f"idx_{category_name}"] = 0

    idx = st.session_state[f"idx_{category_name}"]
    current = shortcut_list[idx]

    # 快捷键信息卡片
    with st.container():
        st.markdown(f"### 🔑 {current['key']}")
        st.info(f"💡 {current['desc']}")
        st.caption(f"📁 分类：{current['cate']}")

    # 按钮顺序：下一个 → 收藏 → 上一个 → 标记错题
    if st.button("🔄 下一个", key=f"next_{category_name}", use_container_width=True):
        if idx < len(shortcut_list) - 1:
            st.session_state[f"idx_{category_name}"] += 1
            st.rerun()
        else:
            st.info("已经是最后一个了")

    if st.button("⭐ 收藏", key=f"fav_{category_name}", use_container_width=True):
        if current["key"] not in data["favorites"]:
            data["favorites"].append(current["key"])
            with open(data_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            st.success("收藏成功！")
        else:
            st.info("已经收藏过了")
        st.rerun()

    if st.button("⏮️ 上一个", key=f"prev_{category_name}", use_container_width=True):
        if idx > 0:
            st.session_state[f"idx_{category_name}"] -= 1
            st.rerun()
        else:
            st.info("已经是第一个了")

    if st.button("❌ 标记错题", key=f"err_{category_name}", use_container_width=True):
        if current["key"] not in data["errors"]:
            data["errors"].append(current["key"])
            with open(data_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            st.warning("已标记为错题！")
        else:
            st.info("已经标记过了")
        st.rerun()
    
# ---------------------- 主界面 ----------------------
st.title("⌨ 办公快捷键大师")
st.divider()

# 三个标签页分别渲染
with tab_excel:
    show_shortcut_module(excel_shortcuts, "Excel")

with tab_word:
    show_shortcut_module(word_shortcuts, "Word")

with tab_wps:
    show_shortcut_module(wps_shortcuts, "WPS")

# 收藏错题查看面板（修复版：解决重复快捷键分类错乱问题）
with st.expander("📂 查看我的收藏 & 错题"):
    # 加到你的「查看我的收藏 & 错题」模块最前面
    st.info("💡 小提示：部分通用快捷键（如Ctrl+C）在不同软件中是一样的，收藏/错题本会按快捷键文本匹配，可能会显示为当前页面的分类，不影响正常学习和使用哦~")
    st.subheader("⭐ 我的收藏")
    if data["favorites"]:
        for fav in data["favorites"]:
            # 遍历所有快捷键，找到和收藏key匹配的条目
            matched = None
            for item in shortcuts:
                if item["key"] == fav:
                    matched = item
                    break
            if matched:
                with st.container():
                    st.markdown(f"**🔑 {matched['key']}**")
                    st.caption(f"📁 分类：{matched['cate']}")
                    st.info(f"💡 {matched['desc']}")
                    if st.button(f"🗑️ 删除收藏", key=f"del_fav_{fav}", use_container_width=True):
                        data["favorites"].remove(fav)
                        with open(data_file, "w", encoding="utf-8") as f:
                            json.dump(data, f, ensure_ascii=False, indent=4)
                        st.rerun()
                    st.divider()
            else:
                st.markdown(f"**🔑 {fav}**")
                st.caption(f"📁 分类：未知")
                st.info(f"💡 无说明")
    else:
        st.write("暂无收藏")

    st.subheader("❌ 我的错题")
    if data["errors"]:
        for err in data["errors"]:
            matched = None
            for item in shortcuts:
                if item["key"] == err:
                    matched = item
                    break
            if matched:
                with st.container():
                    st.markdown(f"**🔑 {matched['key']}**")
                    st.caption(f"📁 分类：{matched['cate']}")
                    st.warning(f"💡 {matched['desc']}")
                    if st.button(f"🗑️ 移除错题", key=f"del_err_{err}", use_container_width=True):
                        data["errors"].remove(err)
                        with open(data_file, "w", encoding="utf-8") as f:
                            json.dump(data, f, ensure_ascii=False, indent=4)
                        st.rerun()
                    st.divider()
            else:
                st.markdown(f"**🔑 {err}**")
                st.caption(f"📁 分类：未知")
                st.warning(f"💡 无说明")
    else:
        st.write("暂无错题")