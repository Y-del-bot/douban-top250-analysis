# -*- coding: utf-8 -*-
"""
豆瓣电影Top250 数据分析工具

爬取豆瓣电影 Top250 榜单数据，自动生成：
- Excel 数据表（豆瓣电影Top250.xlsx）
- 评分分布柱状图（评分分布.png）
- 上映年份分布图（年份分布.png）
- 电影简介词云（简介词云.png）
- 综合静态分析报告（数据分析报告.html）
"""
import os
import re
import time

import matplotlib.pyplot as plt
import pandas as pd
import requests
from bs4 import BeautifulSoup
from wordcloud import WordCloud

# 请求头，防止豆瓣拦截
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0.0.0 Safari/537.36"
}

# 程序生成的所有产物文件
OUTPUT_FILES = [
    "豆瓣电影Top250.xlsx",
    "评分分布.png",
    "年份分布.png",
    "简介词云.png",
    "数据分析报告.html",
]


def clean_old_files():
    """清理上次运行生成的旧文件，避免残留。"""
    for f in OUTPUT_FILES:
        if os.path.exists(f):
            os.remove(f)
    print("已自动清理旧文件")


def show_menu():
    """打印功能菜单并返回用户选择。"""
    print("=" * 45)
    print("豆瓣电影Top250 数据分析工具")
    print("=" * 45)
    print("可选执行功能：")
    print("1 = 仅爬取电影列表，导出Excel")
    print("2 = 爬取数据 + 评分分布柱状图")
    print("3 = 爬取数据 + 年份统计 + 简介词云 + 分析报告")
    print("4 = 一键执行全部完整流程（推荐演示用）")
    return input("请输入数字1/2/3/4：")


def crawl_top250():
    """爬取豆瓣 Top250 全部 10 页数据，返回 DataFrame。"""
    movie_list = []
    print("\n开始抓取豆瓣Top250数据，请稍候……")
    for page in range(10):
        print(f"正在抓取第 {page + 1} 页")
        url = f"https://movie.douban.com/top250?start={page * 25}"
        resp = requests.get(url, headers=HEADERS)
        soup = BeautifulSoup(resp.text, "html.parser")

        for item in soup.find_all("div", class_="item"):
            rank = item.find("em").text
            film_name = item.find("span", class_="title").text
            score = item.find("span", class_="rating_num").text
            brief_tag = item.find("span", class_="inq")
            brief = brief_tag.text if brief_tag else ""
            detail_info = item.find("div", class_="bd").p.text.strip()

            movie_list.append({
                "排名": rank,
                "电影名": film_name,
                "评分": score,
                "一句话简介": brief,
                "导演年份信息": detail_info,
            })
        time.sleep(1.1)  # 控制频率，避免被封

    df = pd.DataFrame(movie_list)
    df["评分"] = df["评分"].astype(float)
    return df


def export_excel(df):
    """导出 Excel 数据表。"""
    try:
        df.to_excel("豆瓣电影Top250.xlsx", index=False)
        print("Excel数据导出成功：豆瓣电影Top250.xlsx")
    except PermissionError:
        print("Excel文件被WPS/Excel占用！关闭表格后重新运行即可")


def setup_chinese_font():
    """配置 matplotlib 中文字体。"""
    plt.rcParams["font.sans-serif"] = ["SimHei"]
    plt.rcParams["axes.unicode_minus"] = False


def plot_score_distribution(df):
    """生成评分分布柱状图。"""
    setup_chinese_font()
    score_count = df["评分"].value_counts().sort_index()

    plt.figure(figsize=(11, 6))
    score_count.plot(kind="bar", color="#2E86AB")
    plt.title("豆瓣Top250电影评分分布统计", fontsize=14)
    plt.xlabel("电影评分", fontsize=12)
    plt.ylabel("对应电影数量", fontsize=12)
    plt.tight_layout()
    plt.savefig("评分分布.png", dpi=300)
    plt.close()
    print("评分柱状图已保存：评分分布.png")


def add_year_column(df):
    """用正则从导演年份信息中提取上映年份，写入新列。"""
    pattern = re.compile(r"\d{4}")
    df["上映年份"] = [
        (res[0] if (res := pattern.findall(info)) else "未知")
        for info in df["导演年份信息"]
    ]
    return df


def plot_year_distribution(df):
    """生成上映年份分布图。"""
    setup_chinese_font()
    year_count = df["上映年份"].value_counts().sort_index()

    plt.figure(figsize=(12, 6))
    year_count.plot(kind="bar", color="#6A994E")
    plt.title("豆瓣Top250电影上映年份分布", fontsize=14)
    plt.xlabel("上映年份", fontsize=12)
    plt.ylabel("电影数量", fontsize=12)
    plt.tight_layout()
    plt.savefig("年份分布.png", dpi=300)
    plt.close()
    print("年份分布图表已保存：年份分布.png")


def generate_wordcloud(df):
    """根据一句话简介生成词云，返回是否生成成功。"""
    all_brief = " ".join(df["一句话简介"].dropna().tolist())
    if not all_brief.strip():
        print("暂无有效简介文本，词云无法生成")
        return False

    wc = WordCloud(
        background_color="white",
        font_path="simhei.ttf",
        width=1000,
        height=600,
    ).generate(all_brief)
    wc.to_file("简介词云.png")
    print("简介词云已保存：简介词云.png")
    return True


def build_html_report(wordcloud_exists):
    """生成整合所有图表的静态 HTML 分析报告。"""
    if wordcloud_exists:
        brief_html = """
        <h2>三、电影简介关键词词云</h2>
        <img src="简介词云.png" alt="简介词云">
        """
    else:
        brief_html = """
        <h2>三、电影简介关键词词云</h2>
        <p>注：因豆瓣Top250部分电影无简介文本，词云无法生成。</p>
        """

    html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>豆瓣电影Top250 数据分析报告</title>
    <style>
        body {{text-align: center; font-family: "微软雅黑"; margin: 30px; background-color: #f9f9f9;}}
        .container {{background: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0,0,0,0.1);}}
        h1 {{color: #2E86AB;}}
        h2 {{color: #6A994E; margin-top: 40px;}}
        img {{margin: 10px 0; border: 1px solid #eee; max-width: 90%;}}
        p {{font-size: 16px; color: #333;}}
    </style>
</head>
<body>
    <div class="container">
        <h1>豆瓣电影Top250 数据分析可视化页面</h1>
        <p>本页面由Python程序自动生成，整合所有统计图表与分析结果</p>

        <h2>一、电影评分分布统计</h2>
        <img src="评分分布.png" alt="评分分布图">

        <h2>二、电影上映年份分布统计</h2>
        <img src="年份分布.png" alt="年份分布图">

        {brief_html}
    </div>
</body>
</html>
"""
    with open("数据分析报告.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    print("综合静态网页已生成：数据分析报告.html（浏览器直接打开）")


def main():
    clean_old_files()
    choice = show_menu()

    df = crawl_top250()

    if choice in ("1", "4"):
        export_excel(df)

    if choice in ("2", "4"):
        plot_score_distribution(df)

    if choice in ("3", "4"):
        add_year_column(df)
        plot_year_distribution(df)
        wc_exists = generate_wordcloud(df)
        build_html_report(wc_exists)

    print("\n所有选定流程执行完毕！")


if __name__ == "__main__":
    main()
