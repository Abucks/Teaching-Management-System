# -*- coding: utf-8 -*-
"""
自动生成虚拟学生信息表格（测试用样例数据）
生成格式：xlsx / csv / docx（学生表），xlsx / csv（成绩表）
并有意埋入边界情况用于测试容错：
  - 空行、缺列、学号重复、中文逗号分隔的标签、数字型学号等
"""
import csv
import random
from pathlib import Path
from datetime import date

from openpyxl import Workbook
from docx import Document

OUT = Path(__file__).resolve().parent / "testdata"   # 生成到 tests/testdata（测试默认读取这里）
OUT.mkdir(parents=True, exist_ok=True)

random.seed(20260909)

SURNAMES = "王李张刘陈杨黄赵吴周徐孙马朱胡郭何高林罗郑梁谢宋唐许韩冯邓曹彭曾肖田董袁潘于蒋蔡余杜叶程苏魏吕丁任沈姚卢姜崔钟谭陆汪范金石廖贾夏韦付方白邹孟熊秦邱江尹薛闫段雷侯龙史陶黎贺顾毛郝龚邵万钱严覃武戴莫孔向汤"
GIVEN = ["伟", "芳", "娜", "敏", "静", "磊", "洋", "勇", "艳", "杰", "涛", "明", "超", "秀英", "霞", "平", "刚", "桂英", "建华", "文", "辉", "丽", "晨", "宇", "子涵", "欣怡", "浩然", "梓萱", "雨桐", "思远"]

CLASSES = ["高一(1)班", "高一(2)班", "高二(3)班", "高三(1)班", "高三(2)班"]
TAG_POOL = [["优秀班干部", "数学课代表"], ["英语课代表"], ["体育特长生"], ["学习标兵"], [], ["文艺骨干"], ["编程爱好者"], [], ["值日组长", "物理课代表"]]
TITLE_POOL = [["三好学生"], ["进步之星"], [], ["优秀团员"], ["文明之星"], [], ["学习之星"]]


def random_name():
    return random.choice(SURNAMES) + random.choice(GIVEN)


def make_students(count=60, class_indexes=None):
    """生成 (学号, 姓名, 班级, 标签串, 头衔串) 记录，含部分重复学号"""
    students = []
    used_nos = set()
    for i in range(count):
        cls = CLASSES[random.randrange(len(CLASSES))] if class_indexes is None else CLASSES[random.choice(class_indexes)]
        while True:
            no = f"2026{random.randrange(1000, 9999)}"
            if no not in used_nos:
                used_nos.add(no)
                break
        name = random_name()
        tags = random.choice(TAG_POOL)
        titles = random.choice(TITLE_POOL)
        students.append((no, name, cls, ",".join(tags), "，".join(titles)))
    # 制造 2 条重复学号（不同姓名）用于测试去重逻辑
    students.append((students[0][0], "测试重复A", students[0][2], "", ""))
    students.append((students[3][0], "测试重复B", students[3][2], "", ""))
    # 制造 1 条空姓名行占位（应被跳过）
    students.append(("", "", "", "", ""))
    return students


def write_students_xlsx(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "学生名单"
    ws.append(["学号", "姓名", "班级", "标签", "头衔"])
    for no, name, cls, tags, titles in records:
        if no or name or cls or tags or titles:
            # 学号写成数字格式，考验读取容错（str()转换）
            try:
                no_cell = int(no) if no else None
            except ValueError:
                no_cell = no
            ws.append([no_cell, name, cls, tags, titles])
        else:
            ws.append([None, None, None, None, None])
    # 在表格后追加一行完全空行
    ws.append([None, None, None, None, None])
    wb.save(OUT / "students_sample.xlsx")


def write_students_csv(records):
    p = OUT / "students_sample.csv"
    with open(p, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["学号", "姓名", "班级", "标签", "头衔"])
        for no, name, cls, tags, titles in records:
            w.writerow([no, name, cls, tags, titles])
    # 追加空行
    with open(p, "a", newline="", encoding="utf-8-sig") as f:
        f.write("\n")


def write_students_docx(records):
    doc = Document()
    doc.add_heading("高一学生名单", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.rows[0].cells[0].text = "学号"
    table.rows[0].cells[1].text = "姓名"
    table.rows[0].cells[2].text = "班级"
    table.rows[0].cells[3].text = "标签"
    table.rows[0].cells[4].text = "头衔"
    for no, name, cls, tags, titles in records:
        if not name:
            continue
        row = table.add_row().cells
        row[0].text = no
        row[1].text = name
        row[2].text = cls
        row[3].text = tags
        row[4].text = titles
    doc.save(OUT / "students_sample.docx")


def make_grades(students, count=180):
    """按真实学生姓名生成成绩记录"""
    valid = [s for s in students if s[1]]
    # 用姓名找到对应记录（注意重复学号也保留了姓名，姓名仍唯一）
    names = [s[1] for s in valid]
    grades = []
    subjects = ["语文", "数学", "英语", "物理", "化学", "生物", "历史", "地理", "政治"]
    exam_types = ["期中", "期末", "月考"]
    for _ in range(count):
        name = random.choice(names)
        subj = random.choice(subjects)
        exam_type = random.choice(exam_types)
        score = round(random.uniform(42.0, 99.5), 1)
        exam_date = date(2026, random.randint(3, 9), random.randint(1, 28)).isoformat()
        grades.append((name, subj, score, exam_date, exam_type))
    return grades


def write_grades_xlsx(grades):
    wb = Workbook()
    ws = wb.active
    ws.title = "成绩"
    ws.append(["姓名", "科目", "成绩", "考试日期", "考试类型"])
    for name, subj, score, d, t in grades:
        ws.append([name, subj, score, d, t])
    wb.save(OUT / "grades_sample.xlsx")


def write_grades_csv(grades):
    p = OUT / "grades_sample.csv"
    with open(p, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["姓名", "科目", "成绩", "考试日期", "考试类型"])
        for name, subj, score, d, t in grades:
            w.writerow([name, subj, score, d, t])


def write_xls_files(records, grades):
    """.xls 生成依赖 xlwt（仅测试数据生成用）"""
    try:
        import xlwt
    except ImportError:
        print("[skip] xlwt 未安装，跳过 .xls 样例生成（可用 pip install xlwt）")
        return
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("学生名单")
    for c, h in enumerate(["学号", "姓名", "班级", "标签", "头衔"]):
        ws.write(0, c, h)
    r = 1
    for no, name, cls, tags, titles in records:
        if not name:
            continue
        ws.write(r, 0, no)
        ws.write(r, 1, name)
        ws.write(r, 2, cls)
        ws.write(r, 3, tags)
        ws.write(r, 4, titles)
        r += 1
    wb.save(str(OUT / "students_sample.xls"))

    wb2 = xlwt.Workbook(encoding="utf-8")
    ws2 = wb2.add_sheet("成绩")
    for c, h in enumerate(["姓名", "科目", "成绩", "考试日期", "考试类型"]):
        ws2.write(0, c, h)
    r = 1
    for name, subj, score, d, t in grades:
        ws2.write(r, 0, name)
        ws2.write(r, 1, subj)
        ws2.write(r, 2, float(score))
        ws2.write(r, 3, d)
        ws2.write(r, 4, t)
        r += 1
    wb2.save(str(OUT / "grades_sample.xls"))
    print("[ok] 已生成 .xls 样例")


def main():
    students = make_students(60)
    write_students_xlsx(students)
    write_students_csv(students)
    write_students_docx(students)
    grades = make_grades(students, 220)
    write_grades_xlsx(grades)
    write_grades_csv(grades)
    try:
        write_xls_files(students, grades)
    except Exception as e:
        print(f"[skip] xls 生成失败: {e}")

    print(f"[ok] 样例数据已写入 {OUT}")
    for f in sorted(OUT.iterdir()):
        print("   ", f.name, f.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
