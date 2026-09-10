# -*- coding: utf-8 -*-
"""
逻辑层自动化测试（不启动GUI）：
  1) 编译/导入脚本本身
  2) DatabaseManager 学生/课程/反馈/成绩 CRUD + 查询 + 过滤器 + 搜索
  3) FileImporter 读取自动生成的虚拟学生表格（xlsx/xls/docx/csv）与成绩表（xlsx/xls/csv）
  4) DataImportService 端到端导入（含重复学号去重、空行容错）
运行方式：cd 到本脚本所在目录后 python logic_test.py
"""
import importlib.util
import os
import re
import sys
import tempfile
import traceback
from datetime import date
from pathlib import Path

def _find_source() -> Path:
    """定位主程序源文件（teaching_manager.py 或旧的 deepseek_python_*.py）"""
    here = Path(__file__).resolve().parent
    for base in [here, here.parent, *here.parents]:
        for name in ("teaching_manager.py", "app.py"):
            cand = base / name
            if cand.exists():
                return cand
        hits = sorted(base.glob("deepseek_python_*.py"))
        if hits:
            return hits[0]
    raise SystemExit("未找到主程序源文件（teaching_manager.py）")


TARGET = _find_source()
def _find_testdata() -> Path:
    """样例/测试数据目录：优先 tests/testdata（gen_data.py 生成），其次仓库根 samples/"""
    here = Path(__file__).resolve().parent
    for cand in (here / "testdata", here.parent / "samples"):
        if cand.is_dir():
            return cand
    raise SystemExit("未找到测试数据目录（tests/testdata 或 samples/）")


DATA = _find_testdata()

PASSED = []

FAILED = []
def check(name, cond, detail=""):
    if cond:
        PASSED.append(name)
        print(f"  [PASS] {name}")
    else:
        FAILED.append(name)
        print(f"  [FAIL] {name}  {detail}")


def load_module():
    spec = importlib.util.spec_from_file_location("teaching_mod", TARGET)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    print("=" * 70)
    print("阶段1: 导入目标脚本（含语法/依赖检查）")
    mod = load_module()
    check("模块导入无异常", True)
    for flag, haslib in [("PySide6", mod.HAS_PYSIDE6), ("matplotlib", mod.HAS_MATPLOTLIB),
                         ("openpyxl", mod.HAS_OPENPYXL), ("python-docx", mod.HAS_DOCX),
                         ("xlrd", mod.HAS_XLRD)]:
        check(f"依赖 {flag} 可用", haslib)

    # 在临时目录内新建独立 DB（DatabaseManager 为单例，测试进程独享）
    workdir = tempfile.mkdtemp(prefix="tms_logic_")
    db = mod.DatabaseManager(str(Path(workdir) / "test.db"))
    check("数据库初始化", db.conn is not None)

    print("=" * 70)
    print("阶段2: 学生 CRUD")
    s1 = mod.Student(student_no="20260001", name="张三", class_name="高一(1)班",
                     tags=["课代表", "优秀"], titles=["三好学生"])
    s1_id = db.add_student(s1)
    check("新增学生返回ID>0", s1_id > 0, f"id={s1_id}")
    s2 = mod.Student(student_no="20260002", name="李四", class_name="高一(1)班", tags=[], titles=[])
    s2_id = db.add_student(s2)
    s3 = mod.Student(student_no="20261001", name="王五", class_name="高二(3)班", tags=["文艺"], titles=[])
    s3_id = db.add_student(s3)
    got = db.get_student_by_id(s1_id)
    check("按ID读取学生", got is not None and got.name == "张三" and got.student_no == "20260001")
    check("标签/头衔JSON往返", got.tags == ["课代表", "优秀"] and got.titles == ["三好学生"],
          f"tags={got.tags} titles={got.titles}")

    got.name = "张三丰"
    got.class_name = "高一(2)班"
    ok = db.update_student(got)
    got2 = db.get_student_by_id(s1_id)
    check("更新学生", ok and got2.name == "张三丰" and got2.class_name == "高一(2)班")

    all_s = db.get_all_students()
    check("get_all_students 数量=3", len(all_s) == 3, f"n={len(all_s)}")
    hits = db.search_students("李")
    check("按姓名搜索", len(hits) == 1 and hits[0].name == "李四")
    hits = db.search_students("高一")
    check("按班级搜索", len(hits) >= 1, f"n={len(hits)}")
    check("删除学生", db.delete_student(s2_id) and db.get_student_by_id(s2_id) is None)

    print("=" * 70)
    print("阶段3: 课程 CRUD")
    c1 = mod.Course(name="数学", course_date="2026-09-10", start_time="08:00", end_time="09:40",
                    teacher="陈老师", location="A101", color="#4F6EF7")
    c1_id = db.add_course(c1)
    c2 = mod.Course(name="英语", course_date="2026-09-10", start_time="10:00", end_time="11:40",
                    teacher="王老师", location="A102")
    c2_id = db.add_course(c2)
    c3 = mod.Course(name="物理", course_date="2026-09-11", start_time="08:00", end_time="09:40")
    c3_id = db.add_course(c3)
    check("新增课程", c1_id > 0 and c2_id > 0)
    bydate = db.get_courses_by_date("2026-09-10")
    check("按日期查课程=2", len(bydate) == 2, f"n={len(bydate)}")
    check("duration_minutes=100", bydate[0].duration_minutes == 100, f"dur={bydate[0].duration_minutes}")
    c1b = db.get_courses_by_date("2026-09-10")[0]
    c1b.name = "高等数学"
    check("更新课程", db.update_course(c1b))
    check("删除课程", db.delete_course(c3_id) and len(db.get_all_courses()) == 2)

    print("=" * 70)
    print("阶段4: 反馈 CRUD + 过滤器")
    f1 = mod.Feedback(date="2026-09-09", time="14:30", class_name="高一(1)班", category="学习",
                      content="课堂积极", priority=3)
    f1_id = db.add_feedback(f1)
    f2 = mod.Feedback(date="2026-09-09", time="15:00", class_name="高二(3)班", category="生活",
                      content="宿舍报修", priority=1)
    f2_id = db.add_feedback(f2)
    f3 = mod.Feedback(date="2026-09-10", time="09:00", class_name="高一(1)班", category="心理",
                      content="考前焦虑", priority=5)
    f3_id = db.add_feedback(f3)
    check("新增反馈", f1_id > 0 and f2_id > 0 and f3_id > 0)
    res = db.get_feedbacks_by_filter(date_str="2026-09-09")
    check("按日期过滤=2", len(res) == 2, f"n={len(res)}")
    res = db.get_feedbacks_by_filter(date_str="2026-09-09", class_name="高一(1)班")
    check("按日期+班级过滤=1", len(res) == 1, f"n={len(res)}")
    res = db.get_feedbacks_by_filter(category="学习")
    check("按类别过滤=1", len(res) == 1 and res[0].category == "学习", f"n={len(res)}")
    res = db.get_feedbacks_by_filter()
    check("无条件取全部=3", len(res) == 3, f"n={len(res)}")
    f1b = db.get_feedbacks_by_filter(date_str="2026-09-09")[0]
    f1b.status = "已解决"
    check("更新反馈", db.update_feedback(f1b))
    check("删除反馈", db.delete_feedback(f2_id))

    print("=" * 70)
    print("阶段5: 成绩 CRUD")
    g1 = mod.Grade(student_id=s1_id, student_name="张三丰", course_name="数学", score=92.5,
                   exam_date="2026-06-20", exam_type="期末", full_score=100.0,
                   class_rank=3, grade_rank=15)
    g1_id = db.add_grade(g1)
    g2 = mod.Grade(student_id=s1_id, student_name="张三丰", course_name="英语", score=88.0,
                   exam_date="2026-06-20", exam_type="期末", full_score=100.0)
    g2_id = db.add_grade(g2)
    g3 = mod.Grade(student_id=s3_id, student_name="王五", course_name="数学", score=76.5,
                   exam_date="2026-06-20", exam_type="期末", full_score=100.0)
    db.add_grade(g3)
    s1g = db.get_grades_by_student(s1_id)
    check("按学生取成绩=2", len(s1g) == 2, f"n={len(s1g)}")
    check("分数精度", abs(s1g[0].score - 92.5) < 1e-6 or abs(s1g[1].score - 92.5) < 1e-6)
    cg = db.get_grades_by_course("数学")
    check("按课程取成绩=2", len(cg) == 2, f"n={len(cg)}")
    all_g = db.get_all_grades()
    check("全部成绩=3", len(all_g) == 3, f"n={len(all_g)}")
    g1b = db.get_grades_by_student(s1_id)[0]
    g1b.score = 95.0
    check("更新成绩", db.update_grade(g1b))
    check("删除成绩", db.delete_grade(g2_id) and len(db.get_grades_by_student(s1_id)) == 1)

    print("=" * 70)
    print("阶段5.5: delete_grades_by_student 级联批量删除")
    before_g = len(db.get_grades_by_student(s1_id))
    db.add_grade(mod.Grade(student_id=s1_id, student_name="张三丰", course_name="化学", score=77.0,
                           exam_date="2026-05-10", exam_type="月考", full_score=100.0))
    db.add_grade(mod.Grade(student_id=s1_id, student_name="张三丰", course_name="生物", score=66.0,
                           exam_date="2026-05-11", exam_type="月考", full_score=100.0))
    n_now = len(db.get_grades_by_student(s1_id))
    check("级联删除前成绩数=原+2", n_now == before_g + 2, f"n={n_now}")
    removed = db.delete_grades_by_student(s1_id)
    check("级联删除返回条数正确", removed == n_now, f"removed={removed}")
    check("级联删除后该学生无成绩", len(db.get_grades_by_student(s1_id)) == 0)
    check("级联删除不影响其他学生成绩", len(db.get_grades_by_student(s3_id)) == 1)

    print("=" * 70)
    print("阶段5.6: 便签新字段（文字/表格/图片 + 置顶 + 搜索）")
    # 清空前面阶段遗留的反馈，保证本阶段计数独立
    for _old in db.get_all_feedbacks():
        db.delete_feedback(_old.id)
    check("清空历史反馈", len(db.get_all_feedbacks()) == 0)
    t_fb = db.add_feedback(mod.Feedback(content_type='text', content='便签文字内容',
                                        color=mod.STICKER_COLORS[1]))
    table_data = {'headers': ['科目', '分数'], 'rows': [['数学', '95'], ['英语', '88']]}
    tb_fb = db.add_feedback(mod.Feedback(
        content_type='table', table_json=mod.json.dumps(table_data, ensure_ascii=False)))
    from PySide6.QtGui import QImage as _QImage
    _img = _QImage(24, 12, _QImage.Format_RGB32)
    _img.fill(0xFF123456)
    img_name = mod.save_qimage_file(_img)
    im_fb = db.add_feedback(mod.Feedback(content_type='image', image_path=img_name,
                                         image_caption='示例截图'))
    all_fb = db.get_all_feedbacks()
    check("便签三条已入库", len(all_fb) == 3, f"n={len(all_fb)}")
    got_t = db.get_all_feedbacks()[-1] if False else next(f for f in all_fb if f.id == t_fb)
    check("文字便签字段往返", got_t.content == '便签文字内容'
          and got_t.content_type == 'text' and got_t.color == mod.STICKER_COLORS[1])
    got_tb = next(f for f in all_fb if f.id == tb_fb)
    check("表格便签可解析回表格", got_tb.table_data() == table_data, str(got_tb.table_data()))
    check("表格便签纯文本导出为 TSV", got_tb.plain_text().startswith("科目\t分数"),
          repr(got_tb.plain_text()[:20]))
    got_im = next(f for f in all_fb if f.id == im_fb)
    check("图片便签文件已落盘", mod.image_abs_path(got_im.image_path).exists())
    check("图片便签纯文本回退为说明", got_im.plain_text() == '示例截图')

    check("搜索命中表格内容", any(f.id == tb_fb for f in db.search_feedbacks('数学')))
    check("搜索命中图片说明", any(f.id == im_fb for f in db.search_feedbacks('截图')))
    check("搜索命中文字内容", any(f.id == t_fb for f in db.search_feedbacks('便签文字')))

    db.set_feedback_pinned(t_fb, True)
    ordered = db.get_all_feedbacks()
    check("置顶便签排在最前", ordered[0].id == t_fb and ordered[0].pinned == 1,
          f"first={ordered[0].id} pinned={ordered[0].pinned}")
    db.set_feedback_pinned(t_fb, False)
    check("取消置顶生效", db.get_all_feedbacks()[0].pinned == 0)

    upd = next(f for f in db.get_all_feedbacks() if f.id == tb_fb)
    upd.content_type = 'text'
    upd.content = '改成文字了'
    check("便签类型切换（表格→文字）", db.update_feedback(upd))
    check("切换后内容正确", next(f for f in db.get_all_feedbacks()
                                if f.id == tb_fb).content == '改成文字了')
    check("删除便签", db.delete_feedback(im_fb) and len(db.get_all_feedbacks()) == 2)
    check("删除图片文件", mod.delete_image_file(img_name)
          and not mod.image_abs_path(img_name).exists())

    print("=" * 70)
    print("阶段5.7: 月历接口（按月份统计每天课程数）")
    for d, n in [("2026-09-01", 2), ("2026-09-01", 1), ("2026-09-15", 3), ("2026-10-02", 4)]:
        for k in range(n):
            db.add_course(mod.Course(name=f"课{k}", course_date=d,
                                     start_time=f"{8+k:02d}:00", end_time=f"{9+k:02d}:00"))
    counts = db.get_course_day_counts(2026, 9)
    # 阶段3 曾在 2026-09-10 建过 2 门课，因此 9 月共有 3 天有课
    check("9月标记天数=3", len(counts) == 3, str(counts))
    check("9月1日 3 节课", counts.get("2026-09-01") == 3, str(counts))
    check("9月10日 2 节课（阶段3遗留）", counts.get("2026-09-10") == 2, str(counts))
    check("9月15日 3 节课", counts.get("2026-09-15") == 3, str(counts))
    check("10月不在9月结果内", "2026-10-02" not in counts)
    check("get_all_course_dates 含跨月日期", "2026-10-02" in db.get_all_course_dates())

    print("=" * 70)
    print("阶段5.8: 老版本数据库自动迁移（新增便签字段）")
    import sqlite3 as _sqlite3
    legacy_dir = tempfile.mkdtemp(prefix="tms_legacy_")
    legacy_path = Path(legacy_dir) / "legacy.db"
    conn = _sqlite3.connect(str(legacy_path))
    conn.execute('''CREATE TABLE feedbacks (
        id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, time TEXT DEFAULT '',
        class_name TEXT DEFAULT '', category TEXT DEFAULT '其他', content TEXT DEFAULT '',
        student_name TEXT DEFAULT '', student_id INTEGER, status TEXT DEFAULT '待处理',
        priority INTEGER DEFAULT 1, created_at TEXT DEFAULT '', updated_at TEXT DEFAULT '')''')
    conn.execute("INSERT INTO feedbacks (date, time, category, content) VALUES "
                 "('2026-01-01', '09:00', '学习', '老版本便签内容')")
    conn.commit()
    conn.close()

    mod.DatabaseManager._instance = None            # 重置单例，指向老库
    legacy_db = mod.DatabaseManager(str(legacy_path))
    cols = {row[1] for row in
            legacy_db._get_connection().execute("PRAGMA table_info(feedbacks)").fetchall()}
    check("迁移后新增列齐备",
          {'content_type', 'table_json', 'image_path', 'image_caption', 'color', 'pinned'} <= cols,
          str(sorted(cols)))
    legacy_rows = legacy_db.get_all_feedbacks()
    check("老数据仍可读取", len(legacy_rows) == 1 and legacy_rows[0].content == '老版本便签内容',
          str([r.content for r in legacy_rows]))
    check("老数据补默认值（text/颜色/未置顶）",
          legacy_rows[0].content_type == 'text' and legacy_rows[0].color == '#FFF6C9'
          and legacy_rows[0].pinned == 0,
          f"{legacy_rows[0].content_type}/{legacy_rows[0].color}/{legacy_rows[0].pinned}")
    check("迁移后可写入新便签",
          legacy_db.add_feedback(mod.Feedback(content_type='table',
                                              table_json='{"headers":["a"],"rows":[["1"]]}')) > 0)
    legacy_db.close()
    mod.DatabaseManager._instance = None            # 复位，后续阶段用独立库

    print("=" * 70)
    print("阶段6: FileImporter 读取虚拟学生表格")
    # 6a xlsx 学生表：60真实 + 2重复学号 + 1空行 = 有效60行（重复行也带姓名，会被读取为记录）
    st = mod.FileImporter.import_students_from_file(str(DATA / "students_sample.xlsx"))
    check("xlsx 学生读取>0", len(st) > 50, f"n={len(st)}")
    check("xlsx 含重复学号记录(容错:仍被读出)", any(s['name'].startswith("测试重复") for s in st),
          "（重复学号在导入服务层去重，读取层不拦截）")
    check("xlsx 字段完整性", all(s['name'] and s['student_no'] for s in st if s['name'] != "测试重复A"))
    # 6b csv
    st2 = mod.FileImporter.import_students_from_file(str(DATA / "students_sample.csv"))
    check("csv 学生读取>50", len(st2) > 50, f"n={len(st2)}")
    # 6c docx
    st3 = mod.FileImporter.import_students_from_file(str(DATA / "students_sample.docx"))
    check("docx 学生读取>50", len(st3) > 50, f"n={len(st3)}")
    # 6d xls
    st4 = mod.FileImporter.import_students_from_file(str(DATA / "students_sample.xls"))
    check("xls 学生读取>50", len(st4) > 50, f"n={len(st4)}")
    # 6e 不支持格式
    bad = Path(workdir) / "x.txt"
    bad.write_text("abc", encoding="utf-8")
    try:
        mod.FileImporter.import_students_from_file(str(bad))
        check("不支持格式抛错", False)
    except ValueError as e:
        check("不支持格式抛错", "不支持" in str(e) or "读取文件失败" in str(e), str(e))

    print("阶段7: FileImporter 读取虚拟成绩表")
    gs = mod.FileImporter.import_grades_from_file(str(DATA / "grades_sample.xlsx"))
    check("xlsx 成绩读取>100", len(gs) > 100, f"n={len(gs)}")
    check("xlsx 成绩字段", all('student_name' in g and 'score' in g and 'course_name' in g for g in gs[:5]))
    gs2 = mod.FileImporter.import_grades_from_file(str(DATA / "grades_sample.csv"))
    check("csv 成绩读取>100", len(gs2) > 100, f"n={len(gs2)}")
    gs3 = mod.FileImporter.import_grades_from_file(str(DATA / "grades_sample.xls"))
    check("xls 成绩读取>100", len(gs3) > 100, f"n={len(gs3)}")

    print("=" * 70)
    print("阶段7.5: Excel 真实日期单元格/空日期单元格容错")
    from datetime import datetime as _dt
    import openpyxl as _ox
    _wb = _ox.Workbook()
    _ws = _wb.active
    _ws.append(["姓名", "科目", "成绩", "考试日期", "考试类型"])
    _ws.append(["张三丰", "数学", 90.0, _dt(2026, 6, 1, 8, 30), "期末"])      # 真实日期单元格
    _ws.append(["王五", "语文", 80.0, _dt(2026, 6, 2), "期末"])
    _ws.append(["张三丰", "英语", 70.0, None, "期末"])                        # 空白日期格
    _ws.append(["王五", "英语", "85.5", "2026/6/3", "期末"])                  # 文本日期
    _datefile = Path(workdir) / "grades_dates.xlsx"
    _wb.save(str(_datefile))
    gd = mod.FileImporter.import_grades_from_file(str(_datefile))
    check("日期文件读满3行", len(gd) == 4, f"n={len(gd)}")
    _date_ok = all(re.fullmatch(r"\d{4}-\d{2}-\d{2}", g["exam_date"]) for g in gd)
    check("exam_date 全部规范为YYYY-MM-DD", _date_ok, f"values={[g['exam_date'] for g in gd]}")
    check("真实日期单元格 2026-06-01", gd[0]["exam_date"] == "2026-06-01", gd[0]["exam_date"])
    check("空白日期格回退为今天", gd[2]["exam_date"] == date.today().isoformat(), gd[2]["exam_date"])
    check("文本日期2026/6/3规范化", gd[3]["exam_date"] == "2026-06-03", gd[3]["exam_date"])
    check("空行仍被跳过", all(g["student_name"] for g in gd))

    print("=" * 70)
    print("阶段8: DataImportService 端到端（学生表 -> 成绩表）")
    # 重置单例，使用全新独立 DB 模拟"干净环境下的用户导入"
    mod.DatabaseManager._instance = None
    workdir2 = tempfile.mkdtemp(prefix="tms_import_")
    db2 = mod.DatabaseManager(str(Path(workdir2) / "import.db"))
    check("阶段8独立DB已初始化(不同路径)", str(Path(db2.db_path)) != str(Path(db.db_path))
          and Path(db2.db_path).name == "import.db")
    svc = mod.DataImportService(db2)

    ok_n, errs = svc.import_students(str(DATA / "students_sample.xlsx"))
    print(f"      导入学生: 成功={ok_n}, 错误={len(errs)}")
    check("学生导入成功=60", ok_n == 60, f"ok={ok_n}")
    check("重复学号被去重跳过(=2)", len(errs) == 2 and all("已存在" in e for e in errs),
          f"errs={errs}")
    all_s2 = db2.get_all_students()
    check("库中学生数与成功数一致", len(all_s2) == ok_n, f"db={len(all_s2)} ok={ok_n}")

    # 同一文件再导一次：文件共62条可读记录(60真实+2重复学号行) 学号均已在库 -> 全部跳过
    ok_n2, errs2 = svc.import_students(str(DATA / "students_sample.xlsx"))
    total_rows = len(mod.FileImporter.import_students_from_file(str(DATA / "students_sample.xlsx")))
    check("重复导入全被跳过", ok_n2 == 0 and len(errs2) == total_rows,
          f"ok={ok_n2} errs={len(errs2)} (预期=文件可读行数 {total_rows})")

    # 成绩导入（姓名匹配学生；2个测试专用名 测试重复A/B 不在库中属正常失配）
    g_ok, g_errs = svc.import_grades(str(DATA / "grades_sample.xlsx"))
    print(f"      导入成绩: 成功={g_ok}, 错误={len(g_errs)}")
    check("成绩导入成功=220", g_ok == 220, f"ok={g_ok}")
    g_in_db = db2.get_all_grades()
    check("成绩已入库", len(g_in_db) == g_ok, f"db={len(g_in_db)}")
    known = {s.name: s.id for s in db2.get_all_students()}
    test_only = {"测试重复A", "测试重复B"}
    bad_map = [g for g in g_in_db if g.student_name in known and g.student_id != known[g.student_name]]
    unmapped = [g.student_name for g in g_in_db if g.student_name not in known and g.student_name not in test_only]
    check("已知学生的成绩均正确关联", not bad_map, f"错配={bad_map[:3]}")
    check("未知学生仅限测试专用名(容错不崩溃)", not unmapped, f"意外未匹配={unmapped[:5]}")

    db.close()
    db2.close()

    print("=" * 70)
    print(f"汇总: PASS={len(PASSED)}  FAIL={len(FAILED)}")
    if FAILED:
        print("失败项:")
        for f in FAILED:
            print("   -", f)
        sys.exit(1)
    print("全部逻辑测试通过 ✅")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(2)
