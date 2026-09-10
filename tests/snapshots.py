# -*- coding: utf-8 -*-
"""离屏截图：把新界面（班级总览/班级详情/月历排课/便签墙/成绩报告）渲染成 PNG"""
import importlib.util
import os
import sys
import tempfile
import time
from datetime import date, timedelta
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

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
OUT = Path(__file__).resolve().parent / "shots"
OUT.mkdir(exist_ok=True)

workdir = tempfile.mkdtemp(prefix="tms_shot_")
os.chdir(workdir)
os.environ["TMS_DATA_DIR"] = workdir

spec = importlib.util.spec_from_file_location("teaching_shot", TARGET)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

from PySide6.QtGui import QImage
from PySide6.QtWidgets import QApplication

db = mod.DatabaseManager()
svc = mod.DataImportService(db)
svc.import_students(str(DATA / "students_sample.xlsx"))
svc.import_grades(str(DATA / "grades_sample.xlsx"))

today = date.today()
# 当天与本月多天的课程（体现“月历标记有课的日子”）
for offset, items in {
    0: [("语文", "08:00", "09:40", "张老师", "A101", "#4F6EF7"),
        ("数学", "10:00", "11:40", "李老师", "A102", "#00B894"),
        ("班会", "15:00", "15:40", "王老师", "A103", "#6C5CE7")],
    2: [("英语", "08:00", "09:40", "刘老师", "B201", "#E17055")],
    5: [("物理", "14:00", "15:40", "陈老师", "B202", "#00CEC9")],
    9: [("化学", "09:00", "10:40", "赵老师", "B203", "#FD79A8")],
}.items():
    d = today + timedelta(days=offset)
    for name, s, e, t, loc, color in items:
        db.add_course(mod.Course(name=name, course_date=d.isoformat(), start_time=s,
                                 end_time=e, teacher=t, location=loc, color=color,
                                 description=""))

# 便签墙：文字 / 表格 / 图片
db.add_feedback(mod.Feedback(content_type='text',
                             content='明天上午第三节公开课，记得带实验报告和听课记录本。',
                             color=mod.STICKER_COLORS[0]))
db.add_feedback(mod.Feedback(
    content_type='table',
    table_json=mod.json.dumps({'headers': ['姓名', '语文', '数学', '英语'],
                               'rows': [['张三', '92', '98', '88'],
                                        ['李四', '85', '91', '90'],
                                        ['王五', '78', '84', '81']]}, ensure_ascii=False),
    color=mod.STICKER_COLORS[3]))
img = QImage(360, 200, QImage.Format_RGB32)
for y in range(200):
    for x in range(360):
        img.setPixel(x, y, (0xFF000000 | ((x * 255 // 360) << 16) | ((y * 255 // 200) << 8) | 0x99))
name = mod.save_qimage_file(img)
db.add_feedback(mod.Feedback(content_type='image', image_path=name,
                             image_caption='板书照片示例（含批注位置）',
                             color=mod.STICKER_COLORS[6]))
db.add_feedback(mod.Feedback(content_type='text',
                             content='家长会要点：1) 期中成绩分析　2) 晚自习安排　3) 春游报名',
                             color=mod.STICKER_COLORS[2]))

app = QApplication(sys.argv)
chosen_font = mod.apply_ui_font(app)
print("界面字体:", chosen_font)
w = mod.MainWindow()
w.resize(1360, 860)
w.show()
for _ in range(8):
    app.processEvents()


def wait(ms: int = 420):
    """真实等待若干毫秒（动画需要时间推进，仅 processEvents 不会前进）"""
    from PySide6.QtCore import QElapsedTimer
    timer = QElapsedTimer()
    timer.start()
    while timer.elapsed() < ms:
        app.processEvents()
        time.sleep(0.01)
    app.processEvents()


def shoot(name: str, widget=None):
    wait(320)                       # 等过渡动画结束再截图
    pix = (widget or w).grab()
    path = OUT / f"{name}.png"
    pix.save(str(path))
    print(f"  {name}.png  {path.stat().st_size}")


# 1) 学生管理 · 班级总览
w.switch_page(0)
sm = w.student_module
sm.refresh()
wait(300)
print(f"  班级数={len(sm._classes())} 班级磁贴={sm.class_layout.count()}")
shoot("1_学生管理_班级总览")

# 2) 点入某个班级（班级详情，等滑动动画结束）
first_class = next(iter(sm._classes()))
sm.open_class(first_class)
wait(700)
sm.stack._settle()               # 确保动画完成、页面归位
wait(120)
print(f"  进入班级={first_class} 学生磁贴={sm.tiles_layout.count()} 页索引={sm.stack.currentIndex()}")
shoot("2_学生管理_班级详情")

# 3) 排课 · 月历 + 当天课表
w.switch_page(1)
cm = w.course_module
cm.refresh()
cm.on_date_selected(today)
wait(400)
print(f"  标记天数={len(cm.calendar.day_counts)} 当天课程行={sum(1 for i in range(cm.day_layout.count()) if isinstance(cm.day_layout.itemAt(i).widget(), mod.CourseRow))}")
shoot("3_排课_月历与当天课表")

# 4) 便签墙
w.switch_page(2)
w.feedback_module.refresh()
wait(300)
print(f"  便签数={len(w.feedback_module.feedbacks)}")
shoot("4_便签墙_文字表格图片")

# 5) 成绩 · 个人报告
w.switch_page(3)
gm = w.grade_module
gm.refresh()
if gm.students:
    gm.current_student_id = gm.students[0].id
    gm.refresh_charts()
for _ in range(8):
    app.processEvents()
shoot("5_成绩_统计与个人报告")

w.close()
db.close()
print("screenshots ->", OUT)
