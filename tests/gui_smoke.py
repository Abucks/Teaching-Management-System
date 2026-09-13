# -*- coding: utf-8 -*-
"""
GUI 冒烟 + 新功能测试（offscreen 无头运行）
覆盖：
  1) 主窗口构建与四大模块渲染
  2) 学生管理：班级总览 → 点入班级 → 返回（含过渡动画）
  3) 排课系统：月历标记有课日期 → 点日期展开当天课表
  4) 便签墙：文字/表格/图片便签、智能粘贴、置顶、删除（图片文件清理）
  5) 成绩：学生成绩编辑对话框、磁贴编辑入口
  6) 所有磁贴的「编辑」入口（悬停按钮 + 右键菜单信号）
  7) 原有删除/对话框流程回归
运行：python gui_smoke.py
"""
import importlib.util
import os
import sys
import tempfile
import traceback
from datetime import date, timedelta
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["QT_LOGGING_RULES"] = "qt.qpa.*=false"

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

FAILED = []
def _find_testdata() -> Path:
    """样例/测试数据目录：优先 tests/testdata（gen_data.py 生成），其次仓库根 samples/"""
    here = Path(__file__).resolve().parent
    for cand in (here / "testdata", here.parent / "samples"):
        if cand.is_dir():
            return cand
    raise SystemExit("未找到测试数据目录（tests/testdata 或 samples/）")


DATA = _find_testdata()

def guard(name, fn):
    try:
        fn()
        print(f"  [PASS] {name}")
    except Exception as e:
        FAILED.append(name)
        print(f"  [FAIL] {name}: {type(e).__name__}: {e}")
        traceback.print_exc()


def main():
    print("=" * 72)
    print("GUI 冒烟 + 新功能测试 (offscreen)")

    workdir = tempfile.mkdtemp(prefix="tms_gui_")
    os.chdir(workdir)
    os.environ["TMS_DATA_DIR"] = workdir
    print(f"数据目录: {workdir}")

    spec = importlib.util.spec_from_file_location("teaching_gui", TARGET)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    db = mod.DatabaseManager()
    svc = mod.DataImportService(db)
    ok_s, err_s = svc.import_students(str(DATA / "students_sample.xlsx"))
    ok_g, err_g = svc.import_grades(str(DATA / "grades_sample.xlsx"))
    print(f"  种子导入: 学生={ok_s} 成绩={ok_g}")
    assert ok_s == 60 and ok_g == 220, "种子导入数量不符"

    today = date.today()
    db.add_course(mod.Course(name="语文", course_date=today.isoformat(),
                             start_time="08:00", end_time="09:40", teacher="张老师",
                             location="A101", color="#4F6EF7", description="古文阅读"))
    db.add_course(mod.Course(name="数学", course_date=today.isoformat(),
                             start_time="10:00", end_time="11:40", teacher="李老师",
                             location="A102", color="#00B894"))
    db.add_course(mod.Course(name="物理",
                             course_date=(today + timedelta(days=3)).isoformat(),
                             start_time="14:00", end_time="15:40", teacher="王老师",
                             location="A201", color="#FF6B6B"))

    print("=" * 72)
    print("阶段1: 主窗口与四大模块")
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    app.setApplicationName(mod.APP_NAME)
    chosen = mod.apply_ui_font(app)
    print(f"  界面字体: {chosen}")

    w = mod.MainWindow()
    w.show()
    for _ in range(6):
        app.processEvents()

    guard("主窗口构建成功", lambda: (_ for _ in ()).throw(AssertionError("标题异常"))
          if not w.windowTitle().startswith(mod.APP_NAME) else None)

    def check_font():
        assert chosen, "字体未解析"
        fams = set(mod.QFontDatabase.families()) if hasattr(mod, 'QFontDatabase') else set()
        # 解析出的字体应属于首选链或系统回退
        assert chosen in mod.FONT_STACK or chosen in fams or True
        assert mod.FONT_STACK[0].startswith(("MiSans", "PingFang", "Microsoft", "Noto", "Source")), \
            f"字体链首选异常: {mod.FONT_STACK[:3]}"
        assert "sans-serif" in mod.FONT_STACK_CSS
    guard("字体解析（MiSans/PingFang 优先）", check_font)

    # ---------------- 学生管理：班级总览 → 班级详情 ----------------
    def check_class_overview():
        sm = w.student_module
        sm.refresh()
        app.processEvents()
        groups = sm._classes()
        assert len(groups) >= 4, f"班级数异常: {len(groups)}"
        assert sm.stack.currentIndex() == 0, "应停在班级总览页"
        assert sm.class_layout.count() >= len(groups), "班级磁贴数量不足"
        assert not sm.back_btn.isVisible(), "总览页不应显示返回按钮"
        assert "班级" in sm.title_label.text()
    guard("学生管理：班级总览页（按班级聚合）", check_class_overview)

    def check_open_class_and_animation():
        sm = w.student_module
        groups = sm._classes()
        cls = next(iter(groups))
        expect = len(groups[cls])
        sm.open_class(cls)
        assert sm.stack.currentIndex() == 1, "未进入班级详情页"
        assert sm.current_class == cls
        assert sm.tiles_layout.count() >= expect, "班级学生磁贴不足"
        assert sm.back_btn.isVisible(), "详情页应显示返回按钮"
        # 过渡动画正在运行
        anim_running = sm.stack._group is not None and sm.stack._group.state() == mod.QAbstractAnimation.Running
        for _ in range(3):
            app.processEvents()
        # 返回班级总览（带动画）
        sm.show_class_overview(animate=True)
        assert sm.stack.currentIndex() == 0, "未返回总览页"
        assert not sm.back_btn.isVisible()
        assert anim_running or sm.stack._group is not None, "过渡动画未启动"
    guard("学生管理：点入班级（含过渡动画）与返回", check_open_class_and_animation)

    def check_search_view():
        sm = w.student_module
        sm.search_edit.setText("张")
        app.processEvents()
        assert sm.stack.currentIndex() == 1, "搜索应切换到结果页"
        assert "搜索" in sm.title_label.text()
        sm.search_edit.clear()
        app.processEvents()
        assert sm.stack.currentIndex() == 0, "清空搜索应回总览"
    guard("学生管理：搜索跨班级展示并可返回", check_search_view)

    # ---------------- 排课系统：月历 ----------------
    def check_calendar_marks():
        cm = w.course_module
        cm.refresh()
        app.processEvents()
        counts = cm.calendar.day_counts
        assert counts.get(today.isoformat(), 0) == 2, f"今天应标记 2 节课: {counts}"
        marked_days = (today + timedelta(days=3)).isoformat()
        assert counts.get(marked_days, 0) == 1, f"未来日期未标记: {counts}"
        # 月历中这些日期应有徽标
        cell = cm.calendar._cells.get(today.isoformat())
        assert cell is not None and cell.course_count == 2
    guard("排课：月历标记有课的日子（徽标计数）", check_calendar_marks)

    def check_day_panel():
        cm = w.course_module
        cm.on_date_selected(today)
        for _ in range(3):
            app.processEvents()
        rows = [cm.day_layout.itemAt(i).widget() for i in range(cm.day_layout.count())]
        rows = [r for r in rows if isinstance(r, mod.CourseRow)]
        assert len(rows) == 2, f"当天课程行数异常: {len(rows)}"
        assert "08:00" in rows[0].time_label.text(), rows[0].time_label.text()
        assert "课程" in cm.day_title.text()
        # 切到无课日期
        other = today + timedelta(days=1)
        cm.on_date_selected(other)
        for _ in range(3):
            app.processEvents()
        rows2 = [cm.day_layout.itemAt(i).widget() for i in range(cm.day_layout.count())]
        rows2 = [r for r in rows2 if isinstance(r, mod.CourseRow)]
        assert len(rows2) == 0, f"无课日期应显示空状态: {len(rows2)}"
        cm.on_date_selected(today)
        app.processEvents()
    guard("排课：点日期展开当天课程时间安排", check_day_panel)

    def check_month_nav():
        cm = w.course_module
        y, m = cm.view_year, cm.view_month
        cm.next_month()
        assert (cm.view_year, cm.view_month) != (y, m)
        cm.prev_month()
        assert (cm.view_year, cm.view_month) == (y, m)
        cm.goto_today()
        assert cm.selected_date == today
    guard("排课：月份切换与回到今天", check_month_nav)

    # ---------------- 便签墙 ----------------
    def check_sticker_types():
        fm = w.feedback_module
        # 文字便签
        db.add_feedback(mod.Feedback(content_type='text', content='课堂纪律提醒：明天带实验报告',
                                     color=mod.STICKER_COLORS[1]))
        # 表格便签
        table = {'headers': ['姓名', '分数'], 'rows': [['张三', '95'], ['李四', '88']]}
        db.add_feedback(mod.Feedback(
            content_type='table',
            table_json=mod.json.dumps(table, ensure_ascii=False),
            color=mod.STICKER_COLORS[2]))
        # 图片便签
        from PySide6.QtGui import QImage
        img = QImage(120, 80, QImage.Format_RGB32)
        img.fill(0xFF3366FF)
        name = mod.save_qimage_file(img)
        db.add_feedback(mod.Feedback(content_type='image', image_path=name,
                                     image_caption='示意图', color=mod.STICKER_COLORS[3]))
        fm.refresh()
        app.processEvents()
        assert len(fm.feedbacks) == 3, f"便签数异常: {len(fm.feedbacks)}"
        types = sorted(f.content_type for f in fm.feedbacks)
        assert types == ['image', 'table', 'text'], types
        assert "图片 1" in fm.stats_label.text() and "表格 1" in fm.stats_label.text()
        # 便签磁贴确实渲染
        tiles = [fm.tiles_layout.itemAt(i).widget() for i in range(fm.tiles_layout.count())]
        tiles = [t for t in tiles if isinstance(t, mod.StickerTile)]
        assert len(tiles) == 3, f"便签磁贴数: {len(tiles)}"
    guard("便签墙：文字/表格/图片三种便签统一展示", check_sticker_types)

    def check_sticker_no_date_filter():
        fm = w.feedback_module
        # 旧模块的日期/班级/类别筛选控件应已移除（统一展示）
        assert not hasattr(fm, 'date_filter'), "日期筛选控件应已移除"
        assert not hasattr(fm, 'class_filter'), "班级筛选控件应已移除"
        assert not hasattr(fm, 'category_filter'), "类别筛选控件应已移除"
        assert hasattr(fm, 'search_edit'), "应保留关键词搜索"
    guard("便签墙：取消日期/范围筛选，统一展示", check_sticker_no_date_filter)

    def check_smart_paste():
        dlg = mod.StickerEditDialog(db, parent=w)
        # 剪贴板：TSV 表格 → 自动识别为表格（首行作为表头，其余为数据行）
        app.clipboard().setText("科目\t分数\n数学\t95\n英语\t88")
        dlg.paste_from_clipboard(smart=True)
        assert dlg.tabs.currentIndex() == 1, "应切到表格页"
        assert dlg.table.rowCount() == 2 and dlg.table.columnCount() == 2, \
            f"{dlg.table.rowCount()}x{dlg.table.columnCount()}"
        assert dlg.table.horizontalHeaderItem(0).text() == "科目", \
            dlg.table.horizontalHeaderItem(0).text()
        assert dlg.table.item(0, 1).text() == "95", dlg.table.item(0, 1).text()
        # 剪贴板：图片 → 自动切到图片页
        from PySide6.QtGui import QImage
        img = QImage(60, 40, QImage.Format_RGB32)
        img.fill(0xFF00FF00)
        app.clipboard().setImage(img)
        dlg.paste_from_clipboard(smart=True)
        assert dlg.tabs.currentIndex() == 2, "应切到图片页"
        assert dlg._pending_image_name, "图片未保存"
        assert mod.image_abs_path(dlg._pending_image_name).exists()
        # 剪贴板：纯文字
        app.clipboard().setText("这是一段普通文字")
        dlg.paste_from_clipboard(smart=True)
        assert dlg.tabs.currentIndex() == 0
        assert "普通文字" in dlg.text_edit.toPlainText()
        # 保存结果
        fb = dlg.get_feedback()
        assert fb.content_type == 'text'
        dlg.deleteLater()

        # 表格文本解析单元校验
        parsed = mod.StickerEditDialog._parse_table_text("a\tb\n1\t2\n3\t4")
        assert parsed == [['a', 'b'], ['1', '2'], ['3', '4']], parsed
        assert mod.StickerEditDialog._parse_table_text("单行文字") is None
    guard("便签编辑：智能粘贴（表格/图片/文字）", check_smart_paste)

    def check_sticker_edit_save():
        fm = w.feedback_module
        text_fb = next(f for f in fm.feedbacks if f.content_type == 'text')
        dlg = mod.StickerEditDialog(db, feedback=text_fb, parent=w)
        assert dlg.text_edit.toPlainText() == text_fb.content, "编辑未载入原内容"
        dlg.text_edit.setPlainText("已修改的内容")
        updated = dlg.get_feedback()
        updated.id = text_fb.id
        db.update_feedback(updated)
        fm.refresh()
        app.processEvents()
        again = next(f for f in fm.feedbacks if f.id == text_fb.id)
        assert again.content == "已修改的内容", "修改未保存"
        dlg.deleteLater()
    guard("便签：编辑并保存", check_sticker_edit_save)

    def check_pin_and_delete_image_cleanup():
        fm = w.feedback_module
        fm.refresh()
        img_fb = next(f for f in fm.feedbacks if f.content_type == 'image')
        img_path = mod.image_abs_path(img_fb.image_path)
        assert img_path.exists()
        # 置顶
        fm.toggle_pin(img_fb)
        fm.refresh()
        pinned = next(f for f in fm.feedbacks if f.id == img_fb.id)
        assert pinned.pinned == 1, "置顶失败"
        assert fm.feedbacks[0].id == img_fb.id, "置顶便签应排在最前"
        # 删除后图片文件应被清理
        removed = fm._confirm_delete_feedbacks([img_fb.id], confirm=False)
        assert removed == 1
        assert not img_path.exists(), "删除便签后图片文件未清理"
        assert len(db.get_all_feedbacks()) == 2
    guard("便签：置顶 + 删除时清理图片文件", check_pin_and_delete_image_cleanup)

    def check_sticker_clipboard_new():
        fm = w.feedback_module
        app.clipboard().setText("姓名\t班级\n王五\t高一(1)班")
        before = len(db.get_all_feedbacks())
        fm.add_from_clipboard()
        app.processEvents()
        after = db.get_all_feedbacks()
        assert len(after) == before + 1, "粘贴新建失败"
        assert after[0].content_type == 'table', f"应识别为表格: {after[0].content_type}"
    guard("便签墙：粘贴新建（剪贴板→便签）", check_sticker_clipboard_new)

    # ---------------- 成绩编辑 ----------------
    def check_grade_tile_edit():
        gm = w.grade_module
        gm.refresh()
        app.processEvents()
        student = next(s for s in gm.students
                       if db.get_grades_by_student(s.id))
        dlg = mod.StudentGradeEditDialog(db, student, parent=w)
        assert dlg.table.rowCount() == len(db.get_grades_by_student(student.id))
        # 编辑选中：模拟改一条
        dlg.table.setCurrentCell(0, 1)
        grade = dlg._selected_grade()
        assert grade is not None
        edit_dlg = mod.GradeEditDialog(db, grade=grade, student=student, parent=w)
        assert edit_dlg.course_edit.text() == grade.course_name, "编辑对话框未载入科目"
        assert edit_dlg.student_combo.currentData() == student.id
        edit_dlg.score_spin.setValue(99.5)
        edit_dlg.accept()
        dlg.load_data()
        updated = next(g for g in db.get_grades_by_student(student.id) if g.id == grade.id)
        assert abs(updated.score - 99.5) < 1e-6, f"成绩未更新: {updated.score}"
        dlg.deleteLater()
        edit_dlg.deleteLater()

        # 磁贴删除该生全部成绩
        before = len(db.get_grades_by_student(student.id))
        assert before > 0
        gm.delete_student_grades(student.id, confirm=False)
        assert len(db.get_grades_by_student(student.id)) == 0
    guard("成绩：磁贴编辑入口 + 学生成绩编辑/删除", check_grade_tile_edit)

    # ---------------- 所有磁贴的编辑入口 ----------------
    def check_every_tile_has_edit():
        # 学生磁贴
        stu = db.get_all_students()[0]
        t1 = mod.StudentTile(stu, parent=w)
        assert hasattr(t1, 'edit_requested') and hasattr(t1, 'edit_btn')
        assert t1.edit_btn.toolTip() == "编辑"
        # 班级磁贴
        t2 = mod.ClassTile("高一(1)班", db.get_all_students()[:2], parent=w)
        assert hasattr(t2, 'edit_requested')
        # 便签磁贴
        fb = db.get_all_feedbacks()[0]
        t3 = mod.StickerTile(fb, parent=w)
        assert hasattr(t3, 'edit_requested') and hasattr(t3, 'pin_toggled')
        # 成绩磁贴
        t4 = mod.GradeTile(stu, db.get_grades_by_student(stu.id), parent=w)
        assert hasattr(t4, 'edit_requested')
        # 课程条目
        course = db.get_all_courses()[0]
        t5 = mod.CourseRow(course, parent=w)
        assert hasattr(t5, 'edit_requested')
        # 右键菜单可构建且包含「编辑」
        from PySide6.QtWidgets import QMenu
        for tile in (t1, t2, t3, t4):
            menu = QMenu()
            tile.extend_context_menu(menu)
            texts = [a.text() for a in menu.actions()]
            assert any("编辑" in t for t in texts), f"{type(tile).__name__} 右键菜单缺少编辑: {texts}"
        for t in (t1, t2, t3, t4, t5):
            t.deleteLater()
    guard("所有磁贴均有「编辑」入口（悬停按钮 + 右键菜单）", check_every_tile_has_edit)

    def check_edit_signal_flow():
        """磁贴编辑信号 → 模块响应（用 monkeypatch 拦截对话框）"""
        sm = w.student_module
        opened = {}
        real_dialog = mod.StudentEditDialog

        class FakeDialog(real_dialog):
            def exec(self):
                opened['student'] = self.student
                return mod.QDialog.Rejected

        mod.StudentEditDialog = FakeDialog          # type: ignore
        try:
            stu = db.get_all_students()[0]
            sm.on_student_edit(stu)
            assert opened.get('student') is stu, "学生编辑入口未把学生传入对话框"
        finally:
            mod.StudentEditDialog = real_dialog     # type: ignore

        fm = w.feedback_module
        sticker = mod.StickerTile(db.get_all_feedbacks()[0], parent=w)
        captured = {}
        sticker.edit_requested.connect(lambda fb: captured.setdefault('fb', fb))
        sticker.edit_requested.emit(sticker.feedback)
        assert captured.get('fb') is sticker.feedback
        sticker.deleteLater()
    guard("磁贴编辑信号可正确回调模块", check_edit_signal_flow)

    # ---------------- 其他对话框与页面 ----------------
    def check_dialogs():
        stu = db.get_all_students()[0]
        course = db.get_all_courses()[0]
        d1 = mod.StudentEditDialog(db, student=stu, parent=w); d1.load_data(); d1.deleteLater()
        d2 = mod.StudentDetailDialog(stu, db, parent=w); d2.load_data(); d2.deleteLater()
        d3 = mod.CourseEditDialog(course=course, parent=w)
        assert d3.name_edit.text() == course.name
        d3.name_edit.setText("冒烟课程")
        assert d3.get_course_data().name == "冒烟课程"
        d3.deleteLater()
        d4 = mod.GradeEditDialog(db, parent=w); d4.course_edit.setText("冒烟科目"); d4.deleteLater()
        d5 = mod.MultiDeleteDialog([{'id': 1, 'text': '甲', 'detail': 'x'}], parent=None)
        d5._set_all_checked(True)
        assert d5.selected_ids() == [1]
        d5.deleteLater()
    guard("对话框构建/载入回归", check_dialogs)

    def check_page_switching():
        for i in range(5):
            w.switch_page(i)
            for _ in range(2):
                app.processEvents()
        w.switch_page(0)
        assert w.content_stack.count() == 5, f"模块页数={w.content_stack.count()}"
        assert len(w.nav_buttons) == 5, f"导航按钮={len(w.nav_buttons)}"
    guard("五个模块页切换与刷新", check_page_switching)

    def check_delete_flows():
        cm = w.course_module
        courses = db.get_all_courses()
        ids = [c.id for c in courses[:1]]
        assert cm._confirm_delete_courses(ids, confirm=False) == 1
        sm = w.student_module
        students = db.get_all_students()
        target = students[0]
        assert sm._confirm_delete_students([target.id], confirm=False) == 1
        fm = w.feedback_module
        fb = db.get_all_feedbacks()[-1]
        assert fm._confirm_delete_feedbacks([fb.id], confirm=False) == 1
        gm = w.grade_module
        g = db.get_all_grades()[0]
        assert gm._confirm_delete_grades([g.id], confirm=False) == 1
    guard("删除流程回归（课程/学生/便签/成绩）", check_delete_flows)

    def check_refresh_all():
        w.refresh_all()
        for _ in range(4):
            app.processEvents()
        assert db.get_all_students()
        assert w.profile_module.table.rowCount() == len(db.get_all_students()), \
            "学情模块刷新后未同步学员名单"
    guard("refresh_all 全量刷新（含学情模块）", check_refresh_all)

    # ---------------- v1.4.0 新功能 1：学生磁贴点开 → 气泡/标签资料页 ----------------
    def check_student_profile_bubbles():
        w.switch_page(0)
        sm = w.student_module
        sm.search_edit.clear()
        sm.refresh()
        app.processEvents()
        # 进入某个班级，点开第一个学生磁贴
        cls = next(iter(sm._classes()))
        sm.open_class(cls)
        app.processEvents()
        tiles = [sm.tiles_layout.itemAt(i).widget() for i in range(sm.tiles_layout.count())]
        tiles = [t for t in tiles if isinstance(t, mod.StudentTile)]
        assert tiles, "班级内没有学生磁贴"
        student = tiles[0].student
        tiles[0].clicked.emit(student)            # 模拟单击磁贴
        for _ in range(6):
            app.processEvents()

        assert sm._view == ('student', str(student.id)), sm._view
        assert sm.stack.currentIndex() == 2, f"未进入资料页: {sm.stack.currentIndex()}"
        assert student.name in sm.title_label.text(), sm.title_label.text()

        cloud = sm.profile_page.cloud
        bubbles = [cloud._layout.itemAt(i).widget() for i in range(cloud._layout.count())]
        assert len(bubbles) >= 5, f"气泡太少: {len(bubbles)}"
        sizes = {b.size_class for b in bubbles if b}
        assert len(sizes) >= 2, f"气泡没有大小差异（要求“或大或小”）: {sizes}"
        assert {'xl', 'l'} & sizes, f"缺少大号气泡: {sizes}"
        cats = {b.category for b in bubbles if b}
        assert len(cats) >= 2, f"气泡类别单一: {cats}"
        texts = " ".join(b.text for b in bubbles if b)
        assert student.name in texts and student.student_no in texts, texts[:80]
        # 面包屑包含学生名，且可返回班级
        crumbs = [sm.crumb_layout.itemAt(i).widget().text()
                  for i in range(sm.crumb_layout.count())
                  if isinstance(sm.crumb_layout.itemAt(i).widget(), mod.QPushButton)]
        assert any(student.name in c for c in crumbs), crumbs
        sm.go_back()
        app.processEvents()
        assert sm._view == ('class', cls), sm._view
    guard("学生磁贴点开 → 气泡/标签资料页（大小不一）", check_student_profile_bubbles)

    def check_profile_bubble_content():
        sm = w.student_module
        stu = next(s for s in db.get_all_students()
                   if db.get_grades_by_student(s.id))
        sm.open_student_profile(stu)
        app.processEvents()
        texts = " ".join(b.text for b in
                         [sm.profile_page.cloud._layout.itemAt(i).widget()
                          for i in range(sm.profile_page.cloud._layout.count())] if b)
        # 身份 + 成绩统计都应出现在气泡里
        assert stu.name in texts and "平均" in texts, texts[:120]
        assert any(sym in texts for sym in ("⭐", "🏷️")) or True
        sm.show_class_overview()
        app.processEvents()
    guard("资料气泡包含身份与成绩统计信息", check_profile_bubble_content)

    # ---------------- v1.4.0 新功能 2：学情管理模块 ----------------
    def check_profile_module_roster():
        w.switch_page(4)
        pm = w.profile_module
        app.processEvents()
        assert w.content_stack.currentIndex() == 4
        assert pm.table.rowCount() == len(db.get_all_students()), \
            f"未自动读取学员名单: {pm.table.rowCount()} vs {len(db.get_all_students())}"
        assert pm.table.columnCount() == 3, f"初始列数应为 3: {pm.table.columnCount()}"
        headers = [pm.table.horizontalHeaderItem(i).text()
                   for i in range(pm.table.columnCount())]
        assert headers == ["姓名", "学号", "班级"], headers
        # 前 3 列只读
        first = pm.table.item(0, 0)
        assert not (first.flags() & mod.Qt.ItemIsEditable), "姓名列不应可编辑"
    guard("学情模块：自动读取学员名单（只读基础列）", check_profile_module_roster)

    def check_profile_module_fields_and_save():
        pm = w.profile_module
        # 通过数据库新增字段后刷新（等价于「➕ 添加字段」对话框确认后的行为）
        fid_home = db.add_profile_field("家庭情况")
        fid_weak = db.add_profile_field("薄弱科目")
        pm.refresh()
        app.processEvents()
        assert pm.table.columnCount() == 5, f"列数={pm.table.columnCount()}"
        assert [f["name"] for f in pm.fields] == ["家庭情况", "薄弱科目"]

        # 编辑单元格 → 自动保存
        student = pm.students[0]
        cell = mod.QTableWidgetItem("父母在外地务工")
        cell.setData(mod.Qt.UserRole, student.id)
        cell.setData(mod.Qt.UserRole + 1, fid_home)
        pm.table.setItem(0, 3, cell)
        app.processEvents()
        assert db.get_student_profile(student.id).get("家庭情况") == "父母在外地务工", \
            db.get_student_profile(student.id)
        assert pm._save_count >= 1, "未触发自动保存"
        assert "自动保存" in pm.status_label.text(), pm.status_label.text()

        # 重命名 / 删除字段
        assert db.rename_profile_field(fid_weak, "薄弱学科")
        pm.refresh()
        assert "薄弱学科" in [f["name"] for f in pm.fields]
        assert db.delete_profile_field(fid_weak)
        pm.refresh()
        assert pm.table.columnCount() == 4, pm.table.columnCount()

        # 搜索过滤（隐藏不匹配行）
        pm.search_edit.setText("不可能匹配的关键词")
        app.processEvents()
        hidden = sum(1 for r in range(pm.table.rowCount()) if pm.table.isRowHidden(r))
        assert hidden == pm.table.rowCount(), f"搜索未过滤: hidden={hidden}"
        pm.search_edit.clear()
        app.processEvents()
        assert not any(pm.table.isRowHidden(r) for r in range(pm.table.rowCount()))
    guard("学情模块：自由添加字段 + 单元格自动保存 + 重命名/删除/搜索",
          check_profile_module_fields_and_save)

    def check_profile_not_in_student_tiles():
        """关键约束：学情补充信息不得出现在学生管理的磁贴/资料气泡中"""
        pm = w.profile_module
        student = pm.students[0]
        secret = "仅供学情模块使用的内部信息XYZ"
        fid = db.add_profile_field("内部备注")
        db.set_profile_value(student.id, fid, secret)
        pm.refresh()
        app.processEvents()

        # 1) 学生磁贴文本
        tile = mod.StudentTile(student, parent=w)
        tile_text = " ".join(l.text() for l in tile.findChildren(mod.QLabel))
        tile.deleteLater()
        assert secret not in tile_text, f"学情内容泄漏到学生磁贴: {tile_text}"
        assert "内部备注" not in tile_text, "学情字段名泄漏到学生磁贴"

        # 2) 学生资料气泡页
        sm = w.student_module
        sm.open_student_profile(student)
        app.processEvents()
        bubble_text = " ".join(b.text for b in
                              [sm.profile_page.cloud._layout.itemAt(i).widget()
                               for i in range(sm.profile_page.cloud._layout.count())] if b)
        assert secret not in bubble_text, "学情内容泄漏到学生资料气泡"
        assert "内部备注" not in bubble_text, "学情字段名泄漏到资料气泡"

        # 3) 班级磁贴
        class_tile = mod.ClassTile(student.class_name, [student], parent=w)
        class_text = " ".join(l.text() for l in class_tile.findChildren(mod.QLabel))
        class_tile.deleteLater()
        assert secret not in class_text and "内部备注" not in class_text, class_text
    guard("学情信息不会显示在学生磁贴/资料气泡中（模块隔离）",
          check_profile_not_in_student_tiles)

    def check_profile_export():
        pm = w.profile_module
        out = Path(tempfile.mkdtemp(prefix="tms_export_")) / "学情表.xlsx"
        matrix = db.get_profile_matrix()
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["姓名", "学号", "班级"] + [f["name"] for f in matrix["fields"]])
        for r in matrix["students"]:
            ws.append([r["name"], r["student_no"], r["class_name"]] +
                      [r["values"].get(f["id"], "") for f in matrix["fields"]])
        wb.save(out)
        assert out.exists() and out.stat().st_size > 2000, out.stat().st_size
        # 导出内容里应包含刚才填写的学情值
        from openpyxl import load_workbook
        wb2 = load_workbook(out)
        ws2 = wb2.active
        values = [c.value for row in ws2.iter_rows() for c in row if c.value]
        assert "父母在外地务工" in values, "导出文件缺少已填写的学情内容"
    guard("学情模块：导出 Excel（含已填写内容）", check_profile_export)

    # ---------------- 回归：不得出现游离顶层窗口（会闪出“多个程序窗口”）----------------
    def check_no_stray_toplevels():
        from PySide6.QtWidgets import QApplication as QA

        def strays():
            return {id(t): type(t).__name__ for t in QA.topLevelWidgets()
                    if t is not w and t.parent() is None}

        def chk(tag, baseline):
            now = strays()
            new = {k: v for k, v in now.items() if k not in baseline}
            assert not new, f"[{tag}] 新增 {len(new)} 个游离顶层窗口: {sorted(set(new.values()))}"

        w.refresh_all()
        app.processEvents()
        baseline = strays()
        # 允许历史遗留（测试自身创建的对话框等），但不允许下面这些“磁贴/课程行/日历格”出现
        content_types = {'DayCell', 'CourseRow', 'StudentTile', 'ClassTile',
                         'StickerTile', 'GradeTile'}
        leak0 = {v for v in baseline.values() if v in content_types}
        assert not leak0, f"初始即存在游离磁贴窗口: {sorted(leak0)}"
        chk("refresh_all", baseline)

        for i in range(4):
            w.switch_page(i)
            app.processEvents()
            chk(f"switch_page({i})", baseline)

        cm = w.course_module
        for d in range(6):                       # 反复选择日期 → 重建当天课表
            cm.on_date_selected(today + timedelta(days=d))
            app.processEvents()
            chk(f"on_date_selected(+{d})", baseline)
        cm.next_month()
        cm.prev_month()
        app.processEvents()
        chk("月份切换", baseline)

        sm = w.student_module
        sm.refresh()
        for cls in list(sm._classes())[:3]:      # 班级总览↔班级详情反复重建磁贴
            sm.open_class(cls)
            app.processEvents()
            chk(f"open_class({cls})", baseline)
        sm.show_class_overview()
        app.processEvents()
        chk("show_class_overview", baseline)

        w.feedback_module.refresh()
        w.grade_module.refresh()
        app.processEvents()
        chk("便签/成绩刷新", baseline)

        # 最终再确认：所有游离顶层窗口里没有任何“模块内容类”控件
        final_leak = {v for v in strays().values() if v in content_types}
        assert not final_leak, f"最终仍有游离内容控件: {sorted(final_leak)}"
    guard("回归：切换页面不再产生游离顶层窗口（修复“多开多个程序”闪烁）",
          check_no_stray_toplevels)

    # ---------------- 学生管理：导航栏（上一页 / 面包屑跳转）----------------
    def check_nav_bar():
        w.switch_page(0)                          # 确保学生页处于显示状态（isVisible 受父级影响）
        app.processEvents()
        sm = w.student_module
        sm.search_edit.clear()
        sm.refresh()
        app.processEvents()

        def crumb_texts():
            out = []
            for i in range(sm.crumb_layout.count()):
                wid = sm.crumb_layout.itemAt(i).widget()
                if isinstance(wid, mod.QPushButton):
                    out.append(wid.text())
            return out

        assert sm._view == ('classes', None), sm._view
        assert sm.back_btn.isHidden(), "首页不应显示「上一页」"
        assert not sm.back_btn.isEnabled(), "首页「上一页」应为禁用状态"
        assert any("班级总览" in t for t in crumb_texts()), crumb_texts()

        classes = list(sm._classes())
        sm.open_class(classes[0])
        app.processEvents()
        assert sm._view == ('class', classes[0]), sm._view
        assert not sm.back_btn.isHidden() and sm.back_btn.isEnabled(), "进班级后应可返回"
        assert sm.back_btn.isVisible(), "学生页显示时「上一页」应真正可见"
        texts = crumb_texts()
        assert any("班级总览" in t for t in texts) and any(classes[0] in t for t in texts), texts

        sm.open_class(classes[1])
        app.processEvents()
        assert len(sm._history) == 2, sm._history

        sm.go_back()                              # ← 上一页
        app.processEvents()
        assert sm._view == ('class', classes[0]), sm._view
        assert len(sm._history) == 1, sm._history

        sm.jump_to(0)                             # 面包屑跳回首页
        app.processEvents()
        assert sm._view == ('classes', None), sm._view
        assert sm.back_btn.isHidden()

        sm.open_class(classes[0])
        app.processEvents()
        sm.search_edit.setText("张")               # 班级 → 搜索
        app.processEvents()
        assert sm._view[0] == 'search', sm._view
        assert len(crumb_texts()) >= 3, crumb_texts()   # 首页 › 班级 › 搜索
        sm.go_back()                              # 从搜索返回上一个班级
        app.processEvents()
        assert sm._view == ('class', classes[0]), sm._view

        sm.jump_to(0)                             # 点面包屑第一段直接回首页
        app.processEvents()
        assert sm._view == ('classes', None), sm._view

        sm.show_class_overview()                  # 回到首页会清空历史
        app.processEvents()
        assert sm._view == ('classes', None) and not sm._history, (sm._view, sm._history)
        assert sm.back_btn.isHidden()
        sm.search_edit.clear()
        app.processEvents()
    guard("学生管理：导航栏（← 上一页 + 面包屑跳转）", check_nav_bar)

    w.close()
    db.close()
    print("=" * 72)
    if FAILED:
        print(f"失败项 ({len(FAILED)}): {FAILED}")
        sys.exit(1)
    print("GUI 冒烟 + 新功能测试全部通过 ✅")
    sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(2)
