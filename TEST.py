import sys
import os
import re
import matplotlib
import pandas as pd
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout,
                             QHBoxLayout, QFileDialog, QWidget, QTableWidget, QTableWidgetItem,
                             QMessageBox, QTabWidget, QComboBox, QProgressBar, QTextEdit)
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import datetime


# 清理字体缓存配置
matplotlib._cache_dir = None  # 禁用缓存

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Zen Hei']  # Windows常用字体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

class Student:
    def __init__(self, student_id, name, grade="", class_name=""):
        self.student_id = student_id
        self.name = name
        self.grade = grade  # 新增年级属性
        self.class_name = class_name  # 新增班级属性
        self.missing_experiments = []

    def add_missing_experiment(self, experiment_name):
        self.missing_experiments.append(experiment_name)


class Experiment:
    def __init__(self, name):
        self.name = name
        self.submitted_students = set()

    def add_submitted_student(self, student_id):
        self.submitted_students.add(student_id)

    def get_missing_students(self, all_students):
        all_ids = {student.student_id for student in all_students}
        return all_ids - self.submitted_students

    def get_submission_rate(self, total_students):
        if total_students == 0:
            return 0
        return len(self.submitted_students) / total_students * 100


class Class:
    def __init__(self, name):
        self.name = name
        self.experiments = {}  # 实验名 -> Experiment对象

    def add_experiment(self, experiment_name):
        if experiment_name not in self.experiments:
            self.experiments[experiment_name] = Experiment(experiment_name)
        return self.experiments[experiment_name]

    def get_experiment(self, experiment_name):
        return self.experiments.get(experiment_name)


class Course:
    def __init__(self, name):
        self.name = name
        self.classes = {}  # 班级名 -> Class对象

    def add_class(self, class_name):
        if class_name not in self.classes:
            self.classes[class_name] = Class(class_name)
        return self.classes[class_name]

    def get_class(self, class_name):
        return self.classes.get(class_name)


class DirectoryParser:
    def __init__(self, student_manager):
        self.student_manager = student_manager
        self.courses = {}  # 课程名 -> Course对象
        self.logger = Logger()

    def parse_directory(self, root_path):
        self.courses = {}
        if not os.path.exists(root_path):
            self.logger.log(f"错误：目录不存在 - {root_path}")
            return False

        # 遍历目录结构
        for course_name in os.listdir(root_path):
            course_path = os.path.join(root_path, course_name)
            if not os.path.isdir(course_path):
                continue

            course = self.add_course(course_name)

            for class_name in os.listdir(course_path):
                class_path = os.path.join(course_path, class_name)
                if not os.path.isdir(class_path):
                    continue

                class_obj = course.add_class(class_name)

                for experiment_name in os.listdir(class_path):
                    experiment_path = os.path.join(class_path, experiment_name)
                    if not os.path.isdir(experiment_path):
                        continue

                    experiment = class_obj.add_experiment(experiment_name)

                    # 解析实验目录中的文件
                    self._parse_experiment_files(experiment_path, course_name, class_name, experiment)

        self._update_missing_experiments()
        return True

    def add_course(self, course_name):
        if course_name not in self.courses:
            self.courses[course_name] = Course(course_name)
        return self.courses[course_name]

    def _parse_experiment_files(self, experiment_path, course_name, class_name, experiment):
        file_pattern = re.compile(r'实验(\d+)_(\d+)-(\w+)\.(doc|docx|pdf|txt)')

        for filename in os.listdir(experiment_path):
            file_path = os.path.join(experiment_path, filename)
            if not os.path.isfile(file_path):
                continue

            match = file_pattern.match(filename)
            if not match:
                self.logger.log(f"文件名格式错误: {filename}")
                continue

            experiment_num = match.group(1)
            student_id = match.group(2)
            student_name = match.group(3)

            # 验证学生是否存在
            student = self.student_manager.get_student(student_id)
            if not student:
                self.logger.log(f"学生不在名单中: {student_name}({student_id})")
                continue

            # 检查学生姓名是否匹配
            if student.name != student_name:
                self.logger.log(f"学生姓名不匹配: 文件中为{student_name}，名单中为{student.name}({student_id})")

            experiment.add_submitted_student(student_id)

    def _update_missing_experiments(self):
        # 遍历所有课程-班级-实验，更新学生的缺交实验
        for course_name, course in self.courses.items():
            for class_name, class_obj in course.classes.items():
                for experiment_name, experiment in class_obj.experiments.items():
                    missing_students = experiment.get_missing_students(
                        self.student_manager.get_students_by_class(class_name))

                    for student_id in missing_students:
                        student = self.student_manager.get_student(student_id)
                        if student:
                            student.add_missing_experiment(experiment_name)

    def get_course_names(self):
        return list(self.courses.keys())

    def get_class_names(self, course_name):
        course = self.courses.get(course_name)
        if course:
            return list(course.classes.keys())
        return []

    def get_student_stats(self, course_name, class_name):
        course = self.courses.get(course_name)
        if not course:
            return []

        class_obj = course.get_class(class_name)
        if not class_obj:
            return []

        students = self.student_manager.get_students_by_class(class_name)
        stats = []

        for student in students:
            missing_count = len(student.missing_experiments)
            missing_list = ", ".join(student.missing_experiments)
            stats.append({
                'student_id': student.student_id,
                'name': student.name,
                'grade': student.grade,  # 新增年级
                'class_name': student.class_name,  # 新增班级
                'missing_count': missing_count,
                'missing_list': missing_list
            })

        return stats

    def get_experiment_stats(self, course_name, class_name):
        course = self.courses.get(course_name)
        if not course:
            return []

        class_obj = course.get_class(class_name)
        if not class_obj:
            return []

        total_students = len(self.student_manager.get_students_by_class(class_name))
        stats = []

        for experiment_name, experiment in class_obj.experiments.items():
            missing_students = experiment.get_missing_students(
                self.student_manager.get_students_by_class(class_name))
            missing_names = []

            for student_id in missing_students:
                student = self.student_manager.get_student(student_id)
                if student:
                    missing_names.append(f"{student.name}({student_id})")

            submission_rate = experiment.get_submission_rate(total_students)

            stats.append({
                'experiment_name': experiment_name,
                'submission_rate': submission_rate,
                'missing_students': ", ".join(missing_names)
            })

        return stats

    def get_submission_rates(self, course_name, class_name):
        course = self.courses.get(course_name)
        if not course:
            return [], []

        class_obj = course.get_class(class_name)
        if not class_obj:
            return [], []

        experiments = sorted(class_obj.experiments.values(),
                             key=lambda exp: int(re.search(r'\d+', exp.name).group()) if re.search(r'\d+',
                                                                                                   exp.name) else 0)

        names = [exp.name for exp in experiments]
        rates = [exp.get_submission_rate(len(self.student_manager.get_students_by_class(class_name)))
                 for exp in experiments]

        return names, rates


class StudentManager:
    def __init__(self):
        self.students = {}  # 学号 -> Student对象
        self.classes = {}  # 班级名 -> [Student对象]
        self.logger = Logger()

    def add_student(self, student_id, name, grade="", class_name=""):
        if student_id in self.students:
            self.logger.log(f"警告：重复添加学生 {name}({student_id})")
            return

        student = Student(student_id, name, grade, class_name)  # 传递年级和班级
        self.students[student_id] = student

        if class_name:
            if class_name not in self.classes:
                self.classes[class_name] = []
            self.classes[class_name].append(student)

    def import_from_excel(self, file_path):
        try:
            df = pd.read_excel(file_path)
            # 假设Excel包含学号、姓名、年级和班级四列
            for index, row in df.iterrows():
                student_id = str(row.get('学号', '')).strip()
                name = str(row.get('姓名', '')).strip()
                grade = str(row.get('年级', '')).strip()  # 新增年级解析
                class_name = str(row.get('班级', '')).strip()

                if student_id and name:
                    self.add_student(student_id, name, grade, class_name)

            self.logger.log(f"成功从Excel导入 {len(df)} 名学生")
            return True
        except Exception as e:
            self.logger.log(f"导入Excel失败: {str(e)}")
            return False

    def get_student(self, student_id):
        return self.students.get(student_id)

    def get_students_by_class(self, class_name):
        return self.classes.get(class_name, [])

    def get_all_students(self):
        return list(self.students.values())

    def clear_students(self):
        self.students = {}
        self.classes = {}


class Logger:
    _instance = None

    def __new__(cls):
        if not cls._instance:
            cls._instance = super().__new__(cls)
            cls._instance.logs = []
        return cls._instance

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        full_message = f"[{timestamp}] {message}"
        self._instance.logs.append(full_message)
        print(full_message)  # 同时输出到控制台

    def get_logs(self):
        return self._instance.logs

    def clear_logs(self):
        self._instance.logs = []


class StatisticsExporter:
    @staticmethod
    def export_student_stats_to_excel(student_stats, file_path):
        if not student_stats:
            return False

        # 导出包含年级和班级的学生统计数据
        df = pd.DataFrame(student_stats)
        try:
            df.to_excel(file_path, index=False)
            return True
        except Exception as e:
            Logger().log(f"导出学生统计数据失败: {str(e)}")
            return False

    @staticmethod
    def export_experiment_stats_to_excel(experiment_stats, file_path):
        if not experiment_stats:
            return False

        df = pd.DataFrame(experiment_stats)
        try:
            df.to_excel(file_path, index=False)
            return True
        except Exception as e:
            Logger().log(f"导出实验统计数据失败: {str(e)}")
            return False


class Canvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = plt.Figure(figsize=(width, height), dpi=dpi)
        super().__init__(self.fig)
        self.setParent(parent)


class ERATMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.student_manager = StudentManager()
        self.directory_parser = DirectoryParser(self.student_manager)
        self.logger = Logger()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("实验报告统计分析工具 (ERAT)")
        self.setGeometry(100, 100, 1000, 700)

        # 创建主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 创建顶部控制区域
        control_layout = QHBoxLayout()

        # 学生名单导入按钮
        self.import_students_btn = QPushButton("导入学生名单")
        self.import_students_btn.clicked.connect(self.import_students)
        control_layout.addWidget(self.import_students_btn)

        # 选择实验目录按钮
        self.select_dir_btn = QPushButton("选择实验目录")
        self.select_dir_btn.clicked.connect(self.select_directory)
        control_layout.addWidget(self.select_dir_btn)

        # 刷新按钮
        self.refresh_btn = QPushButton("刷新统计")
        self.refresh_btn.clicked.connect(self.refresh_statistics)
        control_layout.addWidget(self.refresh_btn)

        # 添加到主布局
        main_layout.addLayout(control_layout)

        # 创建下拉选择框
        selection_layout = QHBoxLayout()

        self.course_combo = QComboBox()
        self.course_combo.currentTextChanged.connect(self.on_course_changed)
        selection_layout.addWidget(QLabel("选择课程:"))
        selection_layout.addWidget(self.course_combo)

        self.class_combo = QComboBox()
        self.class_combo.currentTextChanged.connect(self.on_class_changed)
        selection_layout.addWidget(QLabel("选择班级:"))
        selection_layout.addWidget(self.class_combo)

        main_layout.addLayout(selection_layout)

        # 创建进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)

        # 创建标签页
        self.tab_widget = QTabWidget()

        # 学生统计标签页
        self.student_tab = QWidget()
        student_layout = QVBoxLayout(self.student_tab)

        self.student_table = QTableWidget()
        self.student_table.setColumnCount(6)  # 修改为6列
        self.student_table.setHorizontalHeaderLabels(["学号", "姓名", "年级", "班级", "缺交次数", "缺交实验列表"])  # 添加年级和班级列
        student_layout.addWidget(self.student_table)

        self.export_student_btn = QPushButton("导出学生统计")
        self.export_student_btn.clicked.connect(self.export_student_stats)
        student_layout.addWidget(self.export_student_btn)

        self.tab_widget.addTab(self.student_tab, "学生统计")

        # 实验统计标签页
        self.experiment_tab = QWidget()
        experiment_layout = QVBoxLayout(self.experiment_tab)

        self.experiment_table = QTableWidget()
        self.experiment_table.setColumnCount(3)
        self.experiment_table.setHorizontalHeaderLabels(["实验名称", "提交率", "未提交学生"])
        experiment_layout.addWidget(self.experiment_table)

        self.export_experiment_btn = QPushButton("导出实验统计")
        self.export_experiment_btn.clicked.connect(self.export_experiment_stats)
        experiment_layout.addWidget(self.export_experiment_btn)

        self.tab_widget.addTab(self.experiment_tab, "实验统计")

        # 可视化标签页
        self.visualization_tab = QWidget()
        visualization_layout = QVBoxLayout(self.visualization_tab)

        self.canvas = Canvas(self.visualization_tab, width=7, height=4)
        visualization_layout.addWidget(self.canvas)

        self.tab_widget.addTab(self.visualization_tab, "提交率可视化")

        # 日志标签页
        self.log_tab = QWidget()
        log_layout = QVBoxLayout(self.log_tab)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)

        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_logs)
        log_layout.addWidget(self.clear_log_btn)

        self.tab_widget.addTab(self.log_tab, "操作日志")

        main_layout.addWidget(self.tab_widget)

        # 状态栏
        self.statusBar().showMessage("就绪")

        # 初始禁用按钮
        self.select_dir_btn.setEnabled(False)
        self.refresh_btn.setEnabled(False)
        self.course_combo.setEnabled(False)
        self.class_combo.setEnabled(False)
        self.export_student_btn.setEnabled(False)
        self.export_experiment_btn.setEnabled(False)

    def import_students(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择学生名单文件", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            self.progress_bar.show()
            self.progress_bar.setValue(0)
            self.statusBar().showMessage("正在导入学生名单...")
            QApplication.processEvents()

            success = self.student_manager.import_from_excel(file_path)

            self.progress_bar.setValue(100)
            self.progress_bar.hide()

            if success:
                self.statusBar().showMessage(f"成功导入 {len(self.student_manager.get_all_students())} 名学生")
                self.select_dir_btn.setEnabled(True)
                QMessageBox.information(self, "成功", "学生名单导入成功！")
                self.update_logs()
            else:
                self.statusBar().showMessage("学生名单导入失败")
                QMessageBox.critical(self, "错误", "学生名单导入失败，请检查文件格式！")

    def select_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择实验报告目录")

        if dir_path:
            self.progress_bar.show()
            self.progress_bar.setValue(0)
            self.statusBar().showMessage("正在解析目录...")
            QApplication.processEvents()

            success = self.directory_parser.parse_directory(dir_path)

            self.progress_bar.setValue(100)
            self.progress_bar.hide()

            if success:
                self.statusBar().showMessage("目录解析完成")
                self.refresh_btn.setEnabled(True)
                self.course_combo.setEnabled(True)

                # 更新课程下拉框
                self.course_combo.clear()
                self.course_combo.addItems(self.directory_parser.get_course_names())

                if self.directory_parser.get_course_names():
                    self.course_combo.setCurrentIndex(0)

                QMessageBox.information(self, "成功", "目录解析成功！")
                self.update_logs()
            else:
                self.statusBar().showMessage("目录解析失败")
                QMessageBox.critical(self, "错误", "目录解析失败，请检查目录结构！")

    def on_course_changed(self, course_name):
        self.class_combo.clear()
        self.class_combo.addItems(self.directory_parser.get_class_names(course_name))

        if self.directory_parser.get_class_names(course_name):
            self.class_combo.setCurrentIndex(0)
            self.class_combo.setEnabled(True)
        else:
            self.class_combo.setEnabled(False)

    def on_class_changed(self, class_name):
        course_name = self.course_combo.currentText()
        if course_name and class_name:
            self.update_student_stats(course_name, class_name)
            self.update_experiment_stats(course_name, class_name)
            self.update_visualization(course_name, class_name)

            self.export_student_btn.setEnabled(True)
            self.export_experiment_btn.setEnabled(True)

    def refresh_statistics(self):
        course_name = self.course_combo.currentText()
        class_name = self.class_combo.currentText()

        if course_name and class_name:
            self.update_student_stats(course_name, class_name)
            self.update_experiment_stats(course_name, class_name)
            self.update_visualization(course_name, class_name)

            self.statusBar().showMessage("统计数据已刷新")

    def update_student_stats(self, course_name, class_name):
        stats = self.directory_parser.get_student_stats(course_name, class_name)

        self.student_table.setRowCount(len(stats))

        for row, stat in enumerate(stats):
            self.student_table.setItem(row, 0, QTableWidgetItem(stat['student_id']))
            self.student_table.setItem(row, 1, QTableWidgetItem(stat['name']))
            self.student_table.setItem(row, 2, QTableWidgetItem(stat['grade']))  # 添加年级列
            self.student_table.setItem(row, 3, QTableWidgetItem(stat['class_name']))  # 添加班级列
            self.student_table.setItem(row, 4, QTableWidgetItem(str(stat['missing_count'])))
            self.student_table.setItem(row, 5, QTableWidgetItem(stat['missing_list']))

        self.student_table.resizeColumnsToContents()

    def update_experiment_stats(self, course_name, class_name):
        stats = self.directory_parser.get_experiment_stats(course_name, class_name)

        self.experiment_table.setRowCount(len(stats))

        for row, stat in enumerate(stats):
            self.experiment_table.setItem(row, 0, QTableWidgetItem(stat['experiment_name']))

            rate_item = QTableWidgetItem(f"{stat['submission_rate']:.2f}%")
            # 根据提交率设置不同的颜色
            if stat['submission_rate'] < 60:
                rate_item.setBackground(Qt.red)
            elif stat['submission_rate'] < 80:
                rate_item.setBackground(Qt.yellow)
            else:
                rate_item.setBackground(Qt.green)
            self.experiment_table.setItem(row, 1, rate_item)

            self.experiment_table.setItem(row, 2, QTableWidgetItem(stat['missing_students']))

        self.experiment_table.resizeColumnsToContents()

    def update_visualization(self, course_name, class_name):
        names, rates = self.directory_parser.get_submission_rates(course_name, class_name)

        self.canvas.fig.clear()
        ax = self.canvas.fig.add_subplot(111)

        if names and rates:
            ax.bar(names, rates, color='skyblue')
            ax.set_title('实验提交率统计')
            ax.set_xlabel('实验名称')
            ax.set_ylabel('提交率 (%)')
            ax.set_ylim(0, 105)

            # 添加数值标签
            for i, rate in enumerate(rates):
                ax.text(i, rate + 1, f"{rate:.1f}%", ha='center')

            self.canvas.fig.tight_layout()
            self.canvas.draw()
        else:
            ax.text(0.5, 0.5, '暂无数据', ha='center', va='center', transform=ax.transAxes)
            self.canvas.draw()

    def export_student_stats(self):
        course_name = self.course_combo.currentText()
        class_name = self.class_combo.currentText()

        if not (course_name and class_name):
            QMessageBox.warning(self, "警告", "请先选择课程和班级！")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出学生统计数据", f"{course_name}_{class_name}_学生统计.xlsx", "Excel Files (*.xlsx)"
        )

        if file_path:
            stats = self.directory_parser.get_student_stats(course_name, class_name)
            success = StatisticsExporter.export_student_stats_to_excel(stats, file_path)

            if success:
                self.statusBar().showMessage(f"学生统计数据已导出到 {file_path}")
                QMessageBox.information(self, "成功", "学生统计数据导出成功！")
                self.logger.log(f"导出学生统计数据到 {file_path}")
                self.update_logs()
            else:
                self.statusBar().showMessage("学生统计数据导出失败")
                QMessageBox.critical(self, "错误", "学生统计数据导出失败！")

    def export_experiment_stats(self):
        course_name = self.course_combo.currentText()
        class_name = self.class_combo.currentText()

        if not (course_name and class_name):
            QMessageBox.warning(self, "警告", "请先选择课程和班级！")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出实验统计数据", f"{course_name}_{class_name}_实验统计.xlsx", "Excel Files (*.xlsx)"
        )

        if file_path:
            stats = self.directory_parser.get_experiment_stats(course_name, class_name)
            success = StatisticsExporter.export_experiment_stats_to_excel(stats, file_path)

            if success:
                self.statusBar().showMessage(f"实验统计数据已导出到 {file_path}")
                QMessageBox.information(self, "成功", "实验统计数据导出成功！")
                self.logger.log(f"导出实验统计数据到 {file_path}")
                self.update_logs()
            else:
                self.statusBar().showMessage("实验统计数据导出失败")
                QMessageBox.critical(self, "错误", "实验统计数据导出失败！")

    def update_logs(self):
        self.log_text.clear()
        for log in self.logger.get_logs():
            self.log_text.append(log)

    def clear_logs(self):
        self.logger.clear_logs()
        self.log_text.clear()
        self.statusBar().showMessage("日志已清空")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ERATMainWindow()
    window.show()
    sys.exit(app.exec_())
