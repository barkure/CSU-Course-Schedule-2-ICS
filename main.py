import pandas as pd
from datetime import datetime, timedelta
from ics import Calendar, Event
from ics.grammar.parse import ContentLine


class Course:
    def __init__(self, course_name, teacher, location, start_time, end_time):
        self.course_name = course_name
        self.teacher = teacher
        self.location = location
        self.start_time = start_time
        self.end_time = end_time

    def __str__(self):
        return "课程名称：{}，教师：{}，地点：{}，开始：{}, 结束：{}".format(
            self.course_name, self.teacher, self.location, self.start_time, self.end_time)


# 解析周次字符串，返回周次列表
def parse_weeks(weeks_str):
    weeks = []
    if '[周]' in weeks_str:
        weeks_str = weeks_str.replace('[周]', '')
        for week in weeks_str.split(','):
            if '-' in week:
                start, end = map(int, week.split('-'))
                weeks.extend(range(start, end + 1))
            else:
                weeks.append(int(week))
    elif '[单周]' in weeks_str:
        weeks_str = weeks_str.replace('[单周]', '')
        for week in weeks_str.split(','):
            if '-' in week:
                start, end = map(int, week.split('-'))
                # 确保从单数开始
                if start % 2 == 0:
                    start += 1
                weeks.extend(range(start, end + 1, 2))
            else:
                if int(week) % 2 != 0:
                    weeks.append(int(week))
    elif '[双周]' in weeks_str:
        weeks_str = weeks_str.replace('[双周]', '')
        for week in weeks_str.split(','):
            if '-' in week:
                start, end = map(int, week.split('-'))
                # 确保从双数开始
                if start % 2 != 0:
                    start += 1
                weeks.extend(range(start, end + 1, 2))
            else:
                if int(week) % 2 == 0:
                    weeks.append(int(week))
    return weeks


# 读取XLS文件，解析课程数据
def parse_courses(term_start, file_path):
    term_start_time = datetime(term_start[0], term_start[1], term_start[2])
    # 计算 term_start 所在周的周日
    start_week_sunday = term_start_time + timedelta(days=(6 - term_start_time.weekday()))
    # 读取课程表文件
    xls = pd.read_excel(file_path, header=None)
    # 从第4行第2列开始，到第9行第8列结束，即只读取课程表中的课程数据
    course_data = xls.iloc[3:9, 1:8]

    courses = {}
    course_id = 0

    # 定义每天的课程时间
    class_hours = [8, 10, 14, 16, 19]

    # 遍历课程表中的每个格子
    for j in range(len(course_data.columns)):
        for i in range(len(course_data)):
            cell = course_data.iloc[i, j]
            lines = [line for line in cell.split('\n') if line.strip()]
            '''
            这里会得到三种数据，例如：
            1. 没有课：[]
            2. 一门课：['形势与政策', '刘建亚(教授)', '7,13[周]', 'A座317']
            3. 两门课：['地磁场与磁法勘探', '鲁光银(教授)', '1-8[周]', 'A座224', '重力场与重力勘探', '肖晓(副教授)', '9-16[周]', 'B座418']
            可能一个格子中有更多的课程，但列表长度都应该是4的倍数
            '''
            for k in range(0, len(lines), 4):
                course_name = lines[k]
                teacher = lines[k + 1]
                weeks_str = lines[k + 2]
                location = lines[k + 3]
                weeks = parse_weeks(weeks_str)
                for week in weeks:
                    # 计算课程的日期和时间
                    start = start_week_sunday + timedelta(weeks=week - 1, days=j, hours=class_hours[i])
                    end = start + timedelta(hours=1, minutes=40)
                    course = Course(course_name, teacher, location, start, end)
                    courses[course_id] = course
                    course_id += 1
        
    return courses


# 保存为日历文件
def courses_to_ics(courses):
    calendar = Calendar()
    calendar.extra.append(ContentLine(name="X-WR-CALNAME", value="课程表"))
    for course in courses.values():
        event = Event()
        event.name = course.course_name
        event.description = course.teacher
        event.location = course.location
        event.extra.append(ContentLine(name="DTSTART;TZID=Asia/Shanghai", value=course.start_time.strftime('%Y%m%dT%H%M%S')))
        event.extra.append(ContentLine(name="DTEND;TZID=Asia/Shanghai", value=course.end_time.strftime('%Y%m%dT%H%M%S')))
        calendar.events.add(event)
    with open('courses.ics', 'w', encoding='utf-8') as f:
        f.writelines(calendar)

