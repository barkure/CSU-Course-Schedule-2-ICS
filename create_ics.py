from main import parse_courses, courses_to_ics


# 假设学期开始日期为 2024 年 9 月 1 日
term_start = [2024, 9, 1]
courses = parse_courses(term_start, '学生个人课表_8211211224.xls')
courses_to_ics(courses)