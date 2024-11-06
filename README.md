# CSU-Course-Schedule-2-ICS
将中南大学教务系统下载的 xls 格式的课程表解析并生成 iCalendar 文件（.ics），以便导入日历。使用第三方的课程表不是很优雅。
# 如何使用
1. 访问教务系统，点击“打印”下载课表（.xls格式）；

![](./captures/1.png)

2. 在同一页面找到本学期的第一天的日期；

![](./captures/2.png)

3. 克隆本仓库，安装所需模块：

```bash
git clone 

cd CSU-Course-Schedule-2-ICS

pip install -r requirements.txt
```

4. 修改 `create_ics.py` 相关内容，运行得到 `courses.ics` 文件，例如：

```python
# 假设学期开始日期为 2024 年 9 月 1 日
term_start = [2024, 9, 1]
courses = parse_courses(term_start, '学生个人课表_8211211224.xls')
courses_to_ics(courses)
```

5. 在日历导入刚刚获得的文件即可，不同日历软件各有差异，自己摸索。

# 效果展示
![](./captures/video.mp4)