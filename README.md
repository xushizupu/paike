# 排课系统

读取同目录下的 `基本数据.xlsx`，根据“各班任课、各班课时、作息表”三张表自动排课，并支持：

- 班级不排课
- 固定课
- 课程时段约束
- 教师不排课
- 同科同天最多节数、连堂最多节数设置
- 班级课表拖动微调
- 导出班级课表、教师课表、校验报告
- 上传自定义数据、下载默认数据模板
- 自定义班级课表和教师课表表头
- 多人会话隔离，上传数据和设置互不干扰
- 拖动调课冲突校验与确认

## 启动

```bash
pip install -r requirements.txt
python app.py
```

启动后打开 `http://127.0.0.1:8000`。

Windows 下也可以直接双击 `启动排课系统.bat`。

## 输出

导出的 Excel 文件保存在 `output` 目录，每次导出三个文件：

- `班级课程表_日期_序号.xlsx`：每个班一个 sheet
- `教师课程表_日期_序号.xlsx`：每位教师一个 sheet
- `排课校验_日期_序号.xlsx`：包含校验报告和课时统计

排课设置会保存到当前会话目录的 `settings.json`，下次启动自动恢复。

每个用户的数据会保存在独立的 `sessions` 目录，会话空闲超过 1 小时会自动清理。

页面下方的使用步骤说明保存在 `页面使用说明.json`，可直接编辑。

## 部署到 Render

项目已包含 `Dockerfile`，Render 会自动识别并构建。

1. 把项目文件上传到 GitHub 仓库，必须保留 `static/` 文件夹结构。
2. 在 Render 控制台点击 `New`，选择 `Web Service`。
3. 连接 GitHub，选择这个仓库。
4. Render 会自动识别 Dockerfile；实例类型选择 Free。
5. 点击部署，等待构建完成。
6. 打开 Render 提供的 `https://你的服务名.onrender.com`。

Render 会自动设置 `PORT` 环境变量，应用监听 `0.0.0.0`。免费实例会休眠，首次访问可能需要等待冷启动。

防止休眠：可用 UptimeRobot 或 Cron-Job.org 每 5 分钟访问一次 `https://你的服务名.onrender.com/api/ping`。
