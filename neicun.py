import psutil
import datetime
import openpyxl
import time

def monitor():
    # 获取当前时间
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 获取所有进程的信息
    processes = psutil.process_iter()
    # 记录所有进程的 CPU 和内存使用量
    data = []
    for process in processes:
        try:
            # 获取进程名称
            name = process.name()
            # 获取进程占用的 CPU 和内存使用量
            cpu_percent = process.cpu_percent()
            memory_info = process.memory_info().rss / 1024 / 1024
            memory_percent = process.memory_percent()
            data.append([current_time, name, cpu_percent, memory_info, memory_percent])
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    # 将数据写入 Excel 文件
    filename = "C:\\computer memory-" + datetime.datetime.now().strftime("%Y-%m-%d") + ".xlsx"
    try:
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(["Time", "Process Name", "CPU Usage", "Memory Usage (MB)", "Memory Usage (%)"])
    for item in data:
        worksheet.append(item)
    workbook.save(filename)

if __name__ == "__main__":
    while True:
        monitor()
        time.sleep(60)
