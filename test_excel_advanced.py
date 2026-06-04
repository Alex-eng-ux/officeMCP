"""测试 Office MCP Excel 高级功能."""
import os
import sys

# 添加 src 目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
test_file = os.path.join(desktop, "Office_MCP_Excel_Test.xlsx")

print(f"测试文件路径: {test_file}")
print("=" * 50)

try:
    import win32com.client

    # 启动 Excel
    print("1. 启动 Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    # 创建工作簿
    print("2. 创建工作簿...")
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "测试"

    # ===== 测试基础功能 =====
    print("3. 测试基础功能...")

    # 写入数据
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "分数"
    ws.Range("A2").Value = "张三"
    ws.Range("B2").Value = 85
    ws.Range("A3").Value = "李四"
    ws.Range("B3").Value = 92
    ws.Range("A4").Value = "王五"
    ws.Range("B4").Value = 78
    print("   - 写入数据: OK")

    # 设置格式
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = 0x4472C4  # 蓝色
    ws.Range("A1:B1").Font.Color = 0xFFFFFF  # 白色
    print("   - 设置格式: OK")

    # ===== 测试高级功能 =====

    # 1. 合并单元格
    print("4. 测试合并单元格...")
    ws.Range("D1:E1").Merge()
    ws.Range("D1").Value = "汇总信息"
    ws.Range("D1").HorizontalAlignment = -4108  # xlCenter
    print("   - 合并单元格: OK")

    # 2. 边框
    print("5. 测试边框...")
    rng = ws.Range("A1:D4")
    rng.Borders(7).LineStyle = 1  # xlEdgeTop
    rng.Borders(8).LineStyle = 1  # xlEdgeBottom
    rng.Borders(9).LineStyle = 1  # xlEdgeLeft
    rng.Borders(10).LineStyle = 1  # xlEdgeRight
    rng.Borders(11).LineStyle = 1  # xlInsideHorizontal
    rng.Borders(12).LineStyle = 1  # xlInsideVertical
    print("   - 边框: OK")

    # 3. 自动调整列宽
    print("6. 测试自动调整列宽...")
    ws.Columns("A:D").AutoFit()
    print("   - 自动调整列宽: OK")

    # 4. 冻结窗格
    print("7. 测试冻结窗格...")
    ws.Range("A2").Select()
    excel.ActiveWindow.FreezePanes = True
    print("   - 冻结窗格: OK")

    # 5. 数字格式
    print("8. 测试数字格式...")
    ws.Range("B2:B4").NumberFormat = "0.00"
    print("   - 数字格式: OK")

    # 6. 创建图表
    print("9. 测试创建图表...")
    chart = ws.ChartObjects().Add(300, 50, 300, 200).Chart
    chart.ChartType = 51  # xlColumnClustered
    chart.SetSourceData(ws.Range("A1:B4"))
    chart.HasTitle = True
    chart.ChartTitle.Text = "分数图表"
    print("   - 创建图表: OK")

    # 7. 添加工作表
    print("10. 测试添加工作表...")
    ws2 = wb.Worksheets.Add()
    ws2.Name = "数据表"
    ws2.Range("A1").Value = "新工作表数据"
    print("   - 添加工作表: OK")

    # 8. 重命名工作表
    print("11. 测试重命名工作表...")
    ws2.Name = "分析数据"
    print("   - 重命名工作表: OK")

    # 保存
    print("12. 保存文件...")
    wb.SaveAs(test_file)
    print(f"   - 保存: OK -> {test_file}")

    print("=" * 50)
    print("所有 Excel 功能测试通过!")

    input("\n按回车键关闭 Excel...")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    try:
        wb.Close(SaveChanges=False)
        excel.Quit()
    except:
        pass
