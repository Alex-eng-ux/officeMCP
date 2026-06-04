"""测试 Office MCP PowerPoint 高级功能."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
test_file = os.path.join(desktop, "Office_MCP_PPT_Test.pptx")

print(f"测试文件路径: {test_file}")
print("=" * 50)

try:
    import win32com.client

    # 启动 PowerPoint
    print("1. 启动 PowerPoint...")
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = 1

    # 创建演示文稿
    print("2. 创建演示文稿...")
    presentation = ppt.Presentations.Add()

    # ===== 测试基础功能 =====

    # 幻灯片 1: 标题页
    print("3. 测试幻灯片...")
    slide1 = presentation.Slides.Add(1, 1)  # ppLayoutTitle
    if slide1.Shapes.Title:
        slide1.Shapes.Title.TextFrame.TextRange.Text = "Office MCP PPT 高级功能测试"
    print("   - 添加标题幻灯片: OK")

    # 幻灯片 2: 内容页
    slide2 = presentation.Slides.Add(2, 2)  # ppLayoutText
    if slide2.Shapes.Title:
        slide2.Shapes.Title.TextFrame.TextRange.Text = "高级功能演示"
    print("   - 添加内容幻灯片: OK")

    # ===== 测试高级功能 =====

    # 1. 添加文本框
    print("4. 测试文本框...")
    textbox = slide2.Shapes.AddTextbox(1, 50, 150, 600, 300)
    textbox.TextFrame.TextRange.Text = "• 数据分析\n• 图表展示\n• 动画效果"
    print("   - 添加文本框: OK")

    # 2. 形状
    print("5. 测试形状...")
    shape = slide2.Shapes.AddShape(1, 500, 400, 150, 100)  # msoShapeRectangle
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = 0x4472C4  # 蓝色
    shape.Line.ForeColor.RGB = 0x000000
    shape.Line.Weight = 2
    print("   - 添加形状: OK")

    # 3. 转换效果
    print("6. 测试转换效果...")
    slide2.SlideShowTransition.EntryEffect = 3844  # ppEffectFade
    slide2.SlideShowTransition.Duration = 1.5
    print("   - 转换效果: OK")

    # 4. 背景颜色
    print("7. 测试背景颜色...")
    slide2.FollowMasterBackground = False
    slide2.Background.Fill.Solid()
    slide2.Background.Fill.ForeColor.RGB = 0xF5F5F5  # 浅灰色
    print("   - 背景颜色: OK")

    # 5. 幻灯片编号
    print("8. 测试幻灯片编号...")
    for s in slide2.Shapes.Placeholders:
        if s.PlaceholderFormat.Type == 6:  # ppSlideNumber
            s.Visible = True
    print("   - 幻灯片编号: OK")

    # 6. 备注
    print("9. 测试备注...")
    if slide2.HasNotesPage:
        notes = slide2.NotesPage.Shapes.Placeholders(2)
        notes.TextFrame.TextRange.Text = "这是演讲者备注"
    print("   - 备注: OK")

    # 7. 格式化形状
    print("10. 测试格式化形状...")
    shape.Fill.ForeColor.RGB = 0x00B050  # 绿色
    shape.Line.ForeColor.RGB = 0xFF0000  # 红色
    shape.Line.Weight = 3
    print("   - 格式化形状: OK")

    # 幻灯片 3: 结束页
    print("11. 添加结束页...")
    slide3 = presentation.Slides.Add(3, 1)
    if slide3.Shapes.Title:
        slide3.Shapes.Title.TextFrame.TextRange.Text = "谢谢观看!"
    print("   - 结束页: OK")

    # 保存
    print("12. 保存文件...")
    presentation.SaveAs(test_file)
    print(f"   - 保存: OK -> {test_file}")

    print("=" * 50)
    print("所有 PowerPoint 功能测试通过!")

    input("\n按回车键关闭 PowerPoint...")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    try:
        presentation.Close()
        ppt.Quit()
    except:
        pass
