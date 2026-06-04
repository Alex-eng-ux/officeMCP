"""直接测试 COM 创建 PPT."""
import os

# 设置路径
desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
ppt_path = os.path.join(desktop, "Office_MCP_Test.pptx")

print(f"将创建 PPT 到: {ppt_path}")

try:
    import win32com.client

    # 启动 PowerPoint
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True  # 显示出来方便查看

    # 创建演示文稿
    presentation = ppt.Presentations.Add()

    # 幻灯片 1: 标题页
    slide1 = presentation.Slides.Add(1, 1)  # ppLayoutTitle = 1
    if slide1.Shapes.Title:
        slide1.Shapes.Title.TextFrame.TextRange.Text = "Office MCP 测试演示"
    # 添加副标题
    slide1.Shapes.AddTextbox(
        Orientation=1,  # msoTextOrientationHorizontal
        Left=50, Top=200, Width=600, Height=50
    ).TextFrame.TextRange.Text = "通过 MCP + COM 自动化创建"

    # 幻灯片 2: 内容页
    slide2 = presentation.Slides.Add(2, 2)  # ppLayoutText = 2
    if slide2.Shapes.Title:
        slide2.Shapes.Title.TextFrame.TextRange.Text = "支持的工具"

    # 添加要点
    textbox = slide2.Shapes.AddTextbox(
        Orientation=1,
        Left=50, Top=150, Width=600, Height=300
    )
    tf = textbox.TextFrame
    tf.TextRange.Text = "• Word: 创建文档、段落、表格、图片\n"
    tf.TextRange.InsertAfter("• Excel: 单元格、公式、图表\n")
    tf.TextRange.InsertAfter("• PowerPoint: 幻灯片、文本、图片、表格\n")
    tf.TextRange.InsertAfter("• PDF 导出支持\n")
    tf.TextRange.InsertAfter("• 安全路径控制")

    # 幻灯片 3: 总结页
    slide3 = presentation.Slides.Add(3, 1)
    if slide3.Shapes.Title:
        slide3.Shapes.Title.TextFrame.TextRange.Text = "完成!"

    slide3.Shapes.AddTextbox(
        Orientation=1,
        Left=50, Top=200, Width=600, Height=100
    ).TextFrame.TextRange.Text = "Office MCP Server 已成功创建此演示文稿!"

    # 保存
    presentation.SaveAs(ppt_path)
    print(f"PPT 已保存: {ppt_path}")

    input("\n按回车键关闭 PowerPoint...")
    presentation.Close()
    ppt.Quit()

except Exception as e:
    print(f"错误: {e}")
    import traceback
    traceback.print_exc()
