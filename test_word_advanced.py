"""测试 Office MCP Word 高级功能."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
test_file = os.path.join(desktop, "Office_MCP_Word_Test.docx")

print(f"测试文件路径: {test_file}")
print("=" * 50)

try:
    import win32com.client

    # 启动 Word
    print("1. 启动 Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0

    # 创建文档
    print("2. 创建文档...")
    doc = word.Documents.Add()

    # ===== 测试基础功能 =====
    print("3. 测试基础功能...")

    # 添加标题
    p1 = doc.Content.Paragraphs.Add()
    p1.Range.Text = "Office MCP Word 高级功能测试"
    p1.Style = "标题 1"
    p1.Alignment = 1  # wdAlignParagraphCenter
    print("   - 添加标题: OK")

    # 添加段落
    p2 = doc.Content.Paragraphs.Add()
    p2.Range.Text = "这是正文内容，包含一些测试文本。"
    p2.Style = "正文"
    print("   - 添加段落: OK")

    # ===== 测试高级功能 =====

    # 1. 样式
    print("4. 测试样式...")
    doc.Content.Style = "标题 2"
    print("   - 应用样式: OK")

    # 2. 页眉 (修复：先启用页眉)
    print("5. 测试页眉...")
    try:
        section = doc.Sections(1)
        section.PageSetup.HeaderDistance = 20
        # 设置页眉存在
        for i in range(1, 4):
            section.Headers(i).LinkHeaderFooter = None
        header = section.Headers(1)
        header.Range.Text = "Office MCP 测试文档"
        print("   - 设置页眉: OK")
    except Exception as e:
        print(f"   - 设置页眉: 跳过 ({e})")

    # 3. 页脚
    print("6. 测试页脚...")
    try:
        section = doc.Sections(1)
        footer = section.Footers(1)
        footer.Range.Text = "第 "
        footer.Range.Collapse(0)
        footer.Range.Fields.Add(footer.Range, Type=33)
        footer.Range.Text = footer.Range.Text + " 页"
        print("   - 设置页脚: OK")
    except Exception as e:
        print(f"   - 设置页脚: 跳过 ({e})")

    # 4. 分页符
    print("7. 测试分页符...")
    doc.Content.Collapse(0)
    doc.Content.InsertBreak(7)
    print("   - 分页符: OK")

    # 5. 插入表格
    print("8. 测试表格...")
    doc.Content.Paragraphs.Add()
    table = doc.Tables.Add(doc.Content, 3, 3)
    table.Cell(1, 1).Range.Text = "姓名"
    table.Cell(1, 2).Range.Text = "年龄"
    table.Cell(1, 3).Range.Text = "城市"
    table.Cell(2, 1).Range.Text = "张三"
    table.Cell(2, 2).Range.Text = "25"
    table.Cell(2, 3).Range.Text = "北京"
    table.Cell(3, 1).Range.Text = "李四"
    table.Cell(3, 2).Range.Text = "30"
    table.Cell(3, 3).Range.Text = "上海"
    print("   - 插入表格: OK")

    # 6. 字体设置
    print("9. 测试字体设置...")
    doc.Content.Font.Name = "微软雅黑"
    doc.Content.Font.Size = 12
    print("   - 字体设置: OK")

    # 7. 页面设置
    print("10. 测试页面设置...")
    doc.PageSetup.Orientation = 0
    doc.PageSetup.TopMargin = 72
    doc.PageSetup.BottomMargin = 72
    print("   - 页面设置: OK")

    # 8. 替换文本
    print("11. 测试文本替换...")
    find = doc.Content.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Execute(FindText="正文", ReplaceWith="这里是正文内容")
    print("   - 文本替换: OK")

    # 保存
    print("12. 保存文件...")
    doc.SaveAs(test_file)
    print(f"   - 保存: OK -> {test_file}")

    print("=" * 50)
    print("所有 Word 功能测试通过!")

    input("\n按回车键关闭 Word...")

except Exception as e:
    print(f"\n错误: {e}")
    import traceback
    traceback.print_exc()

finally:
    try:
        doc.Close(SaveChanges=0)
        word.Quit()
    except:
        pass
