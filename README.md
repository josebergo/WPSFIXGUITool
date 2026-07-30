# WPSFIXGUITool

Windows 桌面工具，用于检查并修复 DOCX 中不兼容 WPS 的页码域和外链图片。

- 版本：`V1.0.1`
- 作者：`鼎泰高科全球信息部`

版本和作者同时写入程序界面及 EXE 的 Windows 文件属性。

## 界面

程序只有两个操作按钮：

- **选择文件**：选择一个 `.docx` 文件。
- **检查并转换**：检查并生成新的 `-WPS兼容版.docx`。

转换过程中会显示进度条、当前阶段和结果。源文件不会被覆盖。

## 修复内容

- 将“页次/页码”相邻单元格中的损坏或浮动页码对象重建为标准 `PAGE / NUMPAGES` 域。
- 将有缓存结果的 `INCLUDEPICTURE` 外链图片转为真正的内嵌图片。
- 将内嵌 TIFF/BMP 转为 PNG，并保持页面尺寸关系。
- 设置打开文档时更新字段；电脑安装 Word 时同时刷新字段缓存。
- 校验 DOCX ZIP、XML、字段配对和图片关系。
- 在输出文件旁生成 `.report.json` 检查报告。

## 运行源码

```powershell
.\run.ps1
```

或：

```powershell
python app.py
```

## 测试

```powershell
python -m unittest discover -s tests -v
```

## 打包 EXE

```powershell
.\build.ps1
```

生成文件：`dist\WPSFIXGUITool.exe`。

## 安全边界

- 不覆盖输入文件。
- 无缓存数据的失效远程图片不会被空白图替代，而是在报告中提示。
- 转换先写临时文件并完成 CRC 检查，再原子生成最终 DOCX。
