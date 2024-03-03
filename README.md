



# python自动化脚本提示pdf文件信息，生成excel文件

## 项目描述或简介。
- 用pyqt5制作的交互界面。
- 读取文件夹中的每个PDF文件（invoice发票文件）。
- 将单个PDF表单信息的指定字段传输到Excel表格中的一行（Excel列中的标题与PDF表单中的字段匹配）。
- 根据指定的“金额”字段进行所有订单总金额的统计。

## 依融包

-pdfplumber
pdfplumber库是一个用于从PDF文档中提取文本和表格数据的Python库。它可以帮助用户轻松地从PDF文件中提取有用的信息，例如表格、文本、元数据等。pdfplumber库的特点包括：简单易用、速度快、支持多种PDF文件格式、支持从多个页面中提取数据等。pdfplumber库还提供了一些方便的方法来处理提取的数据，例如排序、过滤和格式化等。它是一个非常有用的工具，特别是在需要从大量PDF文件中提取数据时。


-pandas
Pandas是一个强大的分析结构化数据的工具集；它的使用基础是Numpy（提供高性能的矩阵运算）；用于数据挖掘和数据分析，同时也提供数据清洗功能。


-PyQt5
Qt 库是世界上最强大的 GUI 库之一，跨平台，开发语言为 C++(https://www.qt.io). PyQt 是 QT 框架的 Python 语言实现，PyQt5 是用来创建 Python GUI 应用程序的工具包。作为一个跨平台的工具包，PyQt 可以在所有主流操作系统上运行（Unix，Windows，Mac）。


-用pyinstaller打包成exe还需要安装以下几个库：numpy, xlsxwriter, Jinja2, matplotlib


## 使用

1. 运行Pdf2Excel.py文件，出现如下图的界面

![主界面图](https://github.com/wugerr/Pdf2Excel/blob/main/img/WechatIMG166.png?raw=true "主界面图")

2. 点击“选择pdf所在目录”去选择要提取信息的PDF文件所在目录。如下图

![选择pdf所在目录](https://github.com/wugerr/Pdf2Excel/blob/main/img/WechatIMG167.png?raw=true "选择pdf所在目录")

3. 点击“生成excel文件”会自动提取PDF文件到excel中去

被提取的pdf文件数据信息如下图：
![PDF文件](https://github.com/wugerr/Pdf2Excel/blob/main/img/WechatIMG165.png?raw=true "PDF文件")

生成的excel文件如下图：
![Excel文件](https://github.com/wugerr/Pdf2Excel/blob/main/img/WechatIMG164.png?raw=true "Excel文件")

## 联系方式

有关项目的更多信息，请联系我：2700550800@qq.com。
