# PPT备注提取工具

一个简单易用的在线工具，用于从PowerPoint文件中提取备注内容并转换为Markdown格式。

## 功能特点

- 支持.pptx格式的PowerPoint文件
- 提取每张幻灯片的备注内容
- 转换为易读的Markdown格式
- 支持在线使用，无需安装
- 支持文件拖放上传
- 一键复制提取结果

## 本地部署

### 环境要求

- Python 3.6+
- pip（Python包管理器）

### 安装步骤

1. 克隆仓库到本地：
```bash
git clone https://github.com/BruceLee1024/ppt-notes-extractor.git
cd ppt-notes-extractor
```

2. 安装依赖包：
```bash
pip install -r requirements.txt
```

3. 运行应用：
```bash
python app.py
```

4. 在浏览器中访问：http://localhost:5000

### 命令行使用

也可以直接使用命令行工具提取PPT备注：

```bash
python extract_notes.py 你的PPT文件.pptx
```

可选参数：
- `-o, --output`: 指定输出的Markdown文件路径

## 在线使用

访问：[PPT备注提取工具](https://BruceLee1024.github.io/ppt-notes-extractor)

1. 点击上传区域或将PPT文件拖放到页面中
2. 等待处理完成
3. 查看转换结果
4. 点击"复制内容"按钮复制结果

## 贡献

欢迎提交Issue和Pull Request！

## 许可证

MIT License