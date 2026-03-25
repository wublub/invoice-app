# 发票识别工具

这是一个基于 `invoice_app.py` 的桌面发票识别工具。

## 功能

- 选择一个或多个 PDF 发票文件进行识别
- 提取常用字段：
  - 发票日期
  - 发票号
  - 金额
  - 购买详细
  - 购买方名称
  - 购买方纳税人识别号
  - 销售方名称
  - 销售方纳税人识别号
  - 账号
  - 开户银行
- 在界面中查看识别结果
- 导出 Excel

## 运行环境

- Python 3.10 及以上
- Windows 64 位推荐直接使用 `invoice_app.py` 或打包后的 exe

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动方式

### 方式一：直接运行

```bash
python invoice_app.py
```

### 方式二：使用脚本启动

Windows：

```bash
启动.bat
```

Linux / macOS：

```bash
sh 启动.sh
```

## 使用说明

1. 启动程序
2. 选择 PDF 文件或所在文件夹
3. 点击识别
4. 检查结果
5. 按需导出为 Excel

## 打包

如需打包桌面程序，可使用：

```bash
打包.bat
```

或手动执行：

```bash
pyinstaller "发票识别工具.spec"
```

打包完成后，可执行文件通常位于：

```text
dist/发票识别工具.exe
```

## 上传 GitHub 时建议保留的文件

如果你只想上传和 `invoice_app.py` 相关的内容，建议保留：

- `invoice_app.py`
- `invoice_recognizer.py`
- `requirements.txt`
- `invoice_icon.ico`
- `启动.bat`
- `启动.sh`
- `打包.bat`
- `发票识别工具.spec`
- `README.md`
- `LICENSE`
- `LICENSE-AGPL-3.0.txt`
- `.gitignore`

## 双许可（Dual Licensing）

本项目采用双许可模式：

- **AGPL-3.0**：适用于开源分发、修改和网络服务场景
- **商业授权**：如需闭源集成、商用分发或不希望受 AGPL 约束，请联系版权持有人获取商业授权

### 商业授权联系信息填写位置

你后续只需要修改下面这两处占位文本：

1. 本文件中的这一行：

```text
商业授权联系邮箱：YOUR_EMAIL@example.com
```

2. `LICENSE` 文件中的这一行：

```text
Commercial licensing contact: YOUR_EMAIL@example.com
```

把 `YOUR_EMAIL@example.com` 改成你的真实邮箱即可。

### 当前商业授权联系邮箱

商业授权联系邮箱：`YOUR_EMAIL@example.com`

## License

- 开源许可证：AGPL-3.0，详见 `LICENSE-AGPL-3.0.txt`
- 商业授权说明：详见 `LICENSE`

## 说明

- 建议不要将真实发票、测试 PDF、识别结果 Excel、构建产物上传到公开仓库
- 如果遇到扫描件识别不理想，通常需要额外 OCR 支持
- 打包后的 exe 主要面向 **Windows 64 位** 电脑使用；不同操作系统需要分别打包
