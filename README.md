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

## 双许可（Dual Licensing）

本项目采用双许可模式：

- **AGPL-3.0**：适用于开源分发、修改和网络服务场景
- **商业授权**：如需闭源集成、商用分发或不希望受 AGPL 约束，请联系版权持有人获取商业授权


