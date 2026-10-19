<div align="center">

# 📋 Attendance Tool

**考勤表整理工具 · 按姓名筛选生成独立考勤表**

<br>

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![WPS](https://img.shields.io/badge/WPS_Office-0078D4?style=for-the-badge&logo=wps&logoColor=white)

<br>

从总考勤表中按姓名列表筛选人员 · 保留原始格式与列宽 · WPS COM 自动化

[功能特性](#功能特性) · [快速启动](#快速启动) · [使用流程](#使用流程) · [项目结构](#项目结构)

</div>

---

## 功能特性

| 功能 | 说明 |
|------|------|
| 📖 自动读取 | 从总考勤表自动提取全部姓名 |
| 🔍 按名筛选 | 按提供的人员名单精确匹配 |
| 📐 格式保留 | 复制时完整保留列宽、行高、字体、对齐方式 |
| ⚠️ 异常提示 | 未匹配到的人员自动列出提示 |
| 🖱️ GUI 操作 | 图形界面，三步完成 |

---

## 快速启动

### 环境要求

- Python ≥ 3.8
- WPS Office 已安装

### 安装与运行

```bash
pip install -r requirements.txt
python 考勤表整理工具.py
```

---

## 使用流程

```
选择总考勤表 → 输入/导入人员名单 → 点击生成 → 独立考勤表输出
```

1. **选择文件** — 选择包含所有人员的总考勤表（Excel）
2. **输入名单** — 手动输入或从文件导入需要筛选的人员姓名
3. **生成结果** — 点击「生成考勤表」，每人一份独立 Excel 文件

---

## 项目结构

```
attendance-tool/
└── 考勤表整理工具.py    # 主程序（WPS COM 自动化 + GUI）
```

---

## 开源许可

MIT © [H1nk5](https://github.com/H1nk5)
