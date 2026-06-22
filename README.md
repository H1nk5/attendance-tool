<div align="center">

# 📋 Attendance Tool

**考勤表整理工具 · 按姓名筛选生成独立考勤表**

<br>

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![WPS](https://img.shields.io/badge/WPS-0078D4?style=for-the-badge&logo=wps&logoColor=white)

<br>

从总考勤表中按姓名列表筛选人员，生成独立考勤表 · 保留原始格式与列宽

[功能特性](#功能特性) · [快速启动](#快速启动) · [使用说明](#使用说明)

</div>

---

## 功能特性

- 📖 从总考勤表自动读取全部姓名
- 🔍 按名单筛选指定人员
- 📐 复制时保留格式、列宽、行高
- ⚠️ 未匹配人员自动提示
- 🖱️ GUI 界面，操作简单直观

---

## 快速启动

### 环境要求

- Python ≥ 3.8
- WPS Office 已安装

### 安装依赖

```bash
pip install pywin32
```

### 运行

```bash
python 考勤表整理工具.py
```

---

## 使用说明

1. 运行程序
2. 选择总考勤表文件（Excel）
3. 输入或读取人员名单
4. 点击「生成考勤表」
5. 程序自动生成每人独立的考勤表

---

## 开源许可

MIT © [H1nk5](https://github.com/H1nk5)
