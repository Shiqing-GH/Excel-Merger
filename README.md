# Excel 合并工具（ExcelMerger）

把文件夹内所有 Excel 文件的**第一个 sheet** 合并到一个文件，每个 sheet 以原文件名命名。

## 使用方法

1. 从本仓库的 Actions 页面下载最新打包好的 `ExcelMerger.exe`。
2. 双击 exe，弹出对话框后选择存放 Excel 文件的文件夹。
3. 程序自动在该文件夹下生成 `合并结果.xlsx`。

也可以用命令行指定输入输出：

```bash
ExcelMerger.exe 输入文件夹 输出.xlsx
```

## 自己重新打包

推送到 main 分支后，GitHub Actions 会自动打包；也可在 Actions 页面手动运行工作流。
