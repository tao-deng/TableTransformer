# 表格数据提取工具

一个基于 Web 的表格数据提取工具，可以从总表中筛选出指定的数据并导出。

## 功能

- 支持多种表格格式：xlsx, xls, csv, tsv
- 上传总表和筛选表，自定义提取条件
- 支持精确匹配、包含匹配、开头匹配
- 支持导出为 Excel (xlsx) 或 CSV 格式
- 纯前端处理，数据不上传服务器，保护隐私

## 使用方法

1. 上传总表（数据源）
2. 上传筛选表（包含需要提取的条目）
3. 选择筛选列和匹配模式
4. 点击"提取数据"
5. 下载结果

## 部署到 GitHub Pages

1. 将此仓库推送到 GitHub
2. 进入仓库 Settings > Pages
3. 在 Source 中选择 Deploy from a branch
4. 选择 main 分支和 / 根目录
5. 保存后即可通过 `https://你的用户名.github.io/TableTransformer/` 访问

## 技术栈

- 纯 HTML/CSS/JavaScript
- SheetJS (xlsx) - 表格解析库
- GitHub Pages - 静态网站托管
