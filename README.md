# 红色电影问卷 XLSX 样本生成器

这是一个独立的本地前端工具，用来为“红色电影《阮啸仙》对新时代大学生红色基因传承的实践调研问卷”生成合理的 `.xlsx` 样本答案文件。

工具只生成 Excel 样本数据，不包含问卷自动填写、自动提交、代理、验证码处理或浏览器自动化功能。

## 使用方式

双击根目录的 `start.bat`，或运行：

```powershell
.\start.bat
```

脚本会进入 `red_movie_xlsx_tool`，创建或激活 conda 环境 `red-movie-xlsx`，安装必要依赖，然后打开：

```text
http://127.0.0.1:8765
```

在页面中设置样本数量、随机种子和题目 JSON，点击“生成 XLSX”即可。默认题纲按问卷链接 `https://v.wjx.cn/vm/t4Mk5kh.aspx` 的 18 题整理。

生成文件保存在：

```text
E:\2026\question\red_movie_xlsx_tool\generated_samples
```

## 输出格式

- Sheet 名称：`样本答案`
- 第一列：`样本编号`
- 后续列：每列一道题
- 单选题：写入一个选项
- 多选题：严格使用问卷星多选分隔符 `┋` 分隔多个选项，Unicode 为 `U+250B`
- 文件名包含 `wjx_split_250b`，方便和旧文件区分
- 文本题：生成 1-2 句自然中文建议
- 不生成姓名、手机号、身份证等身份字段

注意：不要使用旧的 `|` 或 `；` 文件导入问卷星；最新文件名应类似 `red_movie_wjx_split_250b_20260526_*.xlsx`。

## 测试

```powershell
conda run -n red-movie-xlsx python -m pytest red_movie_xlsx_tool\tests
```
