# 豆瓣书影音游导出 NeoDB Excel

一个轻量的 Tampermonkey 用户脚本，将当前登录豆瓣账号的书影音游标记导出为 NeoDB 可导入的 `.xlsx` 文件。

脚本在浏览器中直接读取豆瓣数据并生成 Excel，不需要安装完整浏览器扩展。

## 支持范围

| 分类 | 状态 |
| --- | --- |
| 电影 | 想看、在看、看过 |
| 音乐 | 想听、在听、听过 |
| 图书 | 想读、在读、读过 |
| 游戏 | 想玩、在玩、玩过 |
| 舞台剧 | 想看、看过 |

暂不包含广播、评论、日记、相册、关注列表、豆列、豆邮和断点续跑。

## 安装

1. 安装 [Tampermonkey](https://www.tampermonkey.net/)。
2. 点击 [安装脚本](https://raw.githubusercontent.com/dengshu2/douban-neodb-export-userscript/main/douban-neodb-export.user.js)。
3. 在 Tampermonkey 页面确认安装。
4. 打开豆瓣并登录账号，然后刷新页面。

## 使用

1. 点击豆瓣页面右下角的“导出豆瓣记录”。
2. 保持页面打开，等待 Excel 下载完成。
3. 打开 NeoDB 的导入页面。
4. 在 “Import Marks and Reviews from Douban” 中选择下载的 `.xlsx` 文件。

也可以从 Tampermonkey 菜单中选择“导出豆瓣书影音游记录”。

## 说明

- 请求间隔固定为 1 秒，避免在短时间内向豆瓣发送过多请求。
- Excel 使用 [SheetJS](https://sheetjs.com/) 在浏览器本地生成。
- 生成的工作表名称遵循 [Doufen](https://github.com/doufen-org/tofu) 导出约定，并兼容 [NeoDB 的 Douban 导入器](https://github.com/neodb-social/neodb/blob/main/journal/importers/douban.py)。
- 脚本不会将导出内容上传到第三方服务。导入 NeoDB 时，文件会由用户主动提交到所选择的 NeoDB 实例。

## 开发

本项目不依赖构建工具。修改后运行：

```powershell
npm test
```

## 致谢

本项目参考了 MIT 许可的 [doufen-org/tofu](https://github.com/doufen-org/tofu)，复用了其浏览器端导出思路和 Excel 工作表约定。

## License

[MIT](LICENSE)
