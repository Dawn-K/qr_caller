# 电话确认二维码网页

## 启动

```bash
npm install
npm run sample:generate
npm run dev
```

打开 `http://127.0.0.1:5500`。

## 使用流程

1. 先阅读页面上的权限说明，点击“授权并打开 Excel”。
2. 在浏览器弹窗里选择你的 `.xlsx/.xls` 文件，并允许读写权限。
3. 在“未处理”栏目中，扫码二维码进行拨号。
4. 通话后点击“接受”或“拒绝”，系统会自动写回并跳到下一条。
5. 如果你拒绝权限，页面会进入阻断态，需点击“重新授权”恢复。

## 示例文件

- `sample/sample.seed.csv`：示例种子数据（纳入 Git）。
- `sample/sample.xlsx`：本地生成文件（不纳入 Git）。

生成示例 Excel：

```bash
npm run sample:generate
```

生成后的 `sample/sample.xlsx` 可直接用于功能验证。

## 仓库说明

- 已忽略 `node_modules/`、`dist/`、`sample/*.xlsx` 等不必要文件。
- 不提交二进制示例，避免仓库历史膨胀；通过脚本从 CSV 一键生成。

## 浏览器建议

- 仅支持 Chrome / Edge（需要 File System Access API）。
- 文件全程在本地浏览器内处理，不会上传到服务器。
