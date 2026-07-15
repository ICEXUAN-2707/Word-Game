# Word-Game

这是一个静态 PWA 单词卡页面，可部署到 GitHub Pages。

## 本地预览

```powershell
python -m http.server 8766 --bind 127.0.0.1
```

打开：

```text
http://127.0.0.1:8766/
```

## GitHub Pages 设置

1. 在 GitHub 新建公开仓库，例如 `word-cards`。
2. 将本目录推送到仓库 `main` 分支。
3. 打开仓库 `Settings -> Pages`。
4. `Source` 选择 `Deploy from a branch`。
5. `Branch` 选择 `main`，目录选择 `/root`，保存。

发布地址：

```text
https://icexuan-2707.github.io/Word-Game/
```
