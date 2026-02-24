# PPT 一键美化 - Office 插件版

基于 Office JS 的 PowerPoint 插件，直接在 PPT 里一键美化。

## 功能

- ✅ 统一字体和字号
- ✅ 优化配色方案
- ✅ 调整元素布局
- ✅ 对齐和间距优化

## 本地开发

### 1. 安装依赖

```bash
npm install
```

### 2. 生成 HTTPS 证书（Office 插件必须用 HTTPS）

```bash
# 安装 mkcert
brew install mkcert  # macOS
# 或 choco install mkcert  # Windows

# 生成证书
mkcert -install
mkcert localhost
```

### 3. 启动开发服务器

```bash
npm run dev
```

### 4. 在 PowerPoint 中加载插件

**Windows:**
1. 打开 PowerPoint
2. 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录
3. 添加 `https://localhost:3000/manifest.xml`
4. 重启 PowerPoint
5. 插入 → 我的加载项 → 共享文件夹 → PPT 一键美化

**Mac:**
1. 打开 PowerPoint
2. 插入 → 加载项 → 我的加载项
3. 点击「管理我的加载项」
4. 上传 manifest.xml

**Web 版 PowerPoint:**
1. 打开 PowerPoint Online
2. 插入 → Office 加载项
3. 上传我的加载项 → 选择 manifest.xml

## 项目结构

```
ppt-beautify-addin/
├── manifest.xml      # Office 插件清单
├── taskpane.html     # 插件界面
├── src/
│   └── taskpane.js   # 核心逻辑
├── assets/           # 图标资源
└── package.json
```

## 配置说明

### 配色方案

在 `src/taskpane.js` 中修改 `COLOR_SCHEMES`:

```javascript
const COLOR_SCHEMES = {
  professional: {
    primary: '#1a365d',    // 主色
    accent: '#ed8936',     // 强调色
    text: '#2d3748',       // 文字色
  }
};
```

### 字体配置

```javascript
const FONT_CONFIG = {
  title: { name: '微软雅黑', size: 36 },
  heading: { name: '微软雅黑', size: 28 },
  body: { name: '微软雅黑', size: 18 },
};
```

## 发布

1. 修改 `manifest.xml` 中的 URL 为生产环境地址
2. 部署静态文件到 HTTPS 服务器
3. 提交到 Microsoft AppSource（可选）

## 注意事项

- Office 插件必须使用 HTTPS
- 首次加载需要信任证书
- 支持 PowerPoint 2016+ 和 PowerPoint Online
