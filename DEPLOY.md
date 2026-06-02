# 老乡鸡经营证照公示 — 静态站点部署说明

## 包内容

本目录为可直接发布的静态 H5 站点，无需构建步骤。

- 入口：`index.html`
- 门店资质：`stores.html`（数据：`data.js`）
- 平台资质：`platform.html`（数据：`platform-data.js`）

门店证照图片使用 OSS 外链，需服务器可访问公网图片地址。

## 部署方式

### 1. Nginx

将本目录下所有文件放到站点根目录，例如 `/var/www/laoxiangji-disclosure/`：

```nginx
server {
    listen 80;
    server_name your-domain.com;
    root /var/www/laoxiangji-disclosure;
    index index.html;

    location / {
        try_files $uri $uri/ /index.html;
    }

    # 静态资源缓存（可选）
    location ~* \.(js|css|png|jpg|jpeg|svg|ico)$ {
        expires 7d;
        add_header Cache-Control "public";
    }
}
```

### 2. 对象存储 + CDN（阿里云 OSS / 腾讯云 COS 等）

1. 开启「静态网站托管」
2. 上传压缩包解压后的全部文件到 Bucket 根目录
3. 默认首页设为 `index.html`
4. 绑定自定义域名并开启 HTTPS

### 3. 任意静态托管

将文件上传到 IIS、Apache、`python -m http.server` 等均可，保证：

- 根路径能访问 `index.html`
- 子页面 `stores.html`、`platform.html` 路径可直达

## 本地快速验证

```bash
cd /path/to/解压目录
python3 -m http.server 8080
```

浏览器打开：http://127.0.0.1:8080

## 更新数据

门店列表在 `data.js` 的 `STORES` 数组；平台在 `platform-data.js` 的 `PLATFORMS`。替换后重新上传整站或仅替换对应 JS 文件即可。

## 外部跳转（门店 deep link）

外部系统可通过 URL 参数 `storeId` 打开指定门店并自动定位，详见 **[EXTERNAL_LINK.md](./EXTERNAL_LINK.md)**。

示例：

```
https://your-domain.com/stores.html?storeId=8037
```
