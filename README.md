# 老乡鸡经营证照公示

移动端 H5 经营证照公示平台，包含 **门店资质** 与 **平台资质** 两大模块。

## 页面结构

| 页面 | 说明 |
|------|------|
| `index.html` | 首页入口，选择资质类型 |
| `stores.html` | 门店资质（原门店信息公示） |
| `platform.html` | 平台资质（小程序、外卖等） |

## 功能

### 门店资质
- 省份 / 城市筛选、门店名称搜索
- 查看营业执照、食品经营许可证
- 点击卡片空白区域查看全部证照，各自独立旋转

### 平台资质
- 平台营业执照、经营许可证、单用途预付卡备案
- 平台数据：`platform-data.js` → `PLATFORMS`

## 本地预览

```bash
cd /Users/barryallen/Projects/laoxiangji-disclosure
python3 -m http.server 8080
```

访问 http://localhost:8080

## 线上地址

**https://laoxiangji-disclosure.onrender.com**

## 数据维护

- 门店：`data.js` → `STORE_LIST`
- 平台：`platform-data.js` → `PLATFORMS`
- 证照弹层逻辑：`license-modal.js`（门店页、平台页共用）
