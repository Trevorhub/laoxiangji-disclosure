# 老乡鸡餐厅信息公示

移动端 H5 餐厅信息公示页面，参考麦当劳公示页布局，采用老乡鸡绿色主题。

## 移动端适配

- `flex.js`：rem 弹性布局（设计稿基准 375px，大屏封顶 540px）
- `viewport-fit=cover`：适配刘海屏 / 全面屏安全区
- 桌面浏览器打开时居中显示手机宽度画布，不无限拉宽

## 功能

- 省份 / 城市二级筛选门店
- 门店名称关键词搜索
- 点击「营业执照」「食品经营许可证」底部弹层预览证照

## 本地预览

```bash
cd /Users/barryallen/Projects/laoxiangji-disclosure
python3 -m http.server 8080
```

手机与电脑同一局域网时，访问 `http://<本机IP>:8080`；本机可直接打开 `http://localhost:8080`。

## 数据对接

门店数据在 `data.js` 的 `STORE_LIST` 中维护；全部门店证照默认使用 `assets/licenses/business-license.png` 与 `food-license.png` 真实示例图。接入后端后，在 `map` 逻辑中按门店返回各自证照 URL 即可。
