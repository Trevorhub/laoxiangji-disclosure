/** 真实证照示例（安徽老乡鸡餐饮有限公司） */
const BUSINESS_LICENSE = "assets/licenses/business-license.png";
const FOOD_LICENSE = "assets/licenses/food-license.png";

/** @typedef {{ business: string, food: string }} Licenses */
/** @typedef {{ id: string, channel: string, name: string, desc: string, licenses: Licenses }} Platform */

/** @type {Platform[]} */
const PLATFORMS = [
  {
    id: "miniapp-wechat",
    channel: "微信",
    name: "老乡鸡微信小程序",
    desc: "官方点餐 / 会员服务",
    licenses: { business: BUSINESS_LICENSE, food: FOOD_LICENSE },
  },
  {
    id: "miniapp-alipay",
    channel: "支付宝",
    name: "老乡鸡支付宝小程序",
    desc: "官方点餐 / 会员服务",
    licenses: { business: BUSINESS_LICENSE, food: FOOD_LICENSE },
  },
  {
    id: "meituan",
    channel: "美团",
    name: "老乡鸡美团外卖",
    desc: "美团外卖平台官方旗舰店",
    licenses: { business: BUSINESS_LICENSE, food: FOOD_LICENSE },
  },
  {
    id: "eleme",
    channel: "饿了么",
    name: "老乡鸡饿了么外卖",
    desc: "饿了么平台官方旗舰店",
    licenses: { business: BUSINESS_LICENSE, food: FOOD_LICENSE },
  },
  {
    id: "corp",
    channel: "主体",
    name: "安徽老乡鸡餐饮有限公司",
    desc: "平台经营主体资质",
    licenses: { business: BUSINESS_LICENSE, food: FOOD_LICENSE },
  },
];
