/** 证照示例图 */
const PLATFORM_BUSINESS_LICENSE = "assets/licenses/business-license.png";
const OPERATING_LICENSE = "assets/licenses/food-license.png";
const PREPAID_CARD_FILING = "assets/licenses/prepaid-card-filing.svg";

/** @typedef {{ business: string, operating: string, prepaid: string }} PlatformLicenses */
/** @typedef {{ id: string, channel: string, name: string, desc: string, licenses: PlatformLicenses }} Platform */

/** 平台资质三类证照 */
const PLATFORM_LICENSE_SAMPLES = {
  business: PLATFORM_BUSINESS_LICENSE,
  operating: OPERATING_LICENSE,
  prepaid: PREPAID_CARD_FILING,
};

/** @type {Platform[]} */
const PLATFORMS = [
  {
    id: "lxj-platform",
    channel: "平台",
    name: "老乡鸡线上经营平台",
    desc: "安徽老乡鸡餐饮有限公司",
    licenses: { ...PLATFORM_LICENSE_SAMPLES },
  },
  {
    id: "miniapp-wechat",
    channel: "微信",
    name: "老乡鸡微信小程序",
    desc: "官方点餐 / 会员服务",
    licenses: { ...PLATFORM_LICENSE_SAMPLES },
  },
  {
    id: "miniapp-alipay",
    channel: "支付宝",
    name: "老乡鸡支付宝小程序",
    desc: "官方点餐 / 会员服务",
    licenses: { ...PLATFORM_LICENSE_SAMPLES },
  },
  {
    id: "meituan",
    channel: "美团",
    name: "老乡鸡美团外卖",
    desc: "美团外卖平台官方旗舰店",
    licenses: { ...PLATFORM_LICENSE_SAMPLES },
  },
  {
    id: "eleme",
    channel: "饿了么",
    name: "老乡鸡饿了么外卖",
    desc: "饿了么平台官方旗舰店",
    licenses: { ...PLATFORM_LICENSE_SAMPLES },
  },
];
