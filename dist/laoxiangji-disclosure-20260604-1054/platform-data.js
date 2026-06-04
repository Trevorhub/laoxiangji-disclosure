/** 证照示例图 */
const PLATFORM_BUSINESS_LICENSE = "assets/licenses/business-license.png";
const PLATFORM_FOOD_LICENSE = "assets/licenses/food-license.png";

/** @typedef {{ business: string, operating: string }} PlatformLicenses */
/** @typedef {{ id: string, channel: string, name: string, desc: string, licenses: PlatformLicenses }} Platform */

/** 平台资质证照：平台营业执照、食品经营许可证 */
const PLATFORM_LICENSE_SAMPLES = {
  business: PLATFORM_BUSINESS_LICENSE,
  operating: PLATFORM_FOOD_LICENSE,
};

/** @type {Platform[]} */
const PLATFORMS = [
  {
    id: "lxj-platform",
    channel: "平台",
    name: "网络平台运营主体",
    desc: "安徽老乡鸡餐饮有限公司",
    licenses: { ...PLATFORM_LICENSE_SAMPLES },
  },
];
