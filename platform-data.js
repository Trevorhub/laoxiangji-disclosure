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
];
