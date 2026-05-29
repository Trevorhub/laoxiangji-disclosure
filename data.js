/** 真实证照示例（安徽老乡鸡餐饮有限公司） */
const BUSINESS_LICENSE = "assets/licenses/business-license.png";
const FOOD_LICENSE = "assets/licenses/food-license.png";

/** @typedef {{ business: string, food: string }} Licenses */
/** @typedef {{ id: string, province: string, city: string, name: string }} StoreBase */

/** @type {StoreBase[]} */
const STORE_LIST = [
  { id: "ah-hf-001", province: "安徽省", city: "合肥市", name: "老乡鸡合肥政务区店" },
  { id: "ah-hf-002", province: "安徽省", city: "合肥市", name: "老乡鸡合肥天鹅湖店" },
  { id: "ah-wh-001", province: "安徽省", city: "芜湖市", name: "老乡鸡芜湖镜湖店" },
  { id: "js-nj-001", province: "江苏省", city: "南京市", name: "老乡鸡南京新街口店" },
  { id: "js-nj-002", province: "江苏省", city: "南京市", name: "老乡鸡南京河西店" },
  { id: "js-sz-001", province: "江苏省", city: "苏州市", name: "老乡鸡苏州工业园区店" },
  { id: "zj-hz-001", province: "浙江省", city: "杭州市", name: "老乡鸡杭州西湖店" },
  { id: "zj-hz-002", province: "浙江省", city: "杭州市", name: "老乡鸡杭州钱江新城店" },
  { id: "sh-001", province: "上海市", city: "上海市", name: "老乡鸡上海陆家嘴店" },
  { id: "sh-002", province: "上海市", city: "上海市", name: "老乡鸡上海徐家汇店" },
  { id: "hb-wh-001", province: "湖北省", city: "武汉市", name: "老乡鸡武汉光谷店" },
  { id: "gd-gz-001", province: "广东省", city: "广州市", name: "老乡鸡广州天河店" },
];

/** 为全部门店统一挂载真实证照 */
/** @type {Array<StoreBase & { licenses: Licenses }>} */
const STORES = STORE_LIST.map((store) => ({
  ...store,
  licenses: {
    business: BUSINESS_LICENSE,
    food: FOOD_LICENSE,
  },
}));
