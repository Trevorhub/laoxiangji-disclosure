/**
 * 门店资质数据（加盟 Excel + 直营 CSV 合并生成）
 * - 加盟门店证照-最新的 2.xlsx（重复门店以本表为准）
 * - 直营营业的店.csv
 *
 * 合并规则：
 * 1) 同门店按 省份+城市+门店名称 合并为单条。
 * 2) 同门店同证照多条时，按 URL 尾部时间戳保留最新一条。
 * 3) 省市名称做归一化（如 安徽/安徽省、上海/上海市 合并）。
 * 4) 加盟与直营重复时，以加盟 Excel 数据覆盖。
 */
const STORES = [
  {
    "id": "8037",
    "province": "上海市",
    "city": "上海市",
    "name": "上海1788广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/464fb55a-45e5-49e2-bae9-a5c60f44ee401766855372336.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4325e78b-a2bf-4294-9141-a3e3208414dc1766855372687.jpg"
    }
  },
  {
    "id": "8104",
    "province": "上海市",
    "city": "上海市",
    "name": "上海5G未来中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4ee67f59-1d90-40fd-97f7-381dcfdc68d91663607234927.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f092311d-c679-493d-a2c3-00c26ba39cae1663607235312.jpg"
    }
  },
  {
    "id": "8074",
    "province": "上海市",
    "city": "上海市",
    "name": "上海IBP总部店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6da2b248-d9d8-49ce-b676-2a3ca54700c71662621829602.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/60441d1d-126e-472e-9558-6c4a07dcb6ab1662621829960.jpg"
    }
  },
  {
    "id": "8096",
    "province": "上海市",
    "city": "上海市",
    "name": "上海LCM置汇旭辉店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e67112b2-9f7f-4e98-95c8-68965ab3d9fb1662621853707.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/43a47fab-1893-4d3d-a7e7-caa8b2ef030f1662621854150.jpg"
    }
  },
  {
    "id": "8061",
    "province": "上海市",
    "city": "上海市",
    "name": "上海SK大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c896560a-3fda-45ce-a793-b5db21573c771664471225283.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e915d0cf-50f6-428d-b1fc-b61338b30a671664471225653.jpg"
    }
  },
  {
    "id": "8051",
    "province": "上海市",
    "city": "上海市",
    "name": "上海七宝宝龙城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7e3b170d-943e-44e9-8558-38d45aa751591775840820240.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a131d9bf-7ca2-432b-a69e-686be1f26ff41775840821021.jpg"
    }
  },
  {
    "id": "8133",
    "province": "上海市",
    "city": "上海市",
    "name": "上海七宝领展广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8f802603-35ee-49a5-b76f-5eb588bb73c71690390950901.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bcb6ee97-0fb9-4a85-b13b-88f91a9b91671690390951808.jpg"
    }
  },
  {
    "id": "8160",
    "province": "上海市",
    "city": "上海市",
    "name": "上海万嘉商业广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/27a9bbbf-9fb0-4b56-b284-c79fdac923d51704215367077.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8481228b-d229-47ef-8af2-dbab348ecbda1704215367583.jpg"
    }
  },
  {
    "id": "8153",
    "province": "上海市",
    "city": "上海市",
    "name": "上海万科天空之城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/419f3633-a33e-4d2f-8ff3-5ec4615710d11696871171222.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2c20ae7f-ed87-4b47-ba0a-5025403e8a431696871191019.jpg"
    }
  },
  {
    "id": "8005",
    "province": "上海市",
    "city": "上海市",
    "name": "上海万科翡翠公园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fac88e40-8027-49ca-b16d-22e1ffcb725b1715706341877.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/af6c4c35-cba6-4fd3-8a76-475c3fb9d36c1715706342417.jpg"
    }
  },
  {
    "id": "8150",
    "province": "上海市",
    "city": "上海市",
    "name": "上海世博绿谷店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/966b5ca8-45d0-4da3-9d2d-8b6c562eb1da1696871135700.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2bb8544d-d70c-4f1b-a102-5b004cdcc6731696871143495.jpg"
    }
  },
  {
    "id": "8124",
    "province": "上海市",
    "city": "上海市",
    "name": "上海东方懿德城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/65b6f944-5692-446d-991b-edf90d90a5f61689267743400.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2a17ea0c-f7c3-4c5d-a10b-ced4d77610591689267744282.jpg"
    }
  },
  {
    "id": "8020",
    "province": "上海市",
    "city": "上海市",
    "name": "上海东苑丽宝店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/73bc502c-2f86-4845-94e9-87f497b00b0f1750179991070.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/deb9c1a4-e44d-4405-84b1-96cd3d766e2a1750179991616.jpg"
    }
  },
  {
    "id": "8122",
    "province": "上海市",
    "city": "上海市",
    "name": "上海中信泰富万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/12a43331-40a3-41a6-9deb-e28bcf99efd41675443757814.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ee1443b5-8d41-443f-8920-686822a72e571675443758212.jpg"
    }
  },
  {
    "id": "8087",
    "province": "上海市",
    "city": "上海市",
    "name": "上海中庚漫游城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/66040649-c391-41b5-b740-ff76f5ad4de21685811730251.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e29d0de0-9c86-4cad-b3bb-5142d5b562c91685811731253.jpg"
    }
  },
  {
    "id": "8032",
    "province": "上海市",
    "city": "上海市",
    "name": "上海中心大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9b7ff7ba-ccb0-41cd-b7d0-d204caf695751765991195823.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8fa448d0-2675-47b4-a4f3-6fa7ac47b55f1765991196529.jpg"
    }
  },
  {
    "id": "8115",
    "province": "上海市",
    "city": "上海市",
    "name": "上海中铁中环店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d4824f2f-293d-4f2e-9f25-59e68b146d001668013656105.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/06f1c9e7-24ba-474b-8b73-5ec310f740971668013656431.jpg"
    }
  },
  {
    "id": "8102",
    "province": "上海市",
    "city": "上海市",
    "name": "上海中骏广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0b17fdda-dd01-4738-a33d-069df3faa44f1668013648316.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/42dde2af-2496-46e8-a874-6effc9e9447c1668013648755.jpg"
    }
  },
  {
    "id": "8065",
    "province": "上海市",
    "city": "上海市",
    "name": "上海乐虹坊店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6b1ada43-f628-4d5c-8d96-aeb7a606dc741662621820906.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d370d999-352a-4a4e-a268-79c818c9bd861662621821289.jpg"
    }
  },
  {
    "id": "8111",
    "province": "上海市",
    "city": "上海市",
    "name": "上海九亭店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c032e93e-b36a-4800-817e-f810064eef151685811735084.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c802aeee-1cdb-4d76-9b33-c8a2147780b81685811735880.jpg"
    }
  },
  {
    "id": "8105",
    "province": "上海市",
    "city": "上海市",
    "name": "上海云飞大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/49a4afef-f96c-43e1-a0b6-5bd273301e311662621860205.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a88a93a2-849a-4b3e-9cc5-b0dac93c5e361662621860644.jpg"
    }
  },
  {
    "id": "8012",
    "province": "上海市",
    "city": "上海市",
    "name": "上海五角场万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c81150ef-fbfd-4c24-9f64-3b3719adb9ff1734541574960.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0ebc8c96-1a7c-4e96-ba2b-a7ea50d5e1fc1734541575614.jpg"
    }
  },
  {
    "id": "1744",
    "province": "上海市",
    "city": "上海市",
    "name": "上海佘山欢乐谷店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1744/1744-上海佘山欢乐谷店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1744/1744-上海佘山欢乐谷店-经营许可证.jpg"
    }
  },
  {
    "id": "8043",
    "province": "上海市",
    "city": "上海市",
    "name": "上海佳兆业店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/17050349-d35a-4903-b150-f2ba3ce2115a1768928808424.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2854c0d0-1dcd-47c3-a303-e03cc765844c1768928808715.jpg"
    }
  },
  {
    "id": "8128",
    "province": "上海市",
    "city": "上海市",
    "name": "上海信业购物中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/039e4a35-1a80-4003-97e4-f050276bc8571676653398343.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0d4ba6ce-38f1-46df-ae2d-4ec16b9eec341676653398692.jpg"
    }
  },
  {
    "id": "8024",
    "province": "上海市",
    "city": "上海市",
    "name": "上海光启城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c08eea24-350b-4018-8523-6701a6af34c11756141577700.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/691fa984-cefe-4529-91ad-44c6ac2e2c041756141577977.jpg"
    }
  },
  {
    "id": "8100",
    "province": "上海市",
    "city": "上海市",
    "name": "上海凌空SOHO店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cceba48a-f4c8-4a4e-9261-f7912f2179c51686762120953.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/feb508d3-1510-413d-93c2-e9ba693bef6d1686762121709.jpg"
    }
  },
  {
    "id": "8038",
    "province": "上海市",
    "city": "上海市",
    "name": "上海凯德星贸店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/54f1be6b-9c8f-4aef-a9ab-d9cc53ee09281766768816550.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/031ef326-152e-4a68-aa0e-a2acc354dcb71766768816953.jpg"
    }
  },
  {
    "id": "8047",
    "province": "上海市",
    "city": "上海市",
    "name": "上海北外滩来福士店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1a148c3b-d85e-4f53-9aec-44ba557bb0ab1775581609248.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/6n0q5meuals39la09kwkrj572_01778464134531.jpg"
    }
  },
  {
    "id": "8125",
    "province": "上海市",
    "city": "上海市",
    "name": "上海华润万家光新店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fb62f62d-7a41-4fba-8908-b7b382e72d971689267744841.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/18b18857-6177-4cf5-b340-03d474ef28951689267745294.jpg"
    }
  },
  {
    "id": "8109",
    "province": "上海市",
    "city": "上海市",
    "name": "上海南桥科技城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ab5d9898-9523-494f-9cef-5eb0440e274b1662621861546.png",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dbae2090-0f42-4420-89b4-2afe455b12b21662621861931.png"
    }
  },
  {
    "id": "8097",
    "province": "上海市",
    "city": "上海市",
    "name": "上海南翔五彩城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8c28980f-cf05-440f-b7bb-eb08a230c7f51662621855032.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1f85df66-c321-4ff6-b7b8-1918c9e44d2c1662621855535.jpg"
    }
  },
  {
    "id": "8013",
    "province": "上海市",
    "city": "上海市",
    "name": "上海古北1699店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/736d5b00-716b-405e-9442-041f6ccce06f1745341556336.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/465eb6ba-e3fc-470b-a525-416892ad1fc21745341556687.jpg"
    }
  },
  {
    "id": "8068",
    "province": "上海市",
    "city": "上海市",
    "name": "上海合川大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fea0fc38-7762-4200-a0ea-9889f65455001662621823507.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fcb3d90e-1469-4c2b-808b-d34252a3082d1662621823875.jpg"
    }
  },
  {
    "id": "8019",
    "province": "上海市",
    "city": "上海市",
    "name": "上海品尊国际中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9668a480-3963-4c2d-88c7-7eca5a3c2a891755536782783.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4278ca4d-e15e-4a76-8365-a0ec7d72aa411755536783489.jpg"
    }
  },
  {
    "id": "8108",
    "province": "上海市",
    "city": "上海市",
    "name": "上海嘉定宝龙广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5fe379e2-bbb8-40b4-b1f9-52554be7a0271668013651852.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2bf34991-4b6b-4890-92db-dd91be594d621668013652268.jpg"
    }
  },
  {
    "id": "8091",
    "province": "上海市",
    "city": "上海市",
    "name": "上海国创中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a054653c-b1e7-49e9-b7af-f0647caeeea91662621847258.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2fd49a0e-3c32-433f-9b78-85940231313a1662621847609.jpg"
    }
  },
  {
    "id": "8092",
    "province": "上海市",
    "city": "上海市",
    "name": "上海国华广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/01a83438-2200-4b65-8c0d-bfe8540d357a1662621848467.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/270953da-3b9f-4755-9604-934cd7bf046d1662621848821.jpg"
    }
  },
  {
    "id": "8049",
    "province": "上海市",
    "city": "上海市",
    "name": "上海复地浦江中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/13ab9bd1-f01f-42c7-b819-1da991b15c571773680823763.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ad695cf3-de81-40e5-8335-5b4d38973ebb1773680824682.jpg"
    }
  },
  {
    "id": "8036",
    "province": "上海市",
    "city": "上海市",
    "name": "上海大唐乐坊店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5c808032-726f-4901-bdbb-ac86b045eab71766164016188.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0b2de622-d727-461f-9ecf-0aa3863076e21766164016528.jpg"
    }
  },
  {
    "id": "8058",
    "province": "上海市",
    "city": "上海市",
    "name": "上海大宁音乐广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/855403c8-2db5-4711-a46b-9c22a0fb8a061662621813556.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a1e55d6b-1fb6-4c31-b3d8-217c5111bce21662621813927.jpg"
    }
  },
  {
    "id": "8009",
    "province": "上海市",
    "city": "上海市",
    "name": "上海大渡河路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ab62f032-9bbc-4f18-9d28-e74a92bbb1441735664756159.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/171517e7-2e8b-4e97-ace7-14898ffbcde11735664756693.jpg"
    }
  },
  {
    "id": "8010",
    "province": "上海市",
    "city": "上海市",
    "name": "上海大都会店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/77eaaf14-f513-40b0-8c5d-12938dafbeb51735837521306.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bec54835-9b5c-4fec-ada2-2e6af8632d081735837521713.jpg"
    }
  },
  {
    "id": "8007",
    "province": "上海市",
    "city": "上海市",
    "name": "上海天目西路第一分公司",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/69e78e30-d99c-4947-bd89-b9cfd63cd5501731690359124.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/02015539-e52c-4470-a235-390b6841baf91731690359914.jpg"
    }
  },
  {
    "id": "8035",
    "province": "上海市",
    "city": "上海市",
    "name": "上海奉贤宝龙广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c0a3ef04-96c0-4ba6-92d0-dfc1aee489591764608811221.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品经营许可证1767157366819.pdf"
    }
  },
  {
    "id": "8050",
    "province": "上海市",
    "city": "上海市",
    "name": "上海守信创意园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3df5cf76-a45d-4f2f-858a-0f8c3a9d915b1662621805049.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3554a5e6-d03e-4e80-92f4-cb5fc48818381662621806068.jpg"
    }
  },
  {
    "id": "8112",
    "province": "上海市",
    "city": "上海市",
    "name": "上海安亭财富广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/50d3886a-431f-449f-845d-a0a00c7e9b281664471250777.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/322781d4-5599-41bf-92d9-5a8ca1293c181664471251104.jpg"
    }
  },
  {
    "id": "8127",
    "province": "上海市",
    "city": "上海市",
    "name": "上海宝山巴黎春天店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6ab67dfe-d6c5-46dd-b937-a2113d725d901689267748849.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/429d76b8-3ddb-4af4-bd24-da965aaa471f1689267749287.jpg"
    }
  },
  {
    "id": "8062",
    "province": "上海市",
    "city": "上海市",
    "name": "上海宝山日月光店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/862af455-e16a-48fc-b4a1-8285b9fc11c71664471227471.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7b5d5824-5e45-46e9-b998-df3afa2b71171664471227794.jpg"
    }
  },
  {
    "id": "8248",
    "province": "上海市",
    "city": "上海市",
    "name": "上海宝龙旭辉店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/新建 Microsoft Word 文档 (2)1779958264202.pdf",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/新建 Microsoft Word 文档 (1)1779958206566.pdf"
    }
  },
  {
    "id": "8030",
    "province": "上海市",
    "city": "上海市",
    "name": "上海尚悦湾店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/925b77d7-9aa8-4203-9086-ac6c93a1a6d01763053619146.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2e7d1ace-705b-4a08-865c-1e96c4efe51e1763053619507.jpg"
    }
  },
  {
    "id": "8229",
    "province": "上海市",
    "city": "上海市",
    "name": "上海展讯中心店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/营业执照1768547827784.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品证1768547839457.jpg"
    }
  },
  {
    "id": "8054",
    "province": "上海市",
    "city": "上海市",
    "name": "上海张江四标科技园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c40a4f6b-97eb-415f-ac93-2b9da78148ae1662621809475.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/36f2024b-e9a9-4061-828f-7e4cacb839d21662621809826.jpg"
    }
  },
  {
    "id": "8034",
    "province": "上海市",
    "city": "上海市",
    "name": "上海徐汇南宁路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a7e67c09-f976-4b6d-91bc-3ebcaf2b62a31763658423110.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ec449598-4b41-4f3e-9504-ffcbc4a8dd5f1763658423844.jpg"
    }
  },
  {
    "id": "8042",
    "province": "上海市",
    "city": "上海市",
    "name": "上海徐汇正大乐城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/55400d87-3220-45d3-951a-c9445ffa92861763744811684.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b4a403b8-d7ce-4d51-9c10-bb27698dafbc1763744812367.jpg"
    }
  },
  {
    "id": "8094",
    "province": "上海市",
    "city": "上海市",
    "name": "上海徐汇绿地缤纷城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e0d9f72f-4e15-47bd-b08f-a58f789cdc5a1662621850890.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3b9a57c0-f007-4e54-96b7-7aa053cc12bb1662621851227.jpg"
    }
  },
  {
    "id": "8130",
    "province": "上海市",
    "city": "上海市",
    "name": "上海悠方广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8b4572ce-2d02-4a90-be83-6b89d873053b1675443758805.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/505daabc-46c1-4244-904c-1521d44e98081675443759162.jpg"
    }
  },
  {
    "id": "8147",
    "province": "上海市",
    "city": "上海市",
    "name": "上海成山巴黎春天店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/51055a2d-919e-4cae-9ed6-ed0f3881ffea1689267752268.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2a4a48db-0948-45e7-af1c-d61e99000a241689267752742.jpg"
    }
  },
  {
    "id": "8113",
    "province": "上海市",
    "city": "上海市",
    "name": "上海文峰广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5e5f283b-09fe-42de-b6b1-1eb76b495f001668013654833.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/373cd3ff-7cb8-490e-98bf-e9c580544d351668013655254.jpg"
    }
  },
  {
    "id": "8090",
    "province": "上海市",
    "city": "上海市",
    "name": "上海新大陆广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1295a8fe-211e-4942-bb81-d7fcb5ca107d1662621846045.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1276a637-955b-4d74-afe6-633a8d3e8a5b1662621846502.jpg"
    }
  },
  {
    "id": "8152",
    "province": "上海市",
    "city": "上海市",
    "name": "上海新宜街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/72232b66-d870-4266-8cbc-635f21e8d9161726765609677.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/50ecb51d-cae3-4822-a737-1c89ace52cd41726765610215.jpg"
    }
  },
  {
    "id": "8095",
    "province": "上海市",
    "city": "上海市",
    "name": "上海新邻天地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c51e457e-7c63-40b1-89f0-8351282b074a1664471244433.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6eea4dbc-847e-4982-8b01-d46a1232b1f61664471244817.jpg"
    }
  },
  {
    "id": "8039",
    "province": "上海市",
    "city": "上海市",
    "name": "上海星荟中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a9136201-9c9f-4c0f-87de-cff29a8979891769015207846.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cfdefa0f-eae6-4df1-95a0-ba82819d38321769015208909.jpg"
    }
  },
  {
    "id": "8064",
    "province": "上海市",
    "city": "上海市",
    "name": "上海普陀绿地缤纷城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/71f5372b-71ac-494c-9bea-addefbe85c531665248828946.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c3a54c28-ee83-470b-a19c-28b567be11701665248829307.jpg"
    }
  },
  {
    "id": "8057",
    "province": "上海市",
    "city": "上海市",
    "name": "上海杨浦保利绿地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1c164617-7afc-4416-9e7a-83479fb49cd21779728835874.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/96141f0d-c89b-4721-a141-fa817dcc0c3c1779728837569.jpg"
    }
  },
  {
    "id": "8016",
    "province": "上海市",
    "city": "上海市",
    "name": "上海松江万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0a0e7a08-ea67-4b54-bdb4-6fadf649a76d1752771985141.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9985157c-19fb-48f5-83c4-205b26dfb06e1752771985532.jpg"
    }
  },
  {
    "id": "8053",
    "province": "上海市",
    "city": "上海市",
    "name": "上海棕榈广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/814b98c0-38cd-4712-8ef7-6d54d8aa03321776704825709.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/37ef2eda-f093-410b-8c29-75662dfdd7351776704826069.jpg"
    }
  },
  {
    "id": "8116",
    "province": "上海市",
    "city": "上海市",
    "name": "上海歌斐中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fe5b6f83-65f7-481a-b725-0981628fd0581704215360693.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8ba712c7-9f05-486e-83de-3b5c2966685b1704215361573.jpg"
    }
  },
  {
    "id": "8080",
    "province": "上海市",
    "city": "上海市",
    "name": "上海汇锦南翔中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7e1c0ed8-7245-4c89-a5b8-3449e64984d21662621835754.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b358709e-3a0a-48e5-bc13-46c288c6b37b1662621836128.jpg"
    }
  },
  {
    "id": "8216",
    "province": "上海市",
    "city": "上海市",
    "name": "上海浦江生活广场店",
    "licenses": {
      "business": "",
      "food": ""
    }
  },
  {
    "id": "8099",
    "province": "上海市",
    "city": "上海市",
    "name": "上海淮海东路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a6423d3e-5358-42ca-be87-0fdf082f555a1704215351595.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f8dd543d-5c75-4fb9-b9f2-911829bac9e91704215352022.jpg"
    }
  },
  {
    "id": "8081",
    "province": "上海市",
    "city": "上海市",
    "name": "上海淮海路巴黎春天店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/675a5bef-00de-4325-9fcd-89013ad152561704215341639.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3148bedf-61d1-4726-8da1-94abcb6527551704215342066.jpg"
    }
  },
  {
    "id": "8106",
    "province": "上海市",
    "city": "上海市",
    "name": "上海漕河泾本部园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ec056f58-c448-4307-b370-0b791e38a3151668013650523.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c2df7690-dd0e-4add-aaff-4cae0c0b05401668013650853.jpg"
    }
  },
  {
    "id": "8070",
    "province": "上海市",
    "city": "上海市",
    "name": "上海火车站友谊商场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/23b97117-5276-434a-a74f-d4d182e704f91662621825916.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/17506d41-e886-4e7a-8a5d-d8443f8088a91662621826318.jpg"
    }
  },
  {
    "id": "8002",
    "province": "上海市",
    "city": "上海市",
    "name": "上海环球金融中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/368181c6-3db5-434a-af71-b29af9a147ea1727197559785.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b17afe97-9cf7-4d89-8508-b1cf7c1220131727197560285.jpg"
    }
  },
  {
    "id": "8086",
    "province": "上海市",
    "city": "上海市",
    "name": "上海田尚坊店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6c0184c9-52e9-4bd3-8956-237d3bec1ffa1662621842226.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/18c1ec09-d6d4-42b6-b2a3-1b4cc877b2de1662621842526.jpg"
    }
  },
  {
    "id": "8011",
    "province": "上海市",
    "city": "上海市",
    "name": "上海真北路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/77c79e4e-8a65-42c7-ac5a-0ab1b1e2f9e91742490350663.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/acdfee6c-dacd-44ef-bd25-76db04aa97cd1742490351254.jpg"
    }
  },
  {
    "id": "8063",
    "province": "上海市",
    "city": "上海市",
    "name": "上海科汇大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/76bc2a2f-b20d-4f7d-a59d-c1e0c41129811662621818312.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b903f3c5-e90e-42bd-a700-94bcf8bdde491662621818675.jpg"
    }
  },
  {
    "id": "8055",
    "province": "上海市",
    "city": "上海市",
    "name": "上海维璟印象城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/52985b86-6ac4-4984-be0f-86ba82a592791775236017117.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/21777259483005.pdf"
    }
  },
  {
    "id": "8048",
    "province": "上海市",
    "city": "上海市",
    "name": "上海虹桥世界中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/00085845-8bbc-456a-a47d-b2d1295e8b3c1662621802587.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品经营许可证1778031421270.png"
    }
  },
  {
    "id": "8149",
    "province": "上海市",
    "city": "上海市",
    "name": "上海虹桥火车站二店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ad945f9c-022d-4882-8cad-d5e1c8af72a41696871098042.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/59d44013-a9ac-4b50-af8e-3a385bf279401696871119580.jpg"
    }
  },
  {
    "id": "8026",
    "province": "上海市",
    "city": "上海市",
    "name": "上海虹桥火车站店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/247aeb15-b7fc-4d24-8a0c-2e4f278209481767546430502.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b821c01b-9c5d-4aea-85d3-7003eff647a11767546430824.jpg"
    }
  },
  {
    "id": "8072",
    "province": "上海市",
    "city": "上海市",
    "name": "上海越秀大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/303e6d7a-050d-4805-b3f5-967ec9190bdc1662621827160.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a79d5a75-c4e5-4656-b238-8cd01ca0d1bf1662621827458.jpg"
    }
  },
  {
    "id": "8156",
    "province": "上海市",
    "city": "上海市",
    "name": "上海金山万达广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/547d8e07-3cef-48ca-a43e-7f21f0c09a4d1699290361417.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4f6b7ff7-8360-424b-95c5-2d38ca75b7481699290385183.jpg"
    }
  },
  {
    "id": "8126",
    "province": "上海市",
    "city": "上海市",
    "name": "上海金桥太茂店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/72c5ff52-e6a7-41eb-b4a1-005b976d6ab71689267746509.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b37b53a7-95f2-4ce0-bd6f-fe3e96b4eb741689267747958.jpg"
    }
  },
  {
    "id": "8114",
    "province": "上海市",
    "city": "上海市",
    "name": "上海金汇天街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/80e17ba0-e06f-42fe-86da-d6cc1c072b5c1670432492977.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f965654d-5072-40e2-82d6-cda2966d22c01670432493287.jpg"
    }
  },
  {
    "id": "8082",
    "province": "上海市",
    "city": "上海市",
    "name": "上海金虹桥国际中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5f3a978f-7ea9-4556-83c2-45a7edbbc4321662621838509.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9bf371dc-8be6-4398-a151-77952e77caeb1662621838855.jpg"
    }
  },
  {
    "id": "8059",
    "province": "上海市",
    "city": "上海市",
    "name": "上海长宁国际中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/27e2d9f1-23e4-4101-876c-4e02580a40871778951235936.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d6a0ed5e-2576-465a-901a-322a9699047d1778951236771.jpg"
    }
  },
  {
    "id": "8120",
    "province": "上海市",
    "city": "上海市",
    "name": "上海陆家嘴广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6b38c15c-5110-4429-8dd3-7b6bc0a438b61689267739321.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/141e9af7-b19a-4b3f-ad2d-aeae8f3ff32a1689267740438.jpg"
    }
  },
  {
    "id": "8022",
    "province": "上海市",
    "city": "上海市",
    "name": "上海陆家嘴软件园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6bf85acf-3888-4396-a54d-f22931afb4801753722369881.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3d44143c-9c1a-462b-8601-71c19a3208b41753722370248.jpg"
    }
  },
  {
    "id": "8006",
    "province": "上海市",
    "city": "上海市",
    "name": "上海集电港路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/30890704-c90b-48fc-9fe3-ebbaaa83025a1727715958983.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/287fb0d9-6af3-426f-9dbc-26eb6b7d9d151727715959449.jpg"
    }
  },
  {
    "id": "8073",
    "province": "上海市",
    "city": "上海市",
    "name": "上海鲁能国际中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8858ac13-1a58-4617-b302-50c3b4e462ba1662621828380.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c79a7063-f06b-47d0-abf9-89cc651613761662621828739.jpg"
    }
  },
  {
    "id": "B012",
    "province": "北京市",
    "city": "北京市",
    "name": "北京上地华联店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/69beb99c-a043-4ab4-91fe-aa5d37c0c25b1662621895104.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8e2fcfe9-ece9-4b81-8a4a-5423210194f51662621895440.jpg"
    }
  },
  {
    "id": "B001",
    "province": "北京市",
    "city": "北京市",
    "name": "北京五道口店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f927cfd3-7707-4c86-9de3-0ac584f0c3031774458460042.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6ab7053b-09ef-414c-9edd-2c22eef7a5d71774458460448.jpg"
    }
  },
  {
    "id": "B004",
    "province": "北京市",
    "city": "北京市",
    "name": "北京呼家楼店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e76e0ee4-a859-444c-8872-b7aa3851e84b1662621884887.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0346f73e-afe9-4726-86c8-0a65d3c971511662621885215.jpg"
    }
  },
  {
    "id": "B003",
    "province": "北京市",
    "city": "北京市",
    "name": "北京太阳宫凯德MALL店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0075a279-0dea-4824-977f-73416e031f1c1779728897339.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8a10937b-58c1-4903-a338-9008400c3bc31779728899592.pdf"
    }
  },
  {
    "id": "B031",
    "province": "北京市",
    "city": "北京市",
    "name": "北京广安路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/21df2b5f-d180-4278-bed2-f133c04969471704215383080.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5f080f6c-28e8-46be-92b6-ecc6f5d0c21a1704215385653.pdf"
    }
  },
  {
    "id": "B019",
    "province": "北京市",
    "city": "北京市",
    "name": "北京总部基地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2de84c47-0dc6-4308-8560-b1c7505408881672074172513.jpg",
      "food": ""
    }
  },
  {
    "id": "B005",
    "province": "北京市",
    "city": "北京市",
    "name": "北京新中关购物中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6663666d-c926-4cdb-8da3-07b78aadfe1d1662621886142.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0240db1f-9d8b-45f2-88f0-aa216734c0891662621886468.jpg"
    }
  },
  {
    "id": "B014",
    "province": "北京市",
    "city": "北京市",
    "name": "北京新辰里酒仙桥店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/18e353df-9f09-4eb1-bedd-e07b55e84eb91662621897450.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/68588b0d-b299-4b02-ba15-d1b31771e3131662621897777.jpg"
    }
  },
  {
    "id": "B030",
    "province": "北京市",
    "city": "北京市",
    "name": "北京方庄时代店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9e317d27-8780-426a-98ec-76574578f9c31701968851171.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ad6e6c04-733a-4298-ae10-311ada8f76ed1701968851857.pdf"
    }
  },
  {
    "id": "B027",
    "province": "北京市",
    "city": "北京市",
    "name": "北京望京东园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/73a4d05d-9d6d-428a-968b-a438dcc69e751698858250902.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f5f0f31e-09cd-43bd-9461-5975bbd94f071698858251166.png"
    }
  },
  {
    "id": "B020",
    "province": "北京市",
    "city": "北京市",
    "name": "北京石景山万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0b8fb734-aebd-45be-bf54-01afabd659f01668877283345.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8946a7e1-5d07-451d-8325-27b26d5f5fbe1668877283653.jpg"
    }
  },
  {
    "id": "B008",
    "province": "北京市",
    "city": "北京市",
    "name": "北京融科天地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/358909f7-8960-4b4c-a695-28076ab6e8cc1662621889681.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/019daa4a-3059-4905-9abc-3549cc59cbcd1662621890000.jpg"
    }
  },
  {
    "id": "B015",
    "province": "北京市",
    "city": "北京市",
    "name": "北京通州万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0c4379ca-f26b-4f7f-9d85-b0f34a96d4141662621898738.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/248fff2d-a7b0-46bf-b0b9-d69b3c3d7e2f1662621899167.png"
    }
  },
  {
    "id": "B029",
    "province": "北京市",
    "city": "北京市",
    "name": "北京金唐新光界店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/73a4d05d-9d6d-428a-968b-a438dcc69e751698858250902.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f5f0f31e-09cd-43bd-9461-5975bbd94f071698858251166.png"
    }
  },
  {
    "id": "B013",
    "province": "北京市",
    "city": "北京市",
    "name": "北京金地中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/96aa8f48-84bb-455f-a354-add672f49f1e1662621896283.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/31a009ba-c4dd-4a4c-89a0-9ae310b069381662621896593.jpg"
    }
  },
  {
    "id": "B016",
    "province": "北京市",
    "city": "北京市",
    "name": "北京龙湖长楹天街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f8c7fdfa-1d58-460c-8c64-c903492eb8d41662621900021.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b8212e45-890c-4902-95a5-bfb3061d3c501662621900461.jpg"
    }
  },
  {
    "id": "1947",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州万达华府店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1947/1947-亳州万达华府店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1947/1947-亳州万达华府店-经营许可证.jpeg"
    }
  },
  {
    "id": "1564",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州上善名郡店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1564/1564-亳州上善名郡店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1564/1564-亳州上善名郡店-经营许可证.jpeg"
    }
  },
  {
    "id": "1678",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州亳芜家园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1678/1678-亳州亳芜家园店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1678/1678-亳州亳芜家园店-经营许可证.jpeg"
    }
  },
  {
    "id": "1793",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州依立腾奥莱店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1793/1793-亳州依立腾奥莱店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1793/1793-亳州依立腾奥莱店-经营许可证.jpeg"
    }
  },
  {
    "id": "1093",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州利辛七彩世界餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1093/1093-亳州利辛七彩世界餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1093/1093-亳州利辛七彩世界餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1075",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州利辛向阳路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1075/1075-亳州利辛向阳路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1075/1075-亳州利辛向阳路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1092",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州利辛淝河大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1092/1092-亳州利辛淝河大道餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1092/1092-亳州利辛淝河大道餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1527",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州华富广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1527/1527-亳州华富广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1527/1527-亳州华富广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1533",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州南半球餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1533/1533-亳州南半球餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1533/1533-亳州南半球餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1662",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州市万达广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1662/1662-亳州市万达广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1662/1662-亳州市万达广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1588",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州市缤纷城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1588/1588-亳州市缤纷城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1588/1588-亳州市缤纷城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1522",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州希夷大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1522/1522-亳州希夷大道餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1522/1522-亳州希夷大道餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1563",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州建安文化广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1563/1563-亳州建安文化广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1563/1563-亳州建安文化广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1826",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州建投东方名府店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1826/1826-亳州建投东方名府店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1826/1826-亳州建投东方名府店-经营许可证.jpeg"
    }
  },
  {
    "id": "1986",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州恒大城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1986/1986-亳州恒大城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1986/1986-亳州恒大城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1525",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州文帝路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1525/1525-亳州文帝路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1525/1525-亳州文帝路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1521",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州旺角广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1521/1521-亳州旺角广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1521/1521-亳州旺角广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1668",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州桐花路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1668/1668-亳州桐花路店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1668/1668-亳州桐花路店-经营许可证.jpeg"
    }
  },
  {
    "id": "1524",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州汇金广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1524/1524-亳州汇金广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1524/1524-亳州汇金广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1669",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州涡阳新街里店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1669/1669-亳州涡阳新街里店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1669/1669-亳州涡阳新街里店-经营许可证.png"
    }
  },
  {
    "id": "1520",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州涡阳站前路餐厅",
    "licenses": {
      "business": "",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1520/1520-亳州涡阳站前路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1850",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州涡阳青牛广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1850/1850-亳州涡阳青牛广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1850/1850-亳州涡阳青牛广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1609",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州百大购物广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1609/1609-亳州百大购物广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1609/1609-亳州百大购物广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1428",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州蒙城万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1428/1428-亳州蒙城万达广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1428/1428-亳州蒙城万达广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1683",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州蒙城喜客甄选店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1683/1683-亳州蒙城喜客甄选店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1683/1683-亳州蒙城喜客甄选店-经营许可证.jpg"
    }
  },
  {
    "id": "1519",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州蒙城天河广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1519/1519-亳州蒙城天河广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1519/1519-亳州蒙城天河广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1517",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州蒙城宝业餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1517/1517-亳州蒙城宝业餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1517/1517-亳州蒙城宝业餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1518",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州蒙城建材城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1518/1518-亳州蒙城建材城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1518/1518-亳州蒙城建材城餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1516",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州蒙城梦蝶广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1516/1516-亳州蒙城梦蝶广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1516/1516-亳州蒙城梦蝶广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1515",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州蒙城玖隆广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1515/1515-亳州蒙城玖隆广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1515/1515-亳州蒙城玖隆广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1667",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州观澜天下店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1667/1667-亳州观澜天下店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1667/1667-亳州观澜天下店-经营许可证.jpg"
    }
  },
  {
    "id": "1526",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州谯城万达餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1526/1526-亳州谯城万达餐厅-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1526/1526-亳州谯城万达餐厅-经营许可证.png"
    }
  },
  {
    "id": "1587",
    "province": "安徽省",
    "city": "亳州市",
    "name": "亳州谯陵路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1587/1587-亳州谯陵路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1587/1587-亳州谯陵路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "2579",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安万鼎银河湾店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/da3256a9-4ec5-437f-aeb7-fbfeff7713c51757697227837.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a470db66-cf38-4345-ab6e-60654e76674c1757697228210.jpg"
    }
  },
  {
    "id": "2291",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安三十铺金港店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6f129ef9-e266-44b6-ae32-3392622b22e91757696832583.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/09747619-8c51-4a47-ba80-48df3c6ee3a41757696833349.jpg"
    }
  },
  {
    "id": "2285",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安上城国际店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a4c133f8-644a-4a5f-8f84-aabfdb7790461777568519562.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b213ade5-2d85-40e4-8e26-0785978023b71777568520269.jpg"
    }
  },
  {
    "id": "2525",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安丽水康城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d4887d96-4fa0-43e7-a914-4bec51ab4f191757697156811.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f9efc93e-c44b-43b4-864c-cb60f9c1cba01757697157176.jpg"
    }
  },
  {
    "id": "2074",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安云路街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0f9eeca2-4ae7-40f7-ac62-0ea64db869091757696540372.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4488a6d1-541d-48b4-a3aa-fb9ba4882da41757696540731.jpg"
    }
  },
  {
    "id": "2287",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安人民南路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/599617eb-7f4d-458d-b681-ba3c1963b2371757696824137.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aa902bf9-0418-480a-860f-22c8b2c769651757696824873.jpg"
    }
  },
  {
    "id": "2700",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安公园华府店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f058c946-c154-4120-92c1-5a5fb045f7091763744628493.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eb62f198-eaa2-43c4-9980-ec3b61d253fe1763744629275.jpg"
    }
  },
  {
    "id": "2295",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安加油站店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d2aee443-4bc6-4ddb-b64c-e9d67e1cf4231757696838702.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/07816067-9b21-4ba1-a9ec-5e609c7c45301757696839097.jpg"
    }
  },
  {
    "id": "2497",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安华邦新华城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/226f770c-125b-411e-85f2-c71fd15236d61757697124440.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/91cde431-d447-4894-adb5-ef494f119de41757697125168.jpg"
    }
  },
  {
    "id": "2262",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安南河佳苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e7e5173d-ff46-4615-9734-89283eb294841766595704974.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b6990e2a-2516-4ac2-bf18-6e863a880ff71766595705676.jpg"
    }
  },
  {
    "id": "2441",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安和谐名城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2c5c4786-c121-49bb-96ef-0ed93b1fa0631757697052662.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d04cfb42-3930-42a9-9969-37b442f1cb581757697053066.jpg"
    }
  },
  {
    "id": "2538",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安嘉利豪庭店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/57c02d5e-ff3a-4f96-b06f-d5582b7a94f91757697183405.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f2d74268-0409-4ad7-aa5a-861c375828691757697183777.jpg"
    }
  },
  {
    "id": "2753",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安大润发店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8caae9bc-d6f7-41a6-97d1-365c8679dc0c1757697402238.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5268ca2f-69ea-45d9-9e92-4cf0a0cf02dd1757697402600.jpg"
    }
  },
  {
    "id": "2244",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安天盈大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cec7e12c-67cb-4c69-bc71-3b6df75a1c491757696755451.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d6fdd26e-46ef-47b3-8d4c-a933366266951757696756135.jpg"
    }
  },
  {
    "id": "1670",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安姚李镇庆丰购物广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1670/1670-六安姚李镇庆丰购物广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1670/1670-六安姚李镇庆丰购物广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1885",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安市独山镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1885/1885-六安市独山镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1885/1885-六安市独山镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2211",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安市皖西大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b7167b54-4679-4cbe-8bf9-8263c57410561757696713015.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/188029f9-9271-43d8-aef3-3a27c98275b51757696713393.jpg"
    }
  },
  {
    "id": "2470",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安恒大御景湾店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7b6227e8-fed7-45c7-bf29-1ac0dbe446871756832813094.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b840a560-1464-496d-86d6-ab044b3706171756832813425.jpg"
    }
  },
  {
    "id": "2385",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安新加坡御苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/05066ab1-0fd7-43cb-9fa2-fe10b81e3c9c1757696977855.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8a63f81c-b35a-4006-80ef-d889159de7841757696978234.jpg"
    }
  },
  {
    "id": "2597",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安新城春天店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4b69194d-6f70-4cbb-a66d-5bb9df67a7c81757697250555.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fe9c6337-aef0-4e68-b36f-3f86eb68f1fd1757697251256.jpg"
    }
  },
  {
    "id": "1739",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安新安镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1739/1739-六安新安镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1739/1739-六安新安镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2311",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安星汇苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/504c5e1b-2f49-4e90-9f0a-dd265ddf42331757696865648.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/38c677d9-9990-4475-8ed9-608ae8c808541757696866453.jpg"
    }
  },
  {
    "id": "2016",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安梅山路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bc2fd37a-33d6-48c7-9cbb-7dd81f9ea9571757696452840.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1391858e-c43b-45e5-abac-69073fd24e701757696453531.jpg"
    }
  },
  {
    "id": "2776",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安正东凯旋店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5ce9ae40-631d-4d55-865b-18937b0d3d051757697439861.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/84098d9e-d195-4341-912a-5fe51912042a1757697440215.jpg"
    }
  },
  {
    "id": "2250",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安浙东商贸城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e75310f2-169c-41fc-95b6-2c94a6f004491757696769502.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fff98c7c-3d4c-4ab1-a4ad-c6b82ca49b951757696770157.jpg"
    }
  },
  {
    "id": "2366",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安海心沙店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/989b5608-e273-4757-8545-618182ca55c61757696944647.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/599b7dad-c88d-4f6e-8b26-218bf647eb741757696945518.jpg"
    }
  },
  {
    "id": "2023",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安皖西路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/471b840c-60ac-4b1b-92a5-e2b951634d261757696463315.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e636ee4c-904b-4938-989c-9a729e07f52c1757696463659.jpg"
    }
  },
  {
    "id": "2347",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安碧桂园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ed6f4f51-e95b-49fb-b6a7-a68f7dfd87a51757696914670.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/15ca53c2-6679-40e5-b356-5878f904641f1757696915060.jpg"
    }
  },
  {
    "id": "2173",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安红达广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/82c6d0b3-1fbe-4c48-b3b5-30979942e0e51757696667007.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c2ee3dd3-ee82-413d-880d-c9746ca72b561757696667718.jpg"
    }
  },
  {
    "id": "2462",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安红达星河城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6550a481-6692-4dc8-add5-aa72e63b459f1757697086664.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/07f29f0e-fa2e-4052-b169-bc37c424bb9e1757697087027.jpg"
    }
  },
  {
    "id": "2230",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安经济开发区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d789171a-4efa-4ed4-8e31-57d2deba76c91757696732043.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5656e2e2-2ffa-43a0-939a-cdac9aececd81757696732741.jpg"
    }
  },
  {
    "id": "2345",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安舒城南溪丽城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b1583a4c-b6b0-4cf2-80de-213aaf5ecc971757696911656.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/04b2505e-32c2-4f76-913d-bbdf1b4f975c1757696912282.jpg"
    }
  },
  {
    "id": "2025",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安舒城古城路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6847632f-1de3-4f3c-ad3c-4f4abf65d2131757696466232.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/98c8aad9-e5f0-47d3-b4c5-424ceee2e3171757696466581.jpg"
    }
  },
  {
    "id": "2153",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安舒城文化广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7487a9d9-9383-4f61-99f2-36a92bfc98d41778950875067.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3e3a74a6-a80a-4d76-8b82-d192da84fb0b1778950875859.jpg"
    }
  },
  {
    "id": "1585",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安舒城杭埠镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1585/1585-六安舒城杭埠镇店-营业执照.mp4",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1585/1585-六安舒城杭埠镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2032",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安舒城梅河路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6e975884-624f-4414-aed0-9b07b9f078961774544436715.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9bdce22b-68e1-4042-a669-5432e854d78f1774544437099.jpg"
    }
  },
  {
    "id": "2399",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安舒城港汇广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/567d55c6-f953-4cd9-b945-1b37f6b013831780160566696.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ced6711a-b99b-4527-b052-7d385269bfa81780160567825.jpg"
    }
  },
  {
    "id": "2745",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安西商农贸城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ee64b2ee-35a5-4373-b73e-4f50415b90d51757697394862.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/23fc735d-c28a-44ef-98fd-aec63ade78981757697395589.jpg"
    }
  },
  {
    "id": "2774",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安金寨一中店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c26c965d-9b4d-46d7-83a1-6bce3719aa191757697435729.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b542381a-a733-447c-9380-2fccc147fa041757697436562.jpg"
    }
  },
  {
    "id": "2383",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安金寨县红军大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d1f6c71e-5452-42e8-b9fc-eb4a9036ab6e1757696975289.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/57369e62-be81-453b-b581-b2e4a18cdf731757696975656.jpg"
    }
  },
  {
    "id": "2342",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安阳光欧洲城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c2aed4fd-a51e-4f4a-9f8d-19fa8e3313991757696903541.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1a8686a0-3a0b-4d11-ab82-4e7416d9bd9a1757696904239.jpg"
    }
  },
  {
    "id": "2807",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安霍山学府金街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bb85f48e-b206-4976-9601-1ab9404012e01757697501820.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/37abbde7-cbd8-4f95-8d0b-7b04b83c236f1757697502218.jpg"
    }
  },
  {
    "id": "2170",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安霍山迎驾大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f36e62eb-04c9-4571-87b7-038a6fe62ab11757696661803.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/912bd4e3-933d-4734-8c92-415e1fb9efe81757696662471.jpg"
    }
  },
  {
    "id": "2498",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安霍山鑫港嘉园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/351b5955-3f06-41d0-bd61-b9686005ad681757697127347.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e33e5295-2ddd-4fe8-97b0-c9e39e3e31e81757697127718.jpg"
    }
  },
  {
    "id": "1710",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安霍邱商之都店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1710/1710-六安霍邱商之都店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1710/1710-六安霍邱商之都店-经营许可证.jpg"
    }
  },
  {
    "id": "2756",
    "province": "安徽省",
    "city": "六安市",
    "name": "六安龙湖山庄店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4b495ecc-9c3f-441f-90f7-c19fadad87311757697405198.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ad7872ef-0ca2-49cf-918a-facde1a0b0951757697405567.jpg"
    }
  },
  {
    "id": "2143",
    "province": "安徽省",
    "city": "六安市",
    "name": "兴皖路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/00e96767-d889-4e81-a1db-7cc9830fbecb1757696622311.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/70c9f3a1-4e33-4998-8aba-bae43dc975ed1757696622718.jpg"
    }
  },
  {
    "id": "2130",
    "province": "安徽省",
    "city": "六安市",
    "name": "叶集店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3851c44e-0f62-4815-9d46-5816d2b7a8071757696605448.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/38290c44-0ffe-42ea-aa32-3e085025ba251757696605839.jpg"
    }
  },
  {
    "id": "2435",
    "province": "安徽省",
    "city": "六安市",
    "name": "寿春小区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8e9b3313-1121-4de7-9bb2-a9829de58a591757697047734.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eb7fee93-6445-4c5b-b0f6-9c4aec85fd6b1757697048420.jpg"
    }
  },
  {
    "id": "2728",
    "province": "安徽省",
    "city": "六安市",
    "name": "舒城万达广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f2a21902-a8d0-4efa-b961-f0a4927165681757697373967.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2c1a2809-f0bd-4210-a171-de2c4d5e4e5c1757697374626.jpg"
    }
  },
  {
    "id": "2566",
    "province": "安徽省",
    "city": "六安市",
    "name": "舒城鼓楼北街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9ccca5ff-644a-4d15-885a-324c0eedd01e1757697221764.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4b8e79aa-c68d-4d7a-96b9-df48b64ccb831757697222130.jpg"
    }
  },
  {
    "id": "2185",
    "province": "安徽省",
    "city": "六安市",
    "name": "舒城龙津大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9b54be9e-8c94-4a9c-a294-46ab81c4eee81757696685692.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4500f19d-5b86-481e-b10e-3d443ecf89ab1757696686332.jpg"
    }
  },
  {
    "id": "2662",
    "province": "安徽省",
    "city": "六安市",
    "name": "金寨明发广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5d8b0822-32ef-47bf-a3db-d35007c101071757697316582.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3e5edc92-6f73-4a65-aa01-22eda244a9bb1757697316957.jpg"
    }
  },
  {
    "id": "2081",
    "province": "安徽省",
    "city": "六安市",
    "name": "霍山分公司",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/77cbb3d0-9bd4-495e-afa0-bc611c24ed4c1757696553826.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4eec7bc9-05f6-4136-bdb1-5ff1f0ffec7d1757696554187.jpg"
    }
  },
  {
    "id": "2288",
    "province": "安徽省",
    "city": "六安市",
    "name": "霍邱县卧阳路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ed9647d5-0b25-443e-8cff-e475cedb48a61757696826725.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f64fbd21-a4b6-42cf-a2e7-9a513364802d1757696827416.jpg"
    }
  },
  {
    "id": "2669",
    "province": "安徽省",
    "city": "六安市",
    "name": "霍邱锦绣新天地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a9ff41a7-51a2-4236-a203-5561de78bf7d1759079019548.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/25b7d8b1-1858-4a5c-905a-38a874cfc4ec1759079019950.jpg"
    }
  },
  {
    "id": "2167",
    "province": "安徽省",
    "city": "合肥市",
    "name": "七里香榭店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4d487a5a-7544-437d-9497-af35f0cc1cf61757696658522.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ed63e69c-0e9c-4ef8-95cd-72bb46a68dda1757696659170.jpg"
    }
  },
  {
    "id": "2309",
    "province": "安徽省",
    "city": "合肥市",
    "name": "万达文旅新城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fc6973a5-707c-406a-8e3d-368aafa3b2081757696860954.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c8a742ab-b164-44ff-9e8f-ff0b22ed1d201757696861615.jpg"
    }
  },
  {
    "id": "2414",
    "province": "安徽省",
    "city": "合肥市",
    "name": "万邻坊店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/af27feaf-bf22-46c8-9b3e-1f2d1443ee171757697023037.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/298558de-78b6-4309-917f-b2f03ac797571757697023402.jpg"
    }
  },
  {
    "id": "2047",
    "province": "安徽省",
    "city": "合肥市",
    "name": "三孝口店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/13844f56-8bcc-447f-8389-2cc645d4f0211757696503071.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a9ac225d-e016-48a6-802d-21005b8f8f6c1757696503431.jpg"
    }
  },
  {
    "id": "2090",
    "province": "安徽省",
    "city": "合肥市",
    "name": "中环城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bf628ada-56a4-4c9c-b3bd-fcf70575dc6c1757696561505.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/38fd6440-4165-479b-aa88-0b735edcdb141757696561878.jpg"
    }
  },
  {
    "id": "2053",
    "province": "安徽省",
    "city": "合肥市",
    "name": "乐活广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c6557354-d728-465e-ad2b-5353e5ab68221757696515969.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b55e5f3b-e2ab-4b9f-ac1c-66e4baa8d4c91757696516632.jpg"
    }
  },
  {
    "id": "2095",
    "province": "安徽省",
    "city": "合肥市",
    "name": "九华山路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6c0ff703-70f8-451b-b615-084aa9481b351757696566681.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/697f48b5-d117-40ed-98b1-f069d853c2031757696567042.jpg"
    }
  },
  {
    "id": "2124",
    "province": "安徽省",
    "city": "合肥市",
    "name": "习友路水岸茗都店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/45833cbb-34e4-48e4-9dad-912338f176c71770224465755.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/20613479-015e-4eff-9540-059293b1c7621770224466449.jpg"
    }
  },
  {
    "id": "2416",
    "province": "安徽省",
    "city": "合肥市",
    "name": "书香门第店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d7da1d12-3b82-4618-872d-e0150f3c20a41757697025915.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ac03ed90-015c-4037-9432-2e5b4d1ce0ff1757697026281.jpg"
    }
  },
  {
    "id": "2206",
    "province": "安徽省",
    "city": "合肥市",
    "name": "亳州路二店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e07bc208-9ffd-4d59-a8c4-6552ae7a21d61757696708050.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5dfcf7b8-51b3-473b-ab9b-5f081cab0aa41757696708441.jpg"
    }
  },
  {
    "id": "2165",
    "province": "安徽省",
    "city": "合肥市",
    "name": "信地城市广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aa836af2-b531-44a6-92dc-df9c02ab1fb91776272475848.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/92f6427f-7793-48c8-865c-5620bb87c9de1776272476490.jpg"
    }
  },
  {
    "id": "2004",
    "province": "安徽省",
    "city": "合肥市",
    "name": "南七分店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8d88582c-5c81-44bf-9f98-2bae270922e31757696428526.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b2877340-09dd-4da4-9547-0cd9df22193d1757696429213.jpg"
    }
  },
  {
    "id": "2741",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥万和广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dc622895-c4b9-40f5-bf63-2a69663641b41757697386476.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3b194b5a-f432-429b-b4e2-7d2341ea5c1c1757697387152.jpg"
    }
  },
  {
    "id": "2426",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥万科城市之光店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9b7eb8bb-ab9a-4a33-b1af-14dd42048d7f1757697038129.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d3dd92ea-c0df-4ce4-9434-069a2cdf91911757697038499.jpg"
    }
  },
  {
    "id": "2420",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥万科森林店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/144f9a96-ad67-4869-9009-5e33e69182b01757697034322.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1c7c24c8-bc50-4e6a-8995-1be28ef74b121757697034681.jpg"
    }
  },
  {
    "id": "2641",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥万科雅庭店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/44c788f1-b1b6-4d5b-af7e-50f89d5067761757697291724.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/639c430c-aecc-45fa-8c42-6e880850e79a1757697292509.jpg"
    }
  },
  {
    "id": "1963",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥下塘丰迪广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1963/1963-合肥下塘丰迪广场店-营业执照.jpeg",
      "food": ""
    }
  },
  {
    "id": "1299",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥下塘比亚迪店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1299/1299-合肥下塘比亚迪店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1299/1299-合肥下塘比亚迪店-经营许可证.jpeg"
    }
  },
  {
    "id": "2324",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥世纪荣廷店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bb679519-88f7-410e-b6bf-4f8cf3e4fa531757696886472.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8e8e3874-04a9-4292-ae5d-2efb52510c3b1757696887133.jpg"
    }
  },
  {
    "id": "2779",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥东海花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d5265d47-fa84-40af-a7a5-a6417608ff291757697447714.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f57c1c26-1544-416f-91b6-4f3d6a77304e1757697448445.jpg"
    }
  },
  {
    "id": "2406",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥中央城邦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a114353a-2b14-42a3-8d80-0ac83ed807e71757697004336.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/367de0fa-09d1-4163-ac48-c89b7d1ce8bd1757697004831.jpg"
    }
  },
  {
    "id": "2720",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥中安创谷店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a8d069a9-10de-430f-996f-424c67ef3ad71757697359653.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fb9db99c-3516-4e7e-9877-d59305a42ccd1757697360361.jpg"
    }
  },
  {
    "id": "2646",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥丰乐亭路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6149e1d9-f25c-4efd-a2d8-778916445c9b1757697300581.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e86a4d0d-249e-4c34-984b-b8aace881bd81757697300969.jpg"
    }
  },
  {
    "id": "1276",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥丰乐服务区东区餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1276/1276-合肥丰乐服务区东区餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1276/1276-合肥丰乐服务区东区餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1275",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥丰乐服务区西区餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1275/1275-合肥丰乐服务区西区餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1275/1275-合肥丰乐服务区西区餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "2006",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥临泉路分店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/414fe674-18cd-425c-b00a-ada5957a3f051757696433766.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1d9e6b40-2da6-4c8e-a8a5-739e1d1499db1757696434469.jpg"
    }
  },
  {
    "id": "2780",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥临湖社区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9582b462-d627-4140-a34b-067967f14e051757697450209.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2a360769-a0a2-4470-a97b-84b6edcbe9ce1757697450861.jpg"
    }
  },
  {
    "id": "2558",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥义井路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f5c7dab7-8436-441c-b719-9d794d27b8021757697215399.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/79260380-fcdc-4790-8ecf-78b874f84e9a1757697216078.jpg"
    }
  },
  {
    "id": "2536",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥乐城超市店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e8688e23-a5a0-4569-a5de-b94f6e965d191757697178316.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c213e2e3-a9c8-43e3-8af4-293c2034e0691757697178689.jpg"
    }
  },
  {
    "id": "1611",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥二里街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1611/1611-合肥二里街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1611/1611-合肥二里街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "2176",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥云顶雅苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b6027c3b-62e7-4fcb-bcfb-0c2081f37fd81757696670015.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7a043fcb-1b37-4c55-be50-6430afcc6fc31757696670393.jpg"
    }
  },
  {
    "id": "2770",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥京商商贸城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2fa5f8be-cc8e-4a36-bf1f-2b544b62efb31757697427400.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/23bcbe50-1a75-42f8-9fb7-bb24333875731757697427773.jpg"
    }
  },
  {
    "id": "2043",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥亳州路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d3368602-0d42-4389-9057-0431f9a6c19a1757696498117.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f8aa9bdb-d1d6-4c9f-a1ba-70be7a8a67e81757696498473.jpg"
    }
  },
  {
    "id": "1983",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥众兴服务区北店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1983/1983-合肥众兴服务区北店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1983/1983-合肥众兴服务区北店-经营许可证.jpeg"
    }
  },
  {
    "id": "1984",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥众兴服务区南店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1984/1984-合肥众兴服务区南店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1984/1984-合肥众兴服务区南店-经营许可证.jpeg"
    }
  },
  {
    "id": "2307",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥保利五月花店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a074e8d3-5d4e-4055-b65a-6b9c288a22531757696856020.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8c7164c6-a008-44be-ac0b-cbdd71f6f6911757696856697.jpg"
    }
  },
  {
    "id": "2403",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥信达天御店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4188d8e4-7966-401a-82cc-0d5c615688f41757697000816.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0c95244b-53ea-4bde-9d55-363cfbef3aae1757697001209.jpg"
    }
  },
  {
    "id": "2134",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥信达好第坊店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d47621f8-a27b-4860-9297-cf56d06593931769706065293.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/72926a29-78ed-40dc-aeea-8a7b780e7e211769706065710.jpg"
    }
  },
  {
    "id": "2803",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥元一柏庄店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f7011e96-0f3f-4df3-ad8f-6888ade9969c1757697497523.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/159d0a11-205a-44b5-8643-9078027dafe51757697497902.jpg"
    }
  },
  {
    "id": "2509",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥光明北部湾店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b92ca841-7300-4ff7-b05b-01c6ecc45b671757697138315.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bd023a35-adb7-4dc2-9bf9-4cdf1f46de6f1757697139157.jpg"
    }
  },
  {
    "id": "2060",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥兴园小区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5f8dd6f2-903a-41ba-a0e5-2d991aab94f91757696529830.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2f9384f5-820f-47b2-a2e7-d548643f5d1f1757696530207.jpg"
    }
  },
  {
    "id": "2671",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥创新大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/72465ed6-963d-4f93-9875-4e7b2dd8a7a81762448619475.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/59009fd2-85c7-44f4-a8c0-189cd16bdc001762448619800.jpg"
    }
  },
  {
    "id": "2487",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥利港银河广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/90fb53da-687c-450b-b3a4-4fbdd11ab2071757697114041.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a607c5d7-be91-4fd7-9ac5-89650777cf2b1757697114728.jpg"
    }
  },
  {
    "id": "2322",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥加侨悦山城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/69078a17-2ae4-43bc-b53f-a07c7589c0ba1757696883742.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e26c9f79-d597-472b-81a5-e6b0a0c0cec41757696884237.jpg"
    }
  },
  {
    "id": "2516",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥包河万达广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e7033c6c-1af2-4f94-b2ae-9089be8dbc821757697145850.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c6c597d9-c3eb-476e-8ab6-6f18d7afbc8b1757697146207.jpg"
    }
  },
  {
    "id": "2055",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥包河花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4f775c97-6ad7-4753-910f-aad4c289ca2c1757696519238.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f08d0339-bd6a-44f2-9551-1f7ba6ba7c851757696519618.jpg"
    }
  },
  {
    "id": "2164",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥包河苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ac736722-7b51-4157-8e9a-9ab2de9b7cef1757696651105.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a105b9cf-eb5f-4aa5-927c-787f7c8909701757696651466.jpg"
    }
  },
  {
    "id": "2737",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥北一环店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/298beb9e-6d15-4166-9b12-a40caa920b261757697383524.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e7dfd92b-f4e2-4163-b371-af08e05bcb501757697383888.jpg"
    }
  },
  {
    "id": "2757",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥北城万科公园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8e820a92-ca51-4f96-a613-04620db6ca271757697407562.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/78002607-a7f9-4ce0-ace6-d72a8120fe7f1757697408239.jpg"
    }
  },
  {
    "id": "2199",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥北城世纪城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2c99fe54-f7dd-4c58-8285-c44b1642c7b91757696699404.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ca7dc001-734c-4aa0-8b12-9807f49733da1757696700096.jpg"
    }
  },
  {
    "id": "2177",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥十里庙店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/13ca3943-3ac8-478e-ad29-015fb634cd5e1757696672827.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8b050540-c56c-455c-8ce0-160b96b4c71e1757696673196.jpg"
    }
  },
  {
    "id": "2535",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥华冶新天地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5714135a-6743-463f-852c-662612448a831757697175786.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ee5ce0ca-3218-49ba-82e2-e5c43222018f1757697176142.jpg"
    }
  },
  {
    "id": "2186",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥华地新街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c9c6ab92-d8d6-4b4a-89df-e2b17f2422e41757696688155.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d3ba9c9d-4355-47db-a9a5-a173c7933aea1757696688801.jpg"
    }
  },
  {
    "id": "2556",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥华润万象汇店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eb443ec4-4cf6-4832-b042-4817d2f0cf661757697212789.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2c666b2e-8984-43f2-9610-a567551741b01757697213488.jpg"
    }
  },
  {
    "id": "2413",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥华润桃源里店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4f370daa-82b7-4479-9d7a-48395a9c5cc21757697020084.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f1cbb1c5-b6bb-4e70-8682-104243ec39081757697020452.jpg"
    }
  },
  {
    "id": "2242",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥南岗店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fc8895c4-b970-4dfb-9019-e1ee932eb0c51757696752526.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/539efdea-993c-40ee-acf8-c4062a4268e21757696753209.jpg"
    }
  },
  {
    "id": "2582",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥南湖春城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f7da3eef-6bfd-4d28-8306-27e4db7873121757697233041.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f519f93d-6995-48f0-898e-fbcf9c6540bf1757697233410.jpg"
    }
  },
  {
    "id": "2642",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥南翔汽车城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/63ab2c08-aead-4fdc-8b73-663e8d65ff831757697294288.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/06b5efe0-ac39-4bf9-9026-10986f391fbb1757697295072.jpg"
    }
  },
  {
    "id": "2643",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥双凤里小区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/165a15bf-5d98-4000-9c67-e77e07e3f4d61757697297287.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/760dc381-c4b3-4a15-8216-10c9ffcff5a71757697297660.jpg"
    }
  },
  {
    "id": "2012",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥合作化北路大润发店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aba08afb-f69c-45a3-9ac7-6020d37f95091757696447117.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/be1d2f94-e722-4c98-8c5a-e4d6356b00d61757696447470.jpg"
    }
  },
  {
    "id": "2625",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥合郢花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/beb4bef7-c520-415b-acfb-936d84572bca1757697270858.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c5382ded-653f-4b2d-b12a-d02a8f14fe831757697271215.jpg"
    }
  },
  {
    "id": "2456",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥同创科技园店（五代）",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/98cc3820-d591-4db8-b7cd-b5bdc7d598ea1757697077818.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7d8d73f1-b462-4c07-85ab-9a0fc60fd8201757697078195.jpg"
    }
  },
  {
    "id": "2302",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥名门湖畔店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/68a545d6-f35e-42d1-b90f-c23d96b5bc7c1757696850298.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/47137f22-0ce4-4fec-930e-a88685486ddb1757696850663.jpg"
    }
  },
  {
    "id": "2409",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥君御世家店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/afbcbf28-4a0b-47b0-9bc1-15766a2578a31757697012294.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a1474063-0794-45d9-b921-8665d871576a1757697012661.jpg"
    }
  },
  {
    "id": "1737",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥商之都优山美地店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1737/1737-合肥商之都优山美地店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1737/1737-合肥商之都优山美地店-经营许可证.jpeg"
    }
  },
  {
    "id": "2778",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥园博小镇店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e9ba2144-2687-484b-8c96-2b6af0c81ae81757697445263.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f15fe7e2-b05f-49ce-b777-07228e43f4541757697445630.jpg"
    }
  },
  {
    "id": "2252",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥国耀星达城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9f2e313f-8ff9-4aee-b585-47829940bd221757696772407.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7649c306-ada8-4a47-966c-aab7e7ce0c4c1757696773063.jpg"
    }
  },
  {
    "id": "2650",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥城隍庙店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f251c5d4-994f-43d0-a0be-979611bac53d1757697306205.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a0b38984-3a3e-4710-8c94-e97069cf2b181757697306555.jpg"
    }
  },
  {
    "id": "2677",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥大众路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/103d4a4f-26de-425c-8966-a9a7bc2425441764349421221.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/26117678-b86e-4e01-883b-6114b932b3591764349421905.jpg"
    }
  },
  {
    "id": "2614",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥大融城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8a8161dc-ef74-4500-b100-3dc782c70e231757697264123.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8241fc0e-7401-4e64-8a9b-13ec976b73f01757697264505.jpg"
    }
  },
  {
    "id": "2427",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥天下锦城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/25f2a809-14f3-4ac7-9365-dc534411db241757697040359.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/92732a2e-90ab-4a39-b780-d1fa81178b5b1757697041047.jpg"
    }
  },
  {
    "id": "2393",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥天玥中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ea598fdc-b6a6-4508-89a2-5548259504f81757696989267.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e8b7c95b-571b-439d-83d9-23260fd2b1f91757696989638.jpg"
    }
  },
  {
    "id": "2461",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥天珑广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/39858e94-7358-48de-a5d5-35a4c8e015181757697083833.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/853b358c-91fe-46ba-b74c-1c8e8d8e89bd1757697084504.jpg"
    }
  },
  {
    "id": "2357",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥天街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/708ed12d-788c-4d7d-9f05-d9fb72a88dd91757696930558.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/81eee266-8e2c-40e2-867d-be516a8591531757696931232.jpg"
    }
  },
  {
    "id": "2555",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥天长路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b496b3ac-6423-4010-8031-a4e6ebf4cb841757697210387.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7f3c7f8c-7e2f-4baf-8c26-cf317dede6a51757697210783.jpg"
    }
  },
  {
    "id": "2806",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥天鹅湖万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5d0ec925-a38a-46c3-8114-ddbc41cbe2781757697500567.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/244e23ab-f569-4003-9315-acb225d685091757697500969.jpg"
    }
  },
  {
    "id": "2598",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥太平洋广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/14c93806-0161-4df1-954a-81a15df220991757697253457.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9032073b-abf8-4506-b6a7-599d0e6377c71757697253822.jpg"
    }
  },
  {
    "id": "2532",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥奥园城市天地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6f30af36-70ba-4930-9522-4c102b41ebe41757697172455.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f4bbea90-724d-4d2a-a30f-331dc970e3d41757697172821.jpg"
    }
  },
  {
    "id": "2315",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥始信路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/719c022a-f61b-4697-9ffa-5329932f67721757696872040.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3a2e21f7-2add-4561-89cd-3457289f6c6a1757696872387.jpg"
    }
  },
  {
    "id": "2712",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥安医一附院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b188c221-b940-4cee-90e1-9754be11e1f61770656626783.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1300b94c-ebb3-4a51-a9c0-64b072429fd31770656627541.jpg"
    }
  },
  {
    "id": "2443",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥安医二附院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/49882cda-fc4f-48b3-8fdc-8899909b33521757697057830.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c2bee3aa-b6fc-42c4-bcbb-76bae7c943851757697058192.jpg"
    }
  },
  {
    "id": "2892",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥安大馨苑校区店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/营业执照1779705416760.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品经营许可证1779705432040.jpg"
    }
  },
  {
    "id": "2527",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥安通广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b9321941-cf1b-48cd-bc5d-8713f246b8bc1757697159393.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/41f13f29-2125-4e9d-a8ac-4997fa4432641757697159761.jpg"
    }
  },
  {
    "id": "2647",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥宝文花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/842f225a-764a-4aa9-9c6c-5c002a2c1b871757697303489.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7e0eb2e7-ae0e-41ab-9491-619127c4e3031757697303983.jpg"
    }
  },
  {
    "id": "1610",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥宿州路商之都餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1610/1610-合肥宿州路商之都餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1610/1610-合肥宿州路商之都餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "2495",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥宿州路新店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3778fb8e-5044-4d6a-bd97-dc08271416861757697121253.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8cb616aa-a701-4b47-b75e-103e5bfd6ca61757697122360.jpg"
    }
  },
  {
    "id": "2373",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥小庙振兴路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bbadc998-f9d0-454d-b173-7c3f406703031757696956234.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/10a0a1bd-cf12-4f23-8a19-4bc8dded51af1757696956932.jpg"
    }
  },
  {
    "id": "2554",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥尚泽大都会店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fa1349a6-39df-4dc8-95b1-a4cf1a2e8a611757697207531.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/655bcee2-f4b2-41e9-acc3-0e85bef167561757697208190.jpg"
    }
  },
  {
    "id": "2222",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥岸上玫瑰店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/36b82d8d-1ba0-42eb-bfb9-6173b35728d31757696723973.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a24afa88-2214-4045-8e99-f8992f1eea101757696724660.jpg"
    }
  },
  {
    "id": "2455",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥巢湖万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/01873500-4fc7-47ef-942f-a173be2ed9381757697074962.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/80c8363c-9c18-408a-bb6d-86cf8b0e21061757697075662.jpg"
    }
  },
  {
    "id": "2758",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥巢湖健康中路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0df8f2e4-31d7-4db2-9516-ad91cd6063661757697410446.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4790b157-627c-4e7d-8964-7a08a8118d741757697410835.jpg"
    }
  },
  {
    "id": "1388",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥巢湖柘皋镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1388/1388-合肥巢湖柘皋镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1388/1388-合肥巢湖柘皋镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1329",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥巢湖槐林镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1329/1329-合肥巢湖槐林镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1329/1329-合肥巢湖槐林镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2626",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥巴黎春天店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/745c17f1-c877-42e5-b643-d0037139a2681757697273098.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a2e46ae7-5352-473d-955e-1bb108b337a21757697273767.jpg"
    }
  },
  {
    "id": "2369",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市保利广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/617866e8-8057-4f50-bc79-b0f5a65c15191757696952507.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bc21bf6f-da9c-4446-b39d-a0bb08b784f61757696953188.jpg"
    }
  },
  {
    "id": "2349",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市创新产业园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6b6b2e77-c4a8-44d5-97fc-066930fe36961757696917546.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ecd99d8b-a7e1-4d22-bd41-761739a76cf11757696918318.jpg"
    }
  },
  {
    "id": "2374",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市圣地雅阁店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3b0a46cf-00ec-412b-8294-9bf222b94b1f1757696959083.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/550520c2-ea88-4dd6-ae96-f046d1fcc7b41757696959464.jpg"
    }
  },
  {
    "id": "2442",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市广视花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/69834f9f-5bc6-4eeb-8231-bcc5853aed8f1757697055320.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0b3a3ed9-7350-4b08-8abc-ddc734c6458c1757697055702.jpg"
    }
  },
  {
    "id": "2530",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市恒泰城果店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cd281fca-9934-4ed5-82cb-22d31208c1291757697167317.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/27460c92-dbb3-49f1-9303-bcd61673c90c1757697167693.jpg"
    }
  },
  {
    "id": "2365",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市湖滨公馆店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4a901403-f409-43b3-bc92-1d5cd11601d91757696941131.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/35db5692-6307-423d-86ac-88f8831ce52f1757696942088.jpg"
    }
  },
  {
    "id": "2354",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市空港新城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ef12074a-c991-4fe3-b777-bb3b3a5231c51757696925499.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d7604848-2761-481d-85b6-b1a0b427e3811757696925836.jpg"
    }
  },
  {
    "id": "2551",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥市金梅路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1135396d-c940-43fd-825f-bb6fdd09c71a1757697202277.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/19cef16d-ffeb-4bcb-9840-153a804de3f21757697202648.jpg"
    }
  },
  {
    "id": "2781",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥广乐家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bdd87347-1521-45d0-b91e-2997a12c4f621757697453361.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/62e0ecf4-9067-4f34-a402-3e716a618eec1757697454235.jpg"
    }
  },
  {
    "id": "2253",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥庐江分公司店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e533cb38-906b-41a2-b714-f458288e54121767114097239.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2513e637-a3b3-4d09-8434-d049aab86c401767114098291.jpg"
    }
  },
  {
    "id": "2539",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥庐江方圆荟店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/41fb11e8-a3f5-4356-b5d3-138be6d392011757697185838.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/370f086b-b193-475b-a9f2-c4c763de25111757697186537.jpg"
    }
  },
  {
    "id": "2129",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥庐江路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/545ef001-3335-4f45-b8b1-8fc078cdf0811757696602918.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/503a1d7f-9736-4e82-bc03-5a523227712a1757696603271.jpg"
    }
  },
  {
    "id": "2788",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥弘盛广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6c8bd67f-05e7-41da-9a79-789c3ca2fe381757697469217.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/328a45de-2aa8-4fe9-a4d1-668c161fe35f1757697469604.jpg"
    }
  },
  {
    "id": "2722",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥弘阳广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cad70373-019d-4179-93a3-4e90991644861775754231441.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7a144bf7-24fc-43f2-bfcb-6feda0aad99b1775754232086.jpg"
    }
  },
  {
    "id": "2531",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥御园华府店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2420ded5-d236-4a34-96bb-e235d215acc61757697169928.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/931cead8-5a25-4280-b5af-8f7a8cea05991757697170297.jpg"
    }
  },
  {
    "id": "2445",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥御景城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fb5d5ba0-5a33-4ae5-a780-32898316ae371757697060466.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9c4dcad1-1b5e-4f80-92c8-ad959ce8964b1757697061176.jpg"
    }
  },
  {
    "id": "2631",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥徽州大道苏果店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/827f2a78-1f4d-4efa-a62b-619101c555c71757697278549.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4685f121-3f73-4344-bda3-cf874be93c5e1757697279246.jpg"
    }
  },
  {
    "id": "2560",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥恒大中央广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/29b8ba0f-042b-4804-adc7-bdd3b9c356e21757697219108.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/acd17a1b-af78-4100-a9df-ab235e2f4b821757697219494.jpg"
    }
  },
  {
    "id": "2635",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥恒大华府店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/147143f3-81f9-4b45-a61f-420479d39b731757697281168.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c19a1d54-0a17-4f8b-aef4-a0aaa8f5dc131757697281829.jpg"
    }
  },
  {
    "id": "2629",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥恒大帝景店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eaa2abd0-8cfe-42bd-93b6-a2914efb76d41757697276118.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/00f665cf-beeb-4118-9430-d47aed98a2cc1757697276794.jpg"
    }
  },
  {
    "id": "2350",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥恒通店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cc724ad1-d046-4a1c-8e18-4b32e5f4f1151757696920166.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/133b0916-a26e-45fb-8664-ee4919f2820f1757696920880.jpg"
    }
  },
  {
    "id": "2196",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥文昌雅居店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/12cc8c0e-0fc0-4cde-97d6-ea55f7c6bcaa1757696694084.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0b58a35b-fd25-48f6-b35e-4d774d24931a1757696694444.jpg"
    }
  },
  {
    "id": "1638",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥新华国际广场",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1638/1638-合肥新华国际广场-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1638/1638-合肥新华国际广场-经营许可证.jpg"
    }
  },
  {
    "id": "2304",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥新华纺店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1d4c4b18-badc-4aa7-9b00-94f8880bd4051757696853012.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8783435b-79b6-48a9-bbf9-917c7ed443511757696853675.jpg"
    }
  },
  {
    "id": "2768",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥新城国际店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c9bb8738-b61a-47fb-b1a0-a4b91cbe3c251757697421412.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f7a12f50-8356-4239-8cd5-ae590f6c0fc61757697422364.jpg"
    }
  },
  {
    "id": "1279",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥新桥服务区北餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1279/1279-合肥新桥服务区北餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1279/1279-合肥新桥服务区北餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1280",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥新桥服务区南餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1280/1280-合肥新桥服务区南餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1280/1280-合肥新桥服务区南餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "2795",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥新桥机场二店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4299497c-4822-4bdd-92dd-dad9ba9767a11757697482629.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f809ddb0-d83d-4c41-b545-2b900ac15b181757697483259.jpg"
    }
  },
  {
    "id": "2638",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥新站广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8b653656-fb8a-45d3-9f66-59e46314c9dc1757697286885.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0e10e56d-d917-4665-85e7-094ddaa7cae91757697287239.jpg"
    }
  },
  {
    "id": "2231",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥旭辉御府店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d51640b0-1f30-4cd5-8b7c-e7c9396fb3291757696734780.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9750b211-db89-4d64-b692-23eb70da7fb11757696735479.jpg"
    }
  },
  {
    "id": "2297",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥时代城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5c004629-52f1-4c24-9973-512e431b08d91757696841341.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b7c7783b-76e7-4b59-8fc2-d83d3fcfdf341757696842256.jpg"
    }
  },
  {
    "id": "2075",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥星火集贸市场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b12bf813-9eb8-4958-9681-3fea4ba544f81757696542975.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8e78db24-9d43-423c-be56-92c3737503461757696543350.jpg"
    }
  },
  {
    "id": "2711",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥星隆国际店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e1da68ca-ecc0-4b20-ad4e-d80cfec935221768842233959.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0bb1ba2d-c1d1-4aa9-9bd1-39df09fb26b01768842234647.jpg"
    }
  },
  {
    "id": "2145",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥望湖城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a7f3f304-c50b-4fd7-b202-ebbcb3d2b4041769706071512.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/15ce8f93-e12a-4a5c-8607-61497fb478621769706071924.jpg"
    }
  },
  {
    "id": "2344",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥柏堰雅苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b67c251a-72c4-4292-bf27-3979724839571757696908797.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d86da3d2-2dae-431f-8129-57d0409a420c1757696909481.jpg"
    }
  },
  {
    "id": "2249",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥柳林大道四季华庭店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c2910c0a-037a-4ba6-8bab-9996b82367341757696766735.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fbe69a06-e62c-45bc-a3da-620d00551f7d1757696767404.jpg"
    }
  },
  {
    "id": "2125",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥桐城南路江南新里程店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/68aa3260-a785-4942-87a2-c9a178c0a3371757696597238.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dacab667-efc1-48ef-bee2-020d0bdb35ae1757696597596.jpg"
    }
  },
  {
    "id": "2308",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥梅山公寓店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/20c69ccc-f4ed-4b93-8ccf-ded37da50cc21757696858811.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/962e7e6c-b7db-4682-ae75-9dfc3e168dc61757696859161.jpg"
    }
  },
  {
    "id": "2281",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥梦溪小镇店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ff28d4b9-b3ab-4459-be40-3cb17b0b8c691776704510287.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2ba4c55d-6b77-4556-a11d-d2de602ddb2e1776704510661.jpg"
    }
  },
  {
    "id": "2109",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥正创东景苑店",
    "licenses": {
      "business": "",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/store-center/images/2109 原2161正创东景苑店食品经营许可证（网络新）2019-10-81618219309031.jpg"
    }
  },
  {
    "id": "2721",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥海恒金屿海岸店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/01fdd7bb-2d7f-405b-8f16-60626e44aba71774976644151.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9c8861f4-48ac-4fa0-9865-d4eefa5b44571774976644962.jpg"
    }
  },
  {
    "id": "2472",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥海棠别院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/127a02c2-5d5b-4a7c-9894-6bebe31775721757697097811.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/50b68c39-6afe-4111-97ee-d8a7c41376da1757697098156.jpg"
    }
  },
  {
    "id": "2520",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥海棠路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/65d0df6b-22a5-4b30-8f35-2e5cd246b9221757697153569.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/67fbae76-3c9d-45d6-bfad-67b169ea3e9e1757697154344.jpg"
    }
  },
  {
    "id": "2453",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥淝南家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8515da74-997c-4ac3-b5fd-deff23bed3c21757697072157.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cbbbbddd-870e-439e-ae9a-402419383a431757697072934.jpg"
    }
  },
  {
    "id": "2326",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥淠河路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9732911b-b7cf-4b9c-b6a9-4ccb1ff355a61757696892062.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0e9fe2b0-4ba1-453e-a750-d65d43787aea1757696892439.jpg"
    }
  },
  {
    "id": "2003",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥淮河路分店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/15d1d480-92b5-466a-85d5-e63aabd583391757696426294.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3f757ea0-26eb-489f-b946-5d87dfb16f301757696426653.jpg"
    }
  },
  {
    "id": "2593",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥温莎广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/465ae77d-019e-4afe-9e0c-1343377d11461757697247788.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bbd85e24-be7e-4c3f-b29c-7e4be6b8b9881757697248164.jpg"
    }
  },
  {
    "id": "2550",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥港汇广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/05e736c7-a0f1-44a1-bb0b-ae6060732fa51757697199382.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/66c583a9-7d4f-450f-895b-d3d663b1fbdc1757697200071.jpg"
    }
  },
  {
    "id": "2537",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥港澳广场二店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3b3ec001-6807-44bc-9a16-7e48c99189541757697180852.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e228d32b-3942-40d9-8be0-6b69cfae01aa1757697181220.jpg"
    }
  },
  {
    "id": "2098",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥港澳广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1c7b077b-1fcd-49fc-966d-be9508560ac31757696572002.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/42a89770-98af-44a1-894e-dab5b712a8ae1757696572681.jpg"
    }
  },
  {
    "id": "2791",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨水花都店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fb2148ee-d82d-472b-b73f-ab9a5fddb0c11757697477876.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bc02672e-9aea-4466-91c8-efa8252adcf51757697478511.jpg"
    }
  },
  {
    "id": "2258",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨河小区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/59cd203e-c486-4815-b091-06abaff64e5a1757696780832.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/419be9bf-ec74-4adc-af8c-d53a73e97fa81757696781185.jpg"
    }
  },
  {
    "id": "2283",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖万科蓝山店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5eb4fd18-d630-4363-a36d-cdc79316c5e31777568516688.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/87794e6e-8694-41e9-a7e7-c53e5d5c11061777568517082.jpg"
    }
  },
  {
    "id": "2284",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖万达城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1dad0572-e9f6-43e6-aabf-df8f89c5340c1757696818417.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d0ec1bdd-6dd0-4032-8ca8-603db3c447901757696818802.jpg"
    }
  },
  {
    "id": "2320",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖万达茂店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/732d1aaa-2506-4952-bb97-5164aefc85e51757696880033.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/223ff8ac-767a-411a-8b88-ef3c5bb9e6051757696880834.jpg"
    }
  },
  {
    "id": "2467",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖佳源广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7d190718-7597-4560-9464-ee2dcf11de4c1757697091914.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4eba8d30-dd19-4edf-b6a9-0137fecee60b1757697092290.jpg"
    }
  },
  {
    "id": "2166",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖假日花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8407e7c9-f060-4e24-aa8f-63fd58d4fed51769187681126.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/136616af-2109-4d50-a58c-f58636d19b0a1769187681775.jpg"
    }
  },
  {
    "id": "2144",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/da1a9973-d7e0-4d5b-93bf-cf08603930e61780160478128.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9e52b258-6e88-41e8-b0f8-933951cf7f1c1780160479962.jpg"
    }
  },
  {
    "id": "2207",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖康园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3ca14b88-ac44-4910-a122-752241a91d991757696710299.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/07a762c7-9c42-45d6-a027-c9a7de7b11501757696710659.jpg"
    }
  },
  {
    "id": "2058",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖新区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f4d16158-9b20-4404-932b-e389b12a02571757696524519.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ede451a1-60d1-4a90-8723-5a9f3355ba3f1757696525202.jpg"
    }
  },
  {
    "id": "2042",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖新区徽杰苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0eaf81d9-9740-4d0a-8f7f-04836fa574d61757696495497.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/991fcb33-2a97-4122-9175-fb0ab200332e1757696495982.jpg"
    }
  },
  {
    "id": "2388",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖杭州路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d305f670-814b-4f8a-b17a-502cc192a7a31757696983288.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/05780308-4377-4af3-b132-f9b802252d621757696983671.jpg"
    }
  },
  {
    "id": "2666",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖菊园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7ed7a361-bf0b-47fa-8b47-ac9076ae20e81760547813545.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2744ccbc-bb6c-41d0-88f7-acbdc1429c051760547813935.jpg"
    }
  },
  {
    "id": "2394",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖金融港店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/514065ff-04fb-456b-9fb8-434600e8be791757696991909.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c4377afb-a23b-4cd8-af36-717dbfae9c321757696992272.jpg"
    }
  },
  {
    "id": "2451",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥滨湖顺园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c636f377-7328-4d32-a48e-75787cfc9d791757697069639.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0a0aa649-e829-4734-93b2-4439ec73dc251757697070244.jpg"
    }
  },
  {
    "id": "2078",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥火车站店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/10bd2504-6524-413e-b91c-d5c06ac42fe81757696548414.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c9094dc0-a618-484b-8423-bb119ae8d5ac1757696549083.jpg"
    }
  },
  {
    "id": "2654",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥玫瑰绅城花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e0b74ec0-86a2-4ac0-b435-1377e25128261757697308809.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/06316d94-17d2-4da3-af2a-036a5c10c80a1757697309175.jpg"
    }
  },
  {
    "id": "2636",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥琥珀五环城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0845bb38-7e6a-4819-a911-4e5c49a884611757697283460.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6e3c5434-5fe8-4a77-9bcd-8325d8fea2e41757697284366.jpg"
    }
  },
  {
    "id": "2182",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥琥珀名城和园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/592fea98-a6c7-43dd-bc3e-12ad8e502e421757696678186.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3d3b1680-1e1c-4e0d-8898-071e0a0d643d1757696678855.jpg"
    }
  },
  {
    "id": "2282",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥瑞地公馆店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ed6a6c62-41af-4f28-9a25-e3c73ddf2dfb1774976515675.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fe1a17b6-096a-4ae4-84e2-683c9738442e1774976516028.jpg"
    }
  },
  {
    "id": "2316",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥瑞徽苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/70d2ec7a-adbf-44e8-896c-da494a810c191757696874446.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8dc8c3e5-965b-46be-bbed-89f0bd6893e71757696875123.jpg"
    }
  },
  {
    "id": "2429",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥瑶海万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2d7eceed-5d1e-4c9d-9897-66ff9c13cbc31757697043331.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9b948db5-60de-4084-9334-ddf7aac30d871757697043716.jpg"
    }
  },
  {
    "id": "2588",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥瑶海西路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c257c7ce-6729-4ba4-8421-0ace3fbf9fb71757697241798.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1d0b1548-8589-4072-b9c3-805923c0e17b1757697242374.jpg"
    }
  },
  {
    "id": "2246",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥百乐门广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f7653bf0-11f8-4fe2-983c-5ee052395a281757696760616.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/84797b78-fa1e-4c14-98b1-1a0a304807da1757696761607.jpg"
    }
  },
  {
    "id": "2468",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥百大奥莱生活广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/03b87080-8abd-4fc3-954e-5d10afcc48e71757697094142.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6d47844f-25cc-45df-b21c-ed87628187e51757697094500.jpg"
    }
  },
  {
    "id": "2760",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥盛世家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0f13f9e2-0fbd-498f-ab69-a0946a9cbccc1757697413322.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/08d2e12d-8b94-4cc0-b5ee-b9974b3fce1b1757697413705.jpg"
    }
  },
  {
    "id": "2793",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥省二院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2aefec13-6cf9-4143-9ee0-080e642e0a1b1757697479799.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/67bf4659-49c8-4537-bf22-10ef3da37ca61757697480167.jpg"
    }
  },
  {
    "id": "2639",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥睦邻中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2f557574-9967-4262-9a11-c54920fccd161757697289141.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b8981556-8049-4527-8ab3-094d18350e2e1757697289819.jpg"
    }
  },
  {
    "id": "2094",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥社岗路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bd0046ab-30ed-408f-90f6-5e1aa40836e31757696564193.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c7687cf6-800a-4835-8502-58038247de551757696564550.jpg"
    }
  },
  {
    "id": "2418",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥禹州中央大街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/97bede79-8778-442c-a459-a4af5932e9021757697028756.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/154667b7-9f8c-44fc-87fa-465a8763663d1757697029458.jpg"
    }
  },
  {
    "id": "2481",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥禾润阳光店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/63977ec6-7830-4097-81e9-3f829d473b0e1757697109148.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fa1f118d-aaa3-47cc-b6ad-9d249be07a911757697109519.jpg"
    }
  },
  {
    "id": "2492",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥科大讯飞店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/385c7337-f160-487e-9a1a-020973a5b0481757697117854.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/212a0dc1-cb95-4fed-9a1d-613fc40da3ef1757697118543.jpg"
    }
  },
  {
    "id": "2749",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥立基大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/97d6e900-f3ab-4aa4-9a8d-0ca38958d7e81757697398979.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fd3d9298-06af-4702-9ab0-c6a8c66155e61757697399349.jpg"
    }
  },
  {
    "id": "2029",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥站前路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a351289b-db33-46f3-b06d-623f28575b881757696470985.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f220916e-fd32-4ceb-9fa8-e1cb051b3a0b1757696471621.jpg"
    }
  },
  {
    "id": "2590",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥站塘路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c56f7077-3041-4be6-9ae1-0f1e4129cfca1763571802837.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a2e5f21c-eb63-4a17-acde-f6927d7108821763571803221.jpg"
    }
  },
  {
    "id": "1572",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥第二人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1572/1572-合肥第二人民医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1572/1572-合肥第二人民医院店-经营许可证.jpg"
    }
  },
  {
    "id": "2076",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥紫桐新村小区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c4b02c3f-305a-4f22-91ad-edfcc75569d71757696545475.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f52f858c-dc8f-4488-84d1-ced91cfe19d81757696546193.jpg"
    }
  },
  {
    "id": "2474",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥紫蓬路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/67179c90-845e-4039-8238-4180aaf296921757697100671.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/34ee86bd-61e2-4aee-81f8-049a0de9035e1757697101025.jpg"
    }
  },
  {
    "id": "1346",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥紫蓬镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1346/1346-合肥紫蓬镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1346/1346-合肥紫蓬镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2787",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥置地星光荟店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a1d73a17-afbc-4a24-a8ab-36be45c96f0e1757697466190.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/437d0a6d-4ccd-42d5-b558-e1c556d3620e1757697466554.jpg"
    }
  },
  {
    "id": "2529",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥联投中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3f66e014-d0cb-4484-a1cd-4a6c521d75311757697164703.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3de48cbf-5c0b-4a2b-8ece-15af0f46bbcb1757697165081.jpg"
    }
  },
  {
    "id": "2797",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥聚福家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f82dd1bd-7ec1-45f1-80d9-9af76db7b9a91757697489816.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7864b979-514e-446b-9af8-127fc47cc6491757697490528.jpg"
    }
  },
  {
    "id": "2318",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥肥东军天湖路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c4979c6c-7e73-475e-81a8-7ef85e7ef7621757696877827.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9eaca77a-4359-4e9d-8927-4027b404743e1757696878181.jpg"
    }
  },
  {
    "id": "2511",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥肥东吾悦广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9457b67a-1a71-43ab-9886-1b5c297592481757697141695.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9fff61d9-9b34-4729-9f45-27e42f5e740f1757697142056.jpg"
    }
  },
  {
    "id": "2376",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥肥东幸福家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/133a0e2d-0bff-4089-bf5f-76211099364d1757696961880.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c23ef252-7c88-4e67-a0c4-13f87de8ccbb1757696962229.jpg"
    }
  },
  {
    "id": "2378",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥肥东文一名都店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6daeebe8-70a5-471c-b330-dbed407d81141757696964143.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e5544a0c-ece1-4407-8f95-78cecda543eb1757696964807.jpg"
    }
  },
  {
    "id": "2280",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥肥东斌峰中心城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ce38323e-9e08-4648-b7c1-8bd79c5bc5071757696807356.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a0ec6800-c762-4448-a487-67cf76d8e0ef1757696808235.jpg"
    }
  },
  {
    "id": "2790",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥肥东星光嘉园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9d220b9d-fe21-4f71-b14d-e214a92703ed1757697475590.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/52b9ebf0-3f26-421c-942b-69113175ba0c1757697475944.jpg"
    }
  },
  {
    "id": "2001",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥舒城路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/21bd747c-943f-4978-87fa-93576686fa4d1769187624583.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dedbb7af-5696-400f-87d7-29d1844466581769187624935.jpg"
    }
  },
  {
    "id": "2064",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥茨河路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8809a29c-6844-41d1-b5d4-c75cf61be3b51757696535028.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/34822ac2-e410-4926-854b-1bd5a211fd691757696535402.jpg"
    }
  },
  {
    "id": "1252",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥董岗荟萃园餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1252/1252-合肥董岗荟萃园餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1252/1252-合肥董岗荟萃园餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "2457",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥蒙河路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a6dad937-6d07-40c2-a2d8-95fd4b7e4e051757697080108.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b87fa2dc-f9d1-4e67-886f-141b5271bfd51757697080808.jpg"
    }
  },
  {
    "id": "1658",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥蔚来园区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1658/1658-合肥蔚来园区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1658/1658-合肥蔚来园区店-经营许可证.jpg"
    }
  },
  {
    "id": "2137",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥蜀峰路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ba903101-f728-44cc-9360-ff8e98d6d03e1775149265610.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7ef02648-d681-4904-ba1a-42ca3939d4a01775149265998.jpg"
    }
  },
  {
    "id": "2382",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥融侨悦城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eb6b2e3c-3566-49b3-956e-3c1f44f9f1791757696972711.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c86340dd-669d-4dbf-8978-f3da3a9cfd581757696973068.jpg"
    }
  },
  {
    "id": "2742",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥裕兴家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/16b4f800-7076-46f0-99b9-dd34b1c6a8a41757697389350.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8bcbce9b-cf72-408d-8fec-e6caa570645d1757697389730.jpg"
    }
  },
  {
    "id": "2325",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥裕溪路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2251fb5d-dbe2-4015-b9b9-72c5b8f256f51757696889024.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8605cacd-b2c0-4169-937d-45fed02a68fe1757696889721.jpg"
    }
  },
  {
    "id": "2476",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥观澜华庭店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e2a0e353-187a-495e-bd92-d4a9cb78e5a21757697103807.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cf32fe71-5afb-45cf-a79d-171d541cd3141757697104167.jpg"
    }
  },
  {
    "id": "2544",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥财富广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bc573ff7-bb75-4de7-b688-3541077375b11757697190980.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cdf08a30-411d-4c74-a70c-f546e8b350ba1757697191622.jpg"
    }
  },
  {
    "id": "2517",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥贵池路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/513c1865-6d11-4726-8f3d-bf1ab0aa19f81757697148318.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9030b247-2f6c-4282-bff4-32fe69fbd1cb1757697148693.jpg"
    }
  },
  {
    "id": "2329",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥金中环广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4749589c-76b1-4894-90ce-a8ff63363a831757696895001.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/58369ef0-e2cf-461d-a38c-3ddcee8c65201757696895351.jpg"
    }
  },
  {
    "id": "2061",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥金寨路三店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1df06117-3405-4e06-9eed-7a16b0aff1081757696532027.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ce3f690b-3467-40b3-ba61-b7153c702e021757696532691.jpg"
    }
  },
  {
    "id": "2148",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥金珠路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0ed52db3-e149-44fc-9eed-381afd8857a21769187674051.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8d262464-d693-4eea-ae73-68a1c3b18f291769187674753.jpg"
    }
  },
  {
    "id": "2489",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥金葡萄家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3a48f4f5-d10f-4f78-b985-482f5e32bd5f1756832864338.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1fa450ce-c837-4dbd-8338-27555a2382391756832864652.jpg"
    }
  },
  {
    "id": "2363",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥长丰北城华府店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9de6ddad-0878-4ce8-9103-6753b582f3e71757696935803.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1c1c961a-e18f-4c67-ada6-6b12c25dfacf1757696936501.jpg"
    }
  },
  {
    "id": "2368",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥长丰岗集店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ef491050-afbf-4d93-8360-710557260c1d1757696950267.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5e438867-f0e8-4053-8c09-20322e02eb471757696950659.jpg"
    }
  },
  {
    "id": "2546",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥阜阳路大润发店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aee39872-108e-4d8a-ab6e-9332c58e82ad1757697194202.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fd0decaf-4d29-4a6b-9be4-19e99d8963f61757697194572.jpg"
    }
  },
  {
    "id": "1742",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥陈埠服务区东区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1742/1742-合肥陈埠服务区东区店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1742/1742-合肥陈埠服务区东区店-经营许可证.jpeg"
    }
  },
  {
    "id": "1743",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥陈埠服务区西区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1743/1743-合肥陈埠服务区西区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1743/1743-合肥陈埠服务区西区店-经营许可证.jpg"
    }
  },
  {
    "id": "2450",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥陶冲湖广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ff329493-55df-4d97-934c-63f050708c461757697067050.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1dff673e-cf72-4a66-86d2-5346b1dc99491757697067412.jpg"
    }
  },
  {
    "id": "2272",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥陶然居店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fceeaf30-65a0-477c-97b8-cc1d7bf153e11757696798461.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品经营许可证1775268433364.jpg"
    }
  },
  {
    "id": "2040",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥青阳南路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/df8c9506-ebb8-435b-bc6c-d96aa2daf2701769014838329.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/63953a30-991f-499b-bff8-2fbf8179caea1769014838721.jpg"
    }
  },
  {
    "id": "2693",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥青龙潭路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cae84c01-12db-4637-bfe3-f925e39802f61770224638519.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/11446ea8-956a-40a4-a56f-088d48c1b4791770224638894.jpg"
    }
  },
  {
    "id": "2786",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥顺丰产业园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/237d260a-aa40-4c6e-8ae5-3f36076d67781757697462714.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/68728b1a-8abd-490b-bd2f-09edfc1f17e31757697463568.jpg"
    }
  },
  {
    "id": "2009",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥颍上路分店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a6c299be-524d-42f6-89a5-8405293eae941757696441431.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5a99d233-2e29-4ce3-80e1-0f087a14d0721757696442069.jpg"
    }
  },
  {
    "id": "2402",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥风格城市店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eba08bf7-73fa-4b1b-a5ad-4fa273b7c36d1757696998147.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5e791a86-7f15-4a5d-b37c-c13d7bfab16f1757696998532.jpg"
    }
  },
  {
    "id": "1991",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥高刘镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1991/1991-合肥高刘镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1991/1991-合肥高刘镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2743",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥高速时代公馆店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e45a5b4d-04d1-4f63-bc43-0e2cb963ce5c1757697391855.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a184f33b-fc50-4f17-a9ab-3459e88ba1ec1757697392216.jpg"
    }
  },
  {
    "id": "2836",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥高速时代广场店",
    "licenses": {
      "business": "",
      "food": ""
    }
  },
  {
    "id": "2233",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥高铁南站店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c9b8107f-cc47-44fc-8b5f-2e7836f4f5c31770224501351.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1351252f-e215-444c-afb2-5e416a3d4d2c1770224502024.jpg"
    }
  },
  {
    "id": "2158",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥鸿兴苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fa463eb8-46ee-412c-a1d2-f174ec790b871775840473808.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/89bea342-d6f0-4ae2-8395-6290db3193891775840474462.jpg"
    }
  },
  {
    "id": "2586",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥黄山路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2dfe784c-28bc-4e11-b707-31bbca51658a1757697238515.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1c88b161-d076-4757-8554-1970ad0efcce1757697238935.jpg"
    }
  },
  {
    "id": "2337",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥鼎元公馆店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/22475835-cd05-4369-b80e-bc1b6c2eab131757696897328.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/64bab91e-e857-4b7b-be2c-25ee719a2b951757696898195.jpg"
    }
  },
  {
    "id": "2673",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥龙城嘉苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7bed6ec8-c8b6-4c0b-89f5-8fd427ea13711757697328570.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/db688dd3-5691-4514-af37-dd5569330d511757697329256.jpg"
    }
  },
  {
    "id": "2485",
    "province": "安徽省",
    "city": "合肥市",
    "name": "合肥龙湖天街广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/43feb319-4ef9-4c01-9a78-d54c47810b341757697111749.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/495baa69-6d37-4554-8f9b-1c20935b47151757697112131.jpg"
    }
  },
  {
    "id": "2293",
    "province": "安徽省",
    "city": "合肥市",
    "name": "和信大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/801dc9ba-4384-47b8-a0fc-30534540deb91780160527953.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dc9258d5-22d6-457e-a0fd-ad28cbd0bba71780160529779.jpg"
    }
  },
  {
    "id": "2007",
    "province": "安徽省",
    "city": "合肥市",
    "name": "和平路分店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e39c91fc-7f31-4efd-b75f-a68ae59293f51757696436509.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/39f3d815-0d63-45bb-8b6e-2f7e86b644a51757696437157.jpg"
    }
  },
  {
    "id": "2015",
    "province": "安徽省",
    "city": "合肥市",
    "name": "国际花都店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4d5606fc-c296-4eea-93ad-8bf6b0b3fb121768582830418.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f64ef2e8-823b-4782-bf3f-f55cd144a9841768582831278.jpg"
    }
  },
  {
    "id": "2254",
    "province": "安徽省",
    "city": "合肥市",
    "name": "大溪地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c2ce5d73-8b4a-4e7f-af5d-a85f87f4cdd91757696777470.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3aad7d96-07a9-493f-ae4f-3b9e1eb4e10f1757696778133.jpg"
    }
  },
  {
    "id": "2038",
    "province": "安徽省",
    "city": "合肥市",
    "name": "天柱路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9f1feacf-91a5-458e-8901-c3dc039fb2ba1757696484289.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8d5ee2b3-d11c-425f-8430-b5c6193d1b5a1757696484676.jpg"
    }
  },
  {
    "id": "2107",
    "province": "安徽省",
    "city": "合肥市",
    "name": "太宁花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/530b4945-b76b-429a-8182-90090ff9e1291757696577846.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e85077d0-4268-4afb-8f2f-0c8203b9da841757696578498.jpg"
    }
  },
  {
    "id": "2017",
    "province": "安徽省",
    "city": "合肥市",
    "name": "太湖路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5ce4b38f-b987-4fb1-a392-c05591a264c81770224434084.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5f6f264e-c5c6-4b22-8d68-91e6d63b57491770224434453.jpg"
    }
  },
  {
    "id": "2021",
    "province": "安徽省",
    "city": "合肥市",
    "name": "宁国北路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/46366500-816c-4251-90a9-27150cd936581757696460781.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b2b6acbb-35a5-47d7-b0dd-65f0a02a7adb1757696461135.jpg"
    }
  },
  {
    "id": "2008",
    "province": "安徽省",
    "city": "合肥市",
    "name": "宁国路分店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c310ae01-96c5-4d1e-8e10-8956af48a7e71757696438987.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9be17772-157d-47c8-8e3a-275bc5c293531757696439639.jpg"
    }
  },
  {
    "id": "2172",
    "province": "安徽省",
    "city": "合肥市",
    "name": "安徽大剧院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aeff7d67-7fc7-4966-9a8e-b6e7c2b80e751757696664693.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dc7dabee-e48f-45c6-8f71-28126a2b5ff71757696665078.jpg"
    }
  },
  {
    "id": "2096",
    "province": "安徽省",
    "city": "合肥市",
    "name": "寿春路一店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4302890c-3b89-486b-8301-5a7180baa0b31757696569003.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/634a46df-073d-4851-aec0-daf04b0c6c2b1757696569657.jpg"
    }
  },
  {
    "id": "2736",
    "province": "安徽省",
    "city": "合肥市",
    "name": "巢湖东方新世界店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e455b5b2-6ecf-4b77-b2f6-3ce4ed3073d11757697380931.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/094a196b-cab0-4402-9670-2147e7a4f3261757697381309.jpg"
    }
  },
  {
    "id": "2541",
    "province": "安徽省",
    "city": "合肥市",
    "name": "巢湖凤凰之家店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/47635116-e7da-4654-89e6-67d2cfc832e71757697188437.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ac1697cf-6c27-4171-9189-8dc4df73e58e1757697189124.jpg"
    }
  },
  {
    "id": "2149",
    "province": "安徽省",
    "city": "合肥市",
    "name": "巢湖向阳路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0aca950e-04f9-44a6-8642-5b3ae4e608131768842070646.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/470b3759-186b-435d-978c-e88101681bd11768842071397.jpg"
    }
  },
  {
    "id": "1639",
    "province": "安徽省",
    "city": "合肥市",
    "name": "巢湖商之都",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1639/1639-巢湖商之都-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1639/1639-巢湖商之都-经营许可证.jpg"
    }
  },
  {
    "id": "1811",
    "province": "安徽省",
    "city": "合肥市",
    "name": "巢湖黄麓半岛商业广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1811/1811-巢湖黄麓半岛商业广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1811/1811-巢湖黄麓半岛商业广场店-经营许可证.jpg"
    }
  },
  {
    "id": "2381",
    "province": "安徽省",
    "city": "合肥市",
    "name": "庐江世纪华府店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/87d1b67d-b2d8-4323-bb6f-473e5fbb5dd11757696970160.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1cd216a1-f11c-417a-b7e2-cf6919cbbeee1757696970549.jpg"
    }
  },
  {
    "id": "2300",
    "province": "安徽省",
    "city": "合肥市",
    "name": "庐江中心城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/688ac707-5d43-4413-b019-04f634c16b711757696845287.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/962a539d-93df-4023-8865-90361ffffa101757696845656.jpg"
    }
  },
  {
    "id": "2355",
    "province": "安徽省",
    "city": "合肥市",
    "name": "庐江安德利广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/65efc41d-414a-4b3a-9399-5d34660704871757696928147.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d20fd02f-2dd9-46fa-b903-ce8d88e4f0791757696928519.jpg"
    }
  },
  {
    "id": "1801",
    "province": "安徽省",
    "city": "合肥市",
    "name": "庐江汤池鑫隆店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1801/1801-庐江汤池鑫隆店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1801/1801-庐江汤池鑫隆店-经营许可证.jpeg"
    }
  },
  {
    "id": "2223",
    "province": "安徽省",
    "city": "合肥市",
    "name": "御景湾店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e6f2b8a8-3d44-497b-9d53-991fb2187bc11757696726823.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d152d94f-1584-4b37-95a7-d381a80804131757696727184.jpg"
    }
  },
  {
    "id": "2187",
    "province": "安徽省",
    "city": "合肥市",
    "name": "怀宁路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b9ddf267-0537-4e8d-bf1a-7958a714f6181757696690557.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7863d54c-d5f6-4148-9cf2-2884188253891757696691297.jpg"
    }
  },
  {
    "id": "2271",
    "province": "安徽省",
    "city": "合肥市",
    "name": "恒盛皇家花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2dec81dc-238d-4867-baf7-7af6b19b4b581769533302829.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/666d986b-7ebd-4826-804f-f613f82cf40f1769533303537.jpg"
    }
  },
  {
    "id": "2159",
    "province": "安徽省",
    "city": "合肥市",
    "name": "文忠路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/97ed0801-2002-49c9-acc2-36dc9727bfa61775840476286.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/28591740-7775-4170-a7a2-ba86255c62a61775840476930.jpg"
    }
  },
  {
    "id": "2161",
    "province": "安徽省",
    "city": "合肥市",
    "name": "新桥机场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/86f4930e-de8d-4b0e-bc83-8af8eee4c1f71757696648172.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1bf249c2-3dea-4af6-bcd6-39de66310ed21757696648904.jpg"
    }
  },
  {
    "id": "2228",
    "province": "安徽省",
    "city": "合肥市",
    "name": "新海尚宸家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a19d08e5-3d52-4b18-bfba-c12314ba4cb71757696729110.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9fb64ed0-7773-43e4-8b47-6c732112a77e1757696729781.jpg"
    }
  },
  {
    "id": "2020",
    "province": "安徽省",
    "city": "合肥市",
    "name": "望江西路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e7899ab5-15de-4efe-bac5-e3897754d7981770915632411.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a8fdd38f-b26a-4a43-98f8-dd162609c5951770915632754.jpg"
    }
  },
  {
    "id": "2114",
    "province": "安徽省",
    "city": "合肥市",
    "name": "桐城南路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6a073175-c599-40ed-b24e-58d8d7a9db141768842056743.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6279e85b-1dff-42d8-857b-0acaba7f744f1768842057337.jpg"
    }
  },
  {
    "id": "2301",
    "province": "安徽省",
    "city": "合肥市",
    "name": "永和家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5ba25e86-7d15-4df9-bf5a-5532c85180e41757696847747.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a9595ed9-d1a8-456a-9232-98e7272492061757696848103.jpg"
    }
  },
  {
    "id": "2245",
    "province": "安徽省",
    "city": "合肥市",
    "name": "沃野花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f3733652-89ce-466e-8868-1db2ca2026881757696757969.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/05f13365-daff-4982-ad09-4b36f85604781757696758636.jpg"
    }
  },
  {
    "id": "2277",
    "province": "安徽省",
    "city": "合肥市",
    "name": "泰盛广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2bd512de-bcdb-495f-8fc0-6bb49f6c51bf1757696805019.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/18f2e4ac-56b7-4736-b8a7-11bed5f976ec1757696805413.jpg"
    }
  },
  {
    "id": "2059",
    "province": "安徽省",
    "city": "合肥市",
    "name": "海恒社区店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c5245d5c-ec44-46fa-bb57-608796d5d97c1757696527307.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bb500c59-5cff-4164-93d5-f91b71eded431757696527659.jpg"
    }
  },
  {
    "id": "2198",
    "province": "安徽省",
    "city": "合肥市",
    "name": "滁州路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ee650407-2158-43c0-bdb2-fc6b7b971e301757696697195.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/61617887-41a6-4e7b-9ee5-cc1bcbc86a9d1757696697578.jpg"
    }
  },
  {
    "id": "2121",
    "province": "安徽省",
    "city": "合肥市",
    "name": "潜山路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/890fb1d4-1ee7-43b4-bc8b-f3683e6b1dea1761325257843.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/01760d2a-19fa-41bd-8413-477f3b0c17fc1761325258212.jpg"
    }
  },
  {
    "id": "2050",
    "province": "安徽省",
    "city": "合肥市",
    "name": "濉溪东路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/55626ae0-cdcb-4498-8901-ce033a91bab11757696508326.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/69486677-675b-4310-a016-cc437bac4dd41757696508686.jpg"
    }
  },
  {
    "id": "2052",
    "province": "安徽省",
    "city": "合肥市",
    "name": "琅琊山路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/69575570-64ae-4a4f-8268-0e4e0094a6821757696513686.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/91b47778-7adf-4a3e-b86d-f75d58207fb91757696514047.jpg"
    }
  },
  {
    "id": "2039",
    "province": "安徽省",
    "city": "合肥市",
    "name": "瑶海区通达路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0c47dd3a-90cc-48f8-ab9c-3a0d14a599351774976439815.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/477a3881-4b45-4f92-aea8-4b31fe540a741774976440517.jpg"
    }
  },
  {
    "id": "2117",
    "province": "安徽省",
    "city": "合肥市",
    "name": "皖河支路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a450b8fd-ecb8-4260-8f7f-73a44059acd41757696584705.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/220b7ae3-0d4d-4276-9af3-ad773101ca581757696585543.jpg"
    }
  },
  {
    "id": "2367",
    "province": "安徽省",
    "city": "合肥市",
    "name": "祥源城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b6f868e5-6ff8-4540-a9e8-aecf72111c491757696947354.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/23435873-1134-4d71-84f2-3c7b1e8464101757696948027.jpg"
    }
  },
  {
    "id": "2783",
    "province": "安徽省",
    "city": "合肥市",
    "name": "第一人民医院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/19d86db6-a0c2-49cb-a901-e90caad0dfd91757697457510.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f9d2d202-f602-4ac1-8506-d4ffd0038edc1757697457876.jpg"
    }
  },
  {
    "id": "2141",
    "province": "安徽省",
    "city": "合肥市",
    "name": "绩溪路三店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/48e47973-91e8-441f-b003-48848a9776991770224473726.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7a3ec0f6-2e35-47dc-824d-80add6faa1641770224474094.jpg"
    }
  },
  {
    "id": "2049",
    "province": "安徽省",
    "city": "合肥市",
    "name": "翠微路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a4dde7c4-d088-40b0-b2eb-4eeae0bd7a871757696505544.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d79c0512-c623-4648-a945-ed12b25f64201757696506186.jpg"
    }
  },
  {
    "id": "2131",
    "province": "安徽省",
    "city": "合肥市",
    "name": "翡翠商城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6fc7d3c1-c94c-44a3-b7c0-06358c1f9f831757696607683.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/16434445-d65f-4dc9-926e-3802eec5df2f1757696608316.jpg"
    }
  },
  {
    "id": "2035",
    "province": "安徽省",
    "city": "合肥市",
    "name": "肥东县龙泉西路店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/21768833744974.png",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/31768833757670.png"
    }
  },
  {
    "id": "2267",
    "province": "安徽省",
    "city": "合肥市",
    "name": "肥东彩虹新城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d56f02f6-e2e2-4f6b-b6f7-31cb4f6b29921757696789789.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c6ee6612-191f-41bc-9879-93c82aa3ae611757696790469.jpg"
    }
  },
  {
    "id": "2419",
    "province": "安徽省",
    "city": "合肥市",
    "name": "肥东润和尚品店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3f6112a0-0d8c-49a2-aea3-c7113e98a5a71757697031760.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/29e5120f-9ad2-4e77-93b0-da267a4bb9a11757697032119.jpg"
    }
  },
  {
    "id": "2379",
    "province": "安徽省",
    "city": "合肥市",
    "name": "肥东禹洲广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d34640a2-b3fe-4b62-bf89-938d88fb0bd71757696967059.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9777ac55-09af-4c82-a05b-e30a2c825d511757696967446.jpg"
    }
  },
  {
    "id": "2290",
    "province": "安徽省",
    "city": "合肥市",
    "name": "肥东金巨大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a096f5e2-0a70-426d-9f1c-6a2cb594ac581777568524470.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/235062c1-7bbb-4b6a-b5ab-7606a7db5c6c1777568524848.jpg"
    }
  },
  {
    "id": "1677",
    "province": "安徽省",
    "city": "合肥市",
    "name": "肥东长临河镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1677/1677-肥东长临河镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1677/1677-肥东长临河镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1760",
    "province": "安徽省",
    "city": "合肥市",
    "name": "肥西花岗店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1760/1760-肥西花岗店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1760/1760-肥西花岗店-经营许可证.jpg"
    }
  },
  {
    "id": "2528",
    "province": "安徽省",
    "city": "合肥市",
    "name": "花园大道唯美创想城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/31475adf-a833-4f53-b83d-5da42d40cd981764435780225.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c1c7c527-6672-498b-b9ee-dcc4c25f0ff01764435780609.jpg"
    }
  },
  {
    "id": "2045",
    "province": "安徽省",
    "city": "合肥市",
    "name": "莲花广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3ce63018-95c7-46c5-ab50-cf3aa6783bbd1757696500649.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1902c0fc-7015-4b13-8095-2bb1f48e58951757696501036.jpg"
    }
  },
  {
    "id": "2079",
    "province": "安徽省",
    "city": "合肥市",
    "name": "莲花路三店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/471aa690-2e23-44de-abed-70e7ce7416da1757696551277.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8a63a2f6-6418-41f3-963d-f05650caf4741757696551642.jpg"
    }
  },
  {
    "id": "2051",
    "province": "安徽省",
    "city": "合肥市",
    "name": "莲花路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/70a656f6-5f46-4d31-9f9d-2c857f7651d71757696511116.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/449b6ae2-a545-4d7d-b847-54000dabe8171757696511498.jpg"
    }
  },
  {
    "id": "2241",
    "province": "安徽省",
    "city": "合肥市",
    "name": "蓝筹国际大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/74d6498a-074a-40f6-aa9b-ac59037f7ad91757696749792.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dfe71e77-6d20-4322-a20d-976173f86ca41757696750481.jpg"
    }
  },
  {
    "id": "2480",
    "province": "安徽省",
    "city": "合肥市",
    "name": "蓬莱花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3a773749-5510-4bbe-bc45-f46a665f6ade1757697106407.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8237ab76-eb51-4dd1-927f-053f6d9fabce1757697106784.jpg"
    }
  },
  {
    "id": "2548",
    "province": "安徽省",
    "city": "合肥市",
    "name": "蔚蓝商务港店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6b4f79a0-9a1b-4ba4-8f40-8d19c8e057261757697196813.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/655053d7-66ce-4e5e-a7d6-147ecc514a591757697197496.jpg"
    }
  },
  {
    "id": "2313",
    "province": "安徽省",
    "city": "合肥市",
    "name": "蜀山区新产业园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3da09020-c385-4ae1-bc82-0d8fe0e28e751757696869030.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/82b8d01a-396d-477c-912c-bd51960369051757696869408.jpg"
    }
  },
  {
    "id": "2364",
    "province": "安徽省",
    "city": "合肥市",
    "name": "融科梧桐里店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f11549da-241c-4880-b632-fd90e1a218a61757696938777.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a874dd7c-cd1b-43af-a21b-dba92def76ce1757696939111.jpg"
    }
  },
  {
    "id": "2127",
    "province": "安徽省",
    "city": "合肥市",
    "name": "西蜀名苑店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/947962c1-c98d-447a-8a11-3f0c34b8c3791768582862361.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/96956f52-a0f3-49c4-bb5c-37cd86f6c6651768582863296.jpg"
    }
  },
  {
    "id": "2089",
    "province": "安徽省",
    "city": "合肥市",
    "name": "铜陵路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/619b658c-d20b-40e6-a34f-3d724038e7711757696559001.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7292b6b5-8330-4ec6-afbe-c0fbf614869b1757696559376.jpg"
    }
  },
  {
    "id": "2181",
    "province": "安徽省",
    "city": "合肥市",
    "name": "银领时代花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1b6f10b8-6910-4cef-bc46-b123fc2b5fc21774976482939.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2a3b1822-aeed-4015-8853-9513aeceb26f1774976483301.jpg"
    }
  },
  {
    "id": "1940",
    "province": "安徽省",
    "city": "合肥市",
    "name": "长丰玛特大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1940/1940-长丰玛特大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1940/1940-长丰玛特大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1939",
    "province": "安徽省",
    "city": "合肥市",
    "name": "长丰长寿路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1939/1939-长丰长寿路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1939/1939-长丰长寿路店-经营许可证.jpg"
    }
  },
  {
    "id": "2002",
    "province": "安徽省",
    "city": "合肥市",
    "name": "长江东路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f1655225-0b52-44a8-9531-b5817b4c55381770224424741.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9cfb839a-48cb-40eb-9566-3eeb5c6d1b591770224425444.jpg"
    }
  },
  {
    "id": "2028",
    "province": "安徽省",
    "city": "合肥市",
    "name": "长江西路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/16bf3143-d4cb-4eb4-9b4f-364e7055da191757696468764.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5991c2e5-c76d-4c26-b582-9cb08b155edc1757696469132.jpg"
    }
  },
  {
    "id": "2239",
    "province": "安徽省",
    "city": "合肥市",
    "name": "阳光帝景店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dca7f09d-ae78-4113-af5c-22498c2b95d51757696744864.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3a96036c-c442-4025-8483-93792ac7e06b1757696745217.jpg"
    }
  },
  {
    "id": "2030",
    "province": "安徽省",
    "city": "合肥市",
    "name": "青阳北路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9e167b6b-d639-4333-ba82-31a75e2bb8a01768842035360.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c437ef5f-fbee-4ccb-8ac4-b23d79f4a6941768842035738.jpg"
    }
  },
  {
    "id": "2219",
    "province": "安徽省",
    "city": "合肥市",
    "name": "领翔花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f4bd48c0-f389-4286-9929-2afc4ac85fa01757696719008.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/776d75d3-5da4-4552-bc0e-85324a0440231757696719364.jpg"
    }
  },
  {
    "id": "2072",
    "province": "安徽省",
    "city": "合肥市",
    "name": "马鞍山南路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7d272694-fc24-4f6a-bf5a-5383d66c4fab1757696537688.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dfec8793-999a-409b-8c4d-f776ec9718051757696538075.jpg"
    }
  },
  {
    "id": "2237",
    "province": "安徽省",
    "city": "合肥市",
    "name": "高速翡翠湖畔店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9cb8a339-d3f8-4a2f-bfce-e1107efe61941757696741704.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/071ccf8c-6e1a-4227-afc0-6ccecfc48ef71757696742059.jpg"
    }
  },
  {
    "id": "2041",
    "province": "安徽省",
    "city": "合肥市",
    "name": "黄山东路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/afb5dd66-148a-4f45-bdb0-e1685284df671780160443616.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b3cfa610-238f-4cf1-91a2-1f5d63410c691780160444530.jpg"
    }
  },
  {
    "id": "2056",
    "province": "安徽省",
    "city": "合肥市",
    "name": "黄山花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5d8a5a8e-1f33-4e88-8cf3-a18dc1b9f5721757696522204.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/58443316-6eda-484a-8f92-c63eac0419e71757696522634.jpg"
    }
  },
  {
    "id": "2767",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆七街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f84626bb-e7de-4b36-908e-4434b49b068f1757697418970.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/32fb868b-097d-4796-ad73-84939dbc83c91757697419368.jpg"
    }
  },
  {
    "id": "2769",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆人民路步行街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/41919e2e-88f8-4324-84f6-ee0c5c432f071757697424349.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8319e408-88e8-47e0-9064-b877827816651757697425205.jpg"
    }
  },
  {
    "id": "2812",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆吾悦广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8032bc81-d634-443b-a795-7d1b8345fdc81757697507640.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0281d94a-4b9e-41f3-8e0f-3fa3804c526e1757697508234.jpg"
    }
  },
  {
    "id": "2099",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆天柱山路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/221273c8-74c1-4606-b649-3ade085c88d31757696575127.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/434d9f06-ed2e-4748-9cc2-d6973ebfdbb61757696575539.jpg"
    }
  },
  {
    "id": "2610",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆太湖县一店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/营业执照1766625351806.png",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品许可证1766625359546.png"
    }
  },
  {
    "id": "2890",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆太湖程岭美食城店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/6ceff4f4cd1d4f8fbc17fb43ce0f30ce1776653510390.png",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/15fc17c85e1241fb833f6d1a9f3fd0cd1776653532810.png"
    }
  },
  {
    "id": "2684",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆宜城路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c907940f-8868-48ee-a953-f6325c7c9ec61765299982466.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bccdf823-42fa-48b3-b0ab-c44cdbba57dc1765299987159.jpg"
    }
  },
  {
    "id": "2661",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆宿松大润发店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f43d4081-6b0b-4819-8a0f-c3301b47b1ef1757697313948.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/284bde10-6ce9-417d-a7df-359c372817151757697314336.jpg"
    }
  },
  {
    "id": "2410",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆宿松路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/466e7bbc-81c9-4499-a42d-6871f85b7da91757697014902.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/515c83cd-4076-4e0e-8a4c-fe80e0a40e801757697015261.jpg"
    }
  },
  {
    "id": "2732",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆岳西好又多店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9297e5e4-67b0-4feb-8fe8-63c24938ec691757697377236.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1eac4740-0fd7-4174-9f8a-a279c81909a31757697377576.jpg"
    }
  },
  {
    "id": "2785",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆岳西莲云融鑫世纪城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/731bc426-139e-4817-a114-61c7798222b51757697460084.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b2813337-e866-451d-9fd8-363c447d7b6a1757697460745.jpg"
    }
  },
  {
    "id": "2184",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆怀宁和谐家园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c2e0259d-2323-4c1b-b678-03cd7d059e961757696683152.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6e6cadee-81b8-4723-8562-eaa74551bf161757696683790.jpg"
    }
  },
  {
    "id": "2147",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆怀宁步行街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cf807271-54e5-4f3b-8abe-85f6c57729c91773680470800.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d30b9284-4b2c-4892-9fd8-399b4d9589281773680471132.jpg"
    }
  },
  {
    "id": "1979",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆怀宁石牌镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1979/1979-安庆怀宁石牌镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1979/1979-安庆怀宁石牌镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2505",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆怀宁香樟大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/958ec183-db1c-4519-b5ae-7e7b828b6c581757697135394.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d28416cf-fc88-45d2-8a78-9b9dba9bd6cb1757697136070.jpg"
    }
  },
  {
    "id": "2503",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆文苑路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8bd184ac-33f5-4ca2-a326-8a0f39690c1f1757697132959.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1c98f2c1-1a92-4677-9ee9-76452757ff301757697133336.jpg"
    }
  },
  {
    "id": "2789",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆方圆荟店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f0495c1d-2992-4fcc-b1fc-d34c5766feb51757697472231.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/92c016d1-6b51-49d8-a83a-d3ac247a6c261757697472912.jpg"
    }
  },
  {
    "id": "2580",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆望江一店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/营业1766843984789.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品1766843993446.jpg"
    }
  },
  {
    "id": "1651",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆桐城人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1651/1651-安庆桐城人民医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1651/1651-安庆桐城人民医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1530",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆桐城新渡镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1530/1530-安庆桐城新渡镇店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1530/1530-安庆桐城新渡镇店-经营许可证.jpeg"
    }
  },
  {
    "id": "2771",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆桐城盛百店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1a36ec20-9b76-4c46-b350-85c160efa7441757697429712.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/19445d0c-28cf-42d6-81a1-42bc7c602edc1757697430543.jpg"
    }
  },
  {
    "id": "1924",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆桐城范岗镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1924/1924-安庆桐城范岗镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1924/1924-安庆桐城范岗镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1591",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆潜山南岳路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1591/1591-安庆潜山南岳路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1591/1591-安庆潜山南岳路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1592",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆潜山恒太城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1592/1592-安庆潜山恒太城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1592/1592-安庆潜山恒太城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1978",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆潜山源潭镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1978/1978-安庆潜山源潭镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1978/1978-安庆潜山源潭镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2798",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆火车站广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a5b2df18-338c-4d4d-89c0-731710a280311757697492808.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/24a47954-30bc-497f-8fc5-0b1a7d0f84881757697493504.jpg"
    }
  },
  {
    "id": "2777",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆碧桂园红星店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e44bd18b-24dd-4d8e-b222-9c5acf1206991757697442354.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d7c886e0-0bba-40c8-8052-f7895f4fdc241757697442743.jpg"
    }
  },
  {
    "id": "2805",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆绿地紫峰店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2e0d4382-c1df-49eb-8296-bc343d0c6bd81757697499280.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6885596b-a3bf-420f-9ed2-b2be072377121757697499695.jpg"
    }
  },
  {
    "id": "2773",
    "province": "安徽省",
    "city": "安庆市",
    "name": "安庆邻里汇店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4e4af2fa-34f8-4c0c-9ec2-e70eeeac90f71757697433230.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1f63fd50-4478-4202-bc33-63d8e60c73141757697433630.jpg"
    }
  },
  {
    "id": "2248",
    "province": "安徽省",
    "city": "安庆市",
    "name": "岳西县天鹅广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5dad2597-ed75-460a-8540-d5a494d0a6791757696764422.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/72c35697-110c-4487-a75b-5d83e4356efd1757696764771.jpg"
    }
  },
  {
    "id": "2809",
    "province": "安徽省",
    "city": "安庆市",
    "name": "桐城同悦广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/403566c7-6d01-4de7-9b30-226bc00439741757697503632.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a9fc1117-9278-4d6c-97e6-21a2573b194f1757697504187.jpg"
    }
  },
  {
    "id": "2034",
    "province": "安徽省",
    "city": "安庆市",
    "name": "桐城市文城东路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5739c37f-418d-4012-b3f6-09082f439a191757696478520.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a398953c-dd88-40ae-946f-b795bfc4dfc71757696479224.jpg"
    }
  },
  {
    "id": "2240",
    "province": "安徽省",
    "city": "安庆市",
    "name": "桐城市盛唐店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/53676f81-1d6e-419e-9386-c6c7891961e81757696747394.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9d6fb992-6c9a-4e1a-9e36-0c3599a42dfb1757696747775.jpg"
    }
  },
  {
    "id": "1612",
    "province": "安徽省",
    "city": "安庆市",
    "name": "潜山县潜阳路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1612/1612-潜山县潜阳路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1612/1612-潜山县潜阳路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1818",
    "province": "安徽省",
    "city": "安庆市",
    "name": "潜山市立医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1818/1818-潜山市立医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1818/1818-潜山市立医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1170",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城东方燕园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1170/1170-宣城东方燕园店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1170/1170-宣城东方燕园店-经营许可证.jpeg"
    }
  },
  {
    "id": "1032",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城中心医院餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1032/1032-宣城中心医院餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1032/1032-宣城中心医院餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1128",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1128/1128-宣城人民医院店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1128/1128-宣城人民医院店-经营许可证.png"
    }
  },
  {
    "id": "1006",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城大润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1006/1006-宣城大润发餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1006/1006-宣城大润发餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1008",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城宁国人民医院餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1008/1008-宣城宁国人民医院餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1008/1008-宣城宁国人民医院餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1005",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城宁国宁阳路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1005/1005-宣城宁国宁阳路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1005/1005-宣城宁国宁阳路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1725",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城广德升平街鼓角楼店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1725/1725-宣城广德升平街鼓角楼店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1725/1725-宣城广德升平街鼓角楼店-经营许可证.jpeg"
    }
  },
  {
    "id": "1143",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城广德商贸中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1143/1143-宣城广德商贸中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1143/1143-宣城广德商贸中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1041",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城广德大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1041/1041-宣城广德大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1041/1041-宣城广德大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1345",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城广德百大中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1345/1345-宣城广德百大中心店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1345/1345-宣城广德百大中心店-经营许可证.png"
    }
  },
  {
    "id": "1007",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城旌德港德广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1007/1007-宣城旌德港德广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1007/1007-宣城旌德港德广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1154",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城泾县环球缤纷城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1154/1154-宣城泾县环球缤纷城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1154/1154-宣城泾县环球缤纷城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1624",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城洪林服务区北区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1624/1624-宣城洪林服务区北区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1624/1624-宣城洪林服务区北区店-经营许可证.jpg"
    }
  },
  {
    "id": "1623",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城洪林服务区南区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1623/1623-宣城洪林服务区南区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1623/1623-宣城洪林服务区南区店-经营许可证.jpg"
    }
  },
  {
    "id": "1003",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城郎溪县分公司餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1003/1003-宣城郎溪县分公司餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1003/1003-宣城郎溪县分公司餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1090",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城郎溪台客隆店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1090/1090-宣城郎溪台客隆店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1090/1090-宣城郎溪台客隆店-经营许可证.jpeg"
    }
  },
  {
    "id": "1018",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城锦城北路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1018/1018-宣城锦城北路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1018/1018-宣城锦城北路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1144",
    "province": "安徽省",
    "city": "宣城市",
    "name": "宣城麦莎广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1144/1144-宣城麦莎广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1144/1144-宣城麦莎广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1122",
    "province": "安徽省",
    "city": "宣城市",
    "name": "泾县新世界店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1122/1122-泾县新世界店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1122/1122-泾县新世界店-经营许可证.jpeg"
    }
  },
  {
    "id": "1306",
    "province": "安徽省",
    "city": "宣城市",
    "name": "特许店-宣城国购广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1306/1306-特许店-宣城国购广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1306/1306-特许店-宣城国购广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1509",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州万达二餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1509/1509-宿州万达二餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1509/1509-宿州万达二餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1566",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州利群时代广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1566/1566-宿州利群时代广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1566/1566-宿州利群时代广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1767",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州天鹅湾店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1767/1767-宿州天鹅湾店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1767/1767-宿州天鹅湾店-经营许可证.jpg"
    }
  },
  {
    "id": "1507",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州市万达餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1507/1507-宿州市万达餐厅-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1507/1507-宿州市万达餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1650",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州市国购广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1650/1650-宿州市国购广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1650/1650-宿州市国购广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1765",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州市立医院北区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1765/1765-宿州市立医院北区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1765/1765-宿州市立医院北区店-经营许可证.jpg"
    }
  },
  {
    "id": "1511",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州拂晓广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1511/1511-宿州拂晓广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1511/1511-宿州拂晓广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1512",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州新一佳餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1512/1512-宿州新一佳餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1512/1512-宿州新一佳餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1757",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州汴河丽景店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1757/1757-宿州汴河丽景店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1757/1757-宿州汴河丽景店-经营许可证.jpg"
    }
  },
  {
    "id": "1510",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州汴河路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1510/1510-宿州汴河路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1510/1510-宿州汴河路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1017",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州泗县同辉广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1017/1017-宿州泗县同辉广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1017/1017-宿州泗县同辉广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1077",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州泗县清水湾餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1077/1077-宿州泗县清水湾餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1077/1077-宿州泗县清水湾餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1021",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州泗县玉兰大道店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1021/1021-宿州泗县玉兰大道店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1021/1021-宿州泗县玉兰大道店-经营许可证.jpg"
    }
  },
  {
    "id": "1076",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州泗县盛世豪庭餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1076/1076-宿州泗县盛世豪庭餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1076/1076-宿州泗县盛世豪庭餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1514",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州淮海学府餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1514/1514-宿州淮海学府餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1514/1514-宿州淮海学府餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1019",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州灵璧茂和广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1019/1019-宿州灵璧茂和广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1019/1019-宿州灵璧茂和广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1030",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州灵璧莱迪广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1030/1030-宿州灵璧莱迪广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1030/1030-宿州灵璧莱迪广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1902",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州皖北总院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1902/1902-宿州皖北总院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1902/1902-宿州皖北总院店-经营许可证.jpg"
    }
  },
  {
    "id": "1157",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州砀山万达餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1157/1157-宿州砀山万达餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1157/1157-宿州砀山万达餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1015",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州砀山不夜城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1015/1015-宿州砀山不夜城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1015/1015-宿州砀山不夜城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1026",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州砀山县大润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1026/1026-宿州砀山县大润发餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1026/1026-宿州砀山县大润发餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1508",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州磬云路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1508/1508-宿州磬云路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1508/1508-宿州磬云路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1338",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州符离镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1338/1338-宿州符离镇店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1338/1338-宿州符离镇店-经营许可证.jpeg"
    }
  },
  {
    "id": "1754",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州苏宁店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1754/1754-宿州苏宁店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1754/1754-宿州苏宁店-经营许可证.jpg"
    }
  },
  {
    "id": "1020",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州萧县亿洲城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1020/1020-宿州萧县亿洲城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1020/1020-宿州萧县亿洲城餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1513",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州金方世纪城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1513/1513-宿州金方世纪城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1513/1513-宿州金方世纪城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1271",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州高速驿达符离服务区东区餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1271/1271-宿州高速驿达符离服务区东区餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1271/1271-宿州高速驿达符离服务区东区餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1272",
    "province": "安徽省",
    "city": "宿州市",
    "name": "宿州高速驿达符离服务区西区餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1272/1272-宿州高速驿达符离服务区西区餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1272/1272-宿州高速驿达符离服务区西区餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1702",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州万成香格里拉店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1702/1702-池州万成香格里拉店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1702/1702-池州万成香格里拉店-经营许可证.jpeg"
    }
  },
  {
    "id": "1704",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州万盛广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1704/1704-池州万盛广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1704/1704-池州万盛广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1547",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州东至大渡口镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1547/1547-池州东至大渡口镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1547/1547-池州东至大渡口镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1700",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州东至新天地广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1700/1700-池州东至新天地广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1700/1700-池州东至新天地广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1699",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州东至汇金广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1699/1699-池州东至汇金广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1699/1699-池州东至汇金广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1583",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州东至花园服务区东区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1583/1583-池州东至花园服务区东区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1583/1583-池州东至花园服务区东区店-经营许可证.jpg"
    }
  },
  {
    "id": "1584",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州东至花园服务区西区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1584/1584-池州东至花园服务区西区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1584/1584-池州东至花园服务区西区店-经营许可证.jpg"
    }
  },
  {
    "id": "1353",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州九华山游客服务中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1353/1353-池州九华山游客服务中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1353/1353-池州九华山游客服务中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1273",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州升金湖服务区东餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1273/1273-池州升金湖服务区东餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1273/1273-池州升金湖服务区东餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1274",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州升金湖服务区西餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1274/1274-池州升金湖服务区西餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1274/1274-池州升金湖服务区西餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1703",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州商之都店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1703/1703-池州商之都店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1703/1703-池州商之都店-经营许可证.jpg"
    }
  },
  {
    "id": "1701",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州远东国际广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1701/1701-池州远东国际广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1701/1701-池州远东国际广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1882",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州青阳城上城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1882/1882-池州青阳城上城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1882/1882-池州青阳城上城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1705",
    "province": "安徽省",
    "city": "池州市",
    "name": "池州青阳大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1705/1705-池州青阳大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1705/1705-池州青阳大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1505",
    "province": "安徽省",
    "city": "淮北市",
    "name": "淮北南翔云集餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1505/1505-淮北南翔云集餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1505/1505-淮北南翔云集餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1814",
    "province": "安徽省",
    "city": "淮北市",
    "name": "淮北吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1814/1814-淮北吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1814/1814-淮北吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1534",
    "province": "安徽省",
    "city": "淮北市",
    "name": "淮北安邦广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1534/1534-淮北安邦广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1534/1534-淮北安邦广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1778",
    "province": "安徽省",
    "city": "淮北市",
    "name": "淮北恒大中央公园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1778/1778-淮北恒大中央公园店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1778/1778-淮北恒大中央公园店-经营许可证.jpeg"
    }
  },
  {
    "id": "1506",
    "province": "安徽省",
    "city": "淮北市",
    "name": "淮北濉溪新百餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1506/1506-淮北濉溪新百餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1506/1506-淮北濉溪新百餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1602",
    "province": "安徽省",
    "city": "淮北市",
    "name": "淮北碧乐城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1602/1602-淮北碧乐城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1602/1602-淮北碧乐城店-经营许可证.jpg"
    }
  },
  {
    "id": "1503",
    "province": "安徽省",
    "city": "淮北市",
    "name": "淮北老乡鸡三餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1503/1503-淮北老乡鸡三餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1503/1503-淮北老乡鸡三餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1504",
    "province": "安徽省",
    "city": "淮北市",
    "name": "老乡鸡淮北步行街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1504/1504-老乡鸡淮北步行街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1504/1504-老乡鸡淮北步行街店-经营许可证.jpg"
    }
  },
  {
    "id": "1541",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南万茂餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1541/1541-淮南万茂餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1541/1541-淮南万茂餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1452",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南万达广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1452/1452-淮南万达广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1452/1452-淮南万达广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1619",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南三和镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1619/1619-淮南三和镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1619/1619-淮南三和镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1444",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南上东锦城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1444/1444-淮南上东锦城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1444/1444-淮南上东锦城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1752",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南东方医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1752/1752-淮南东方医院店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1752/1752-淮南东方医院店-经营许可证.jpeg"
    }
  },
  {
    "id": "1443",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南中化国际城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1443/1443-淮南中化国际城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1443/1443-淮南中化国际城餐厅-经营许可证.png"
    }
  },
  {
    "id": "1644",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南八公山服务区北区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1644/1644-淮南八公山服务区北区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1644/1644-淮南八公山服务区北区店-经营许可证.jpg"
    }
  },
  {
    "id": "1643",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南八公山服务区南区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1643/1643-淮南八公山服务区南区店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1643/1643-淮南八公山服务区南区店-经营许可证.jpeg"
    }
  },
  {
    "id": "1437",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南八公山餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1437/1437-淮南八公山餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1437/1437-淮南八公山餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1789",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南凤凰湾店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1789/1789-淮南凤凰湾店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1789/1789-淮南凤凰湾店-经营许可证.jpg"
    }
  },
  {
    "id": "1900",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南凤台中山北路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1900/1900-淮南凤台中山北路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1900/1900-淮南凤台中山北路店-经营许可证.jpg"
    }
  },
  {
    "id": "1440",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南凤台明珠大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1440/1440-淮南凤台明珠大道餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1440/1440-淮南凤台明珠大道餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1439",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南凤台未来城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1439/1439-淮南凤台未来城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1439/1439-淮南凤台未来城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1553",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南天一时代城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1553/1553-淮南天一时代城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1553/1553-淮南天一时代城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1586",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南天柱山路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1586/1586-淮南天柱山路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1586/1586-淮南天柱山路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1434",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南寿县二餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1434/1434-淮南寿县二餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1434/1434-淮南寿县二餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1431",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南寿县大顺路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1431/1431-淮南寿县大顺路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1431/1431-淮南寿县大顺路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1432",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南寿县玫瑰公馆餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1432/1432-淮南寿县玫瑰公馆餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1432/1432-淮南寿县玫瑰公馆餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1A02",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南寿县环球港店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A02/1A02-淮南寿县环球港店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A02/1A02-淮南寿县环球港店-经营许可证.jpeg"
    }
  },
  {
    "id": "1551",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南寿县鼎鑫幸福城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1551/1551-淮南寿县鼎鑫幸福城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1551/1551-淮南寿县鼎鑫幸福城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1433",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南尚泰广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1433/1433-淮南尚泰广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1433/1433-淮南尚泰广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1565",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南山南印象店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1565/1565-淮南山南印象店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1565/1565-淮南山南印象店-经营许可证.jpeg"
    }
  },
  {
    "id": "1453",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南市人民医院餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1453/1453-淮南市人民医院餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1453/1453-淮南市人民医院餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1447",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南市淮河新城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1447/1447-淮南市淮河新城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1447/1447-淮南市淮河新城餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1445",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南广场路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1445/1445-淮南广场路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1445/1445-淮南广场路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1446",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南惠利花园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1446/1446-淮南惠利花园店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1446/1446-淮南惠利花园店-经营许可证.jpg"
    }
  },
  {
    "id": "1351",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南毛集店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1351/1351-淮南毛集店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1351/1351-淮南毛集店-经营许可证.jpeg"
    }
  },
  {
    "id": "1454",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南泉山湖餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1454/1454-淮南泉山湖餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1454/1454-淮南泉山湖餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1442",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南潘集世瑞大厦餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1442/1442-淮南潘集世瑞大厦餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1442/1442-淮南潘集世瑞大厦餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1798",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南潘集珠江路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1798/1798-淮南潘集珠江路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1798/1798-淮南潘集珠江路店-经营许可证.jpg"
    }
  },
  {
    "id": "1450",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南火车站餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1450/1450-淮南火车站餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1450/1450-淮南火车站餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1589",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南科技大厦餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1589/1589-淮南科技大厦餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1589/1589-淮南科技大厦餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1796",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南绿茵里店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1796/1796-淮南绿茵里店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1796/1796-淮南绿茵里店-经营许可证.jpeg"
    }
  },
  {
    "id": "1919",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南英伦联邦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1919/1919-淮南英伦联邦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1919/1919-淮南英伦联邦店-经营许可证.jpg"
    }
  },
  {
    "id": "1436",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南蔡新路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1436/1436-淮南蔡新路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1436/1436-淮南蔡新路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1435",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南西城国际店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1435/1435-淮南西城国际店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1435/1435-淮南西城国际店-经营许可证.jpeg"
    }
  },
  {
    "id": "1712",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南谢家集祥云府店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1712/1712-淮南谢家集祥云府店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1712/1712-淮南谢家集祥云府店-经营许可证.jpg"
    }
  },
  {
    "id": "1621",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南迎河服务区东区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1621/1621-淮南迎河服务区东区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1621/1621-淮南迎河服务区东区店-经营许可证.jpg"
    }
  },
  {
    "id": "1622",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南迎河服务区西区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1622/1622-淮南迎河服务区西区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1622/1622-淮南迎河服务区西区店-经营许可证.jpg"
    }
  },
  {
    "id": "1769",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南金域蓝湾店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1769/1769-淮南金域蓝湾店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1769/1769-淮南金域蓝湾店-经营许可证.jpeg"
    }
  },
  {
    "id": "1645",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南香樟苑店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1645/1645-淮南香樟苑店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1645/1645-淮南香樟苑店-经营许可证.jpeg"
    }
  },
  {
    "id": "1448",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南龙湖中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1448/1448-淮南龙湖中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1448/1448-淮南龙湖中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1449",
    "province": "安徽省",
    "city": "淮南市",
    "name": "淮南龙湖路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1449/1449-淮南龙湖路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1449/1449-淮南龙湖路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1467",
    "province": "安徽省",
    "city": "滁州市",
    "name": "全椒城东花园餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1467/1467-全椒城东花园餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1467/1467-全椒城东花园餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1892",
    "province": "安徽省",
    "city": "滁州市",
    "name": "凤阳临淮关镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1892/1892-凤阳临淮关镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1892/1892-凤阳临淮关镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1022",
    "province": "安徽省",
    "city": "滁州市",
    "name": "天长天发广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1022/1022-天长天发广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1022/1022-天长天发广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1049",
    "province": "安徽省",
    "city": "滁州市",
    "name": "来安嘉年华店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1049/1049-来安嘉年华店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1049/1049-来安嘉年华店-经营许可证.jpg"
    }
  },
  {
    "id": "1922",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州七彩联华超市店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1922/1922-滁州七彩联华超市店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1922/1922-滁州七彩联华超市店-经营许可证.jpg"
    }
  },
  {
    "id": "1A12",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州东升花园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A12/1A12-滁州东升花园店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A12/1A12-滁州东升花园店-经营许可证.jpg"
    }
  },
  {
    "id": "1461",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州中州国际广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1461/1461-滁州中州国际广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1461/1461-滁州中州国际广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1463",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州中都大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1463/1463-滁州中都大道餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1463/1463-滁州中都大道餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1593",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州丰乐大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1593/1593-滁州丰乐大道餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1593/1593-滁州丰乐大道餐厅-经营许可证.png"
    }
  },
  {
    "id": "1822",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州乐彩城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1822/1822-滁州乐彩城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1822/1822-滁州乐彩城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1283",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州全椒十字店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1283/1283-滁州全椒十字店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1283/1283-滁州全椒十字店-经营许可证.jpg"
    }
  },
  {
    "id": "1470",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州全椒县城南大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1470/1470-滁州全椒县城南大道餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1470/1470-滁州全椒县城南大道餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1469",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州全椒新江海城市广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1469/1469-滁州全椒新江海城市广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1469/1469-滁州全椒新江海城市广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1277",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州全椒服务区北餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1277/1277-滁州全椒服务区北餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1277/1277-滁州全椒服务区北餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1278",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州全椒服务区南餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1278/1278-滁州全椒服务区南餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1278/1278-滁州全椒服务区南餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1471",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州全椒站前广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1471/1471-滁州全椒站前广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1471/1471-滁州全椒站前广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1567",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州凤阳商贸城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1567/1567-滁州凤阳商贸城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1567/1567-滁州凤阳商贸城店-经营许可证.jpg"
    }
  },
  {
    "id": "1456",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州凤阳联华超市店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1456/1456-滁州凤阳联华超市店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1456/1456-滁州凤阳联华超市店-经营许可证.jpg"
    }
  },
  {
    "id": "1606",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州凯迪置地广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1606/1606-滁州凯迪置地广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1606/1606-滁州凯迪置地广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1590",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1590/1590-滁州吾悦广场店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1590/1590-滁州吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1016",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州天长吾悦广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1016/1016-滁州天长吾悦广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1016/1016-滁州天长吾悦广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1027",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州天长市苏果店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1027/1027-滁州天长市苏果店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1027/1027-滁州天长市苏果店-经营许可证.jpg"
    }
  },
  {
    "id": "1777",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州天长秦栏镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1777/1777-滁州天长秦栏镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1777/1777-滁州天长秦栏镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1458",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州定远县餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1458/1458-滁州定远县餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1458/1458-滁州定远县餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1637",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州定远大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1637/1637-滁州定远大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1637/1637-滁州定远大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1457",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州定远曲阳国际餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1457/1457-滁州定远曲阳国际餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1457/1457-滁州定远曲阳国际餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1409",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州定远炉桥镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1409/1409-滁州定远炉桥镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1409/1409-滁州定远炉桥镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1459",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州定远金山丽景餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1459/1459-滁州定远金山丽景餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1459/1459-滁州定远金山丽景餐厅-经营许可证.png"
    }
  },
  {
    "id": "1753",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州定远金鹏玖玖店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1753/1753-滁州定远金鹏玖玖店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1753/1753-滁州定远金鹏玖玖店-经营许可证.jpeg"
    }
  },
  {
    "id": "1964",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州市全椒店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1964/1964-滁州市全椒店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1964/1964-滁州市全椒店-经营许可证.jpeg"
    }
  },
  {
    "id": "1124",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州明光名都汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1124/1124-滁州明光名都汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1124/1124-滁州明光名都汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1029",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州明光大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1029/1029-滁州明光大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1029/1029-滁州明光大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1031",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州明光市餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1031/1031-滁州明光市餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1031/1031-滁州明光市餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1023",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州来安世纪华联店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1023/1023-滁州来安世纪华联店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1023/1023-滁州来安世纪华联店-经营许可证.jpg"
    }
  },
  {
    "id": "1411",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州来安汊河店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1411/1411-滁州来安汊河店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1411/1411-滁州来安汊河店-经营许可证.jpg"
    }
  },
  {
    "id": "1053",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州来安苏润国际餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1053/1053-滁州来安苏润国际餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1053/1053-滁州来安苏润国际餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1903",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州永乐北路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1903/1903-滁州永乐北路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1903/1903-滁州永乐北路店-经营许可证.jpg"
    }
  },
  {
    "id": "1676",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州白云商厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1676/1676-滁州白云商厦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1676/1676-滁州白云商厦店-经营许可证.jpg"
    }
  },
  {
    "id": "1466",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州紫金广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1466/1466-滁州紫金广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1466/1466-滁州紫金广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1460",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州金鹏琅琊玖玖广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1460/1460-滁州金鹏琅琊玖玖广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1460/1460-滁州金鹏琅琊玖玖广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1542",
    "province": "安徽省",
    "city": "滁州市",
    "name": "滁州龙蟠大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1542/1542-滁州龙蟠大道餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1542/1542-滁州龙蟠大道餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "2339",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "无为县君临四季花都店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d0cde243-5745-4629-94f5-c46c137efd501757696900389.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/06625cc5-dfb5-4c04-9046-98247b9a565d1757696900758.jpg"
    }
  },
  {
    "id": "2183",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "无为安得利店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/16d68140-e878-4f32-b7ef-578d233134151757696680970.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3ae29761-c3f5-425b-bb5c-9ed465b9c4971757696681357.jpg"
    }
  },
  {
    "id": "1973",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "无为竹丝湖服务区东店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1973/1973-无为竹丝湖服务区东店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1973/1973-无为竹丝湖服务区东店-经营许可证.jpg"
    }
  },
  {
    "id": "1974",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "无为竹丝湖服务区西店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1974/1974-无为竹丝湖服务区西店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1974/1974-无为竹丝湖服务区西店-经营许可证.jpg"
    }
  },
  {
    "id": "2502",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "无为米芾广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8c090b30-c31f-4284-ba76-81b767e3d6d61757697130047.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a294867d-380d-4df2-be64-07ace28a66771757697130421.jpg"
    }
  },
  {
    "id": "2446",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖万春花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ae635cdd-4b99-4036-a2a8-8cd4a5349c9f1757697063799.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/246a48cf-35dd-43b8-be2d-4c9da8803bae1757697064167.jpg"
    }
  },
  {
    "id": "2679",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖万达广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3ac7185a-6185-4d62-afba-03378b73b2c81764867820627.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f00c4808-3426-409b-b5e9-136e0b6a2de51764867821370.jpg"
    }
  },
  {
    "id": "2569",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖东方龙城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eaf2249f-4fcf-4b02-8d33-2dd6e362f5891757697223788.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6bcf9f0c-089c-4624-a40d-91e65c5c3d601757697224475.jpg"
    }
  },
  {
    "id": "2391",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖中央城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f2f9dad4-4598-4188-8721-93d1c48355001757696986301.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c514382f-8b82-477d-9e3e-c910a59790ae1757696986690.jpg"
    }
  },
  {
    "id": "2585",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖伟星万悦城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7de13dee-ac52-4622-bdd7-f123957f01f41757697236023.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/97e26d0b-dcdf-4bcb-9b13-03931991c8401757697236398.jpg"
    }
  },
  {
    "id": "2609",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖伟星时代之光店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f5e02fda-2989-4ef6-a3ca-ef2fe00fe2421757697257787.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/df0ae624-303c-4f84-837e-c89fc6129fab1757697258186.jpg"
    }
  },
  {
    "id": "2714",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖伟星玖璋台店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a4c1ab4a-0b14-48be-9e5c-6b0c164af5ae1769015029013.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/39447273-c968-420c-b5e4-82183ea496001769015029376.jpg"
    }
  },
  {
    "id": "2723",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖伟星银湖时代店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/46c5508e-9265-4349-ae3a-31e4bc2224661757697368044.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e304b0d7-fc3b-4e5a-9074-d523d236086d1757697368418.jpg"
    }
  },
  {
    "id": "2221",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖南瑞新城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a3c7e154-0d04-4a01-844c-704924a52c831757696721736.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cd17bfbf-5f40-473c-98d2-5a69f4786a991757696722094.jpg"
    }
  },
  {
    "id": "1575",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖南陵许镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1575/1575-芜湖南陵许镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1575/1575-芜湖南陵许镇店-经营许可证.jpg"
    }
  },
  {
    "id": "2408",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖大桥新城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/52f837cd-f728-4aaa-9b38-91774be9487f1757697009700.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2da52106-c794-402b-97c0-135579430c751757697010074.jpg"
    }
  },
  {
    "id": "2412",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖大润发店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b2b73810-fe1a-4ea0-b108-b5a0710afa301757697017494.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/412ea727-079c-40cc-8398-72d050b717441757697017899.jpg"
    }
  },
  {
    "id": "2407",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖大观花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/487ffd80-eac7-4f26-b07a-ea4f95a51c051757697007014.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f7837440-d8fe-448b-8b51-a1193ae2b8f11757697007377.jpg"
    }
  },
  {
    "id": "2386",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖幸福里店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/984e9fe4-5b3e-4ddd-b9e7-78981e3378e71757696980357.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/58b7be6b-f2b7-4b75-8b54-19ea669cd7f41757696980720.jpg"
    }
  },
  {
    "id": "1941",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖悦达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1941/1941-芜湖悦达广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1941/1941-芜湖悦达广场店-经营许可证.jpg"
    }
  },
  {
    "id": "2464",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖新时代商业街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8864769a-2815-431e-9795-821b4b6713cd1757697089232.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/638f7811-85d7-43f7-b29d-95190421c1241757697089594.jpg"
    }
  },
  {
    "id": "2118",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖新芜路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cb86ef88-2873-4d74-95be-42f0d6166ceb1769187660867.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a45c3dc4-0f93-4495-b840-c22f6562ca4a1769187661246.jpg"
    }
  },
  {
    "id": "2423",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖星隆国际店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bc82c8df-e91f-48dd-874d-2601f698b6a81756832659271.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/994ab4e7-bfdf-4c55-86bd-78c652c6afa91756832659914.jpg"
    }
  },
  {
    "id": "2553",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖柏庄时代广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/87d4a174-8730-490a-b4d0-afc23543fa341757697205314.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1b5adf01-2819-4cde-b3d4-752400f74d291757697205697.jpg"
    }
  },
  {
    "id": "1305",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖沈巷店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1305/1305-芜湖沈巷店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1305/1305-芜湖沈巷店-经营许可证.jpg"
    }
  },
  {
    "id": "1956",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖湾址城东新城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1956/1956-芜湖湾址城东新城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1956/1956-芜湖湾址城东新城店-经营许可证.jpg"
    }
  },
  {
    "id": "1942",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖湾沚静安阳光城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1942/1942-芜湖湾沚静安阳光城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1942/1942-芜湖湾沚静安阳光城店-经营许可证.jpg"
    }
  },
  {
    "id": "2214",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖澳然天成店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d5d008a2-e7b7-4ce7-8a96-8dc63cea51d81757696715547.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/da66a3df-5872-439c-aec5-0ee200642db11757696716271.jpg"
    }
  },
  {
    "id": "2142",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖百线广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/64dde005-d5ae-42cb-b51a-07506c975b9a1757696619815.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/711a6ecd-813d-47cc-a237-2d25052b23ed1757696620179.jpg"
    }
  },
  {
    "id": "2276",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖繁昌华亿广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/79ef7f9a-232e-48a0-ad29-2ae40a6d4de71769187713824.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b04a192a-e46f-41c5-97e5-d87189dabb2b1769187714181.jpg"
    }
  },
  {
    "id": "2884",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖重庆路宜邻中心店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/营1773925842165.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/许1773925848965.jpg"
    }
  },
  {
    "id": "2621",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖鲁港大市场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cfbf923c-62ef-4198-a5a7-cef14a923a6b1757697266876.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ab87189a-6550-45a5-9044-b49be297d4601757697267251.jpg"
    }
  },
  {
    "id": "2264",
    "province": "安徽省",
    "city": "芜湖市",
    "name": "芜湖黄山东路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e6201195-3533-491b-86e3-2f505d9a40d11772470907024.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c766f0b8-0637-46a0-bc74-0892eba76b311772470907399.jpg"
    }
  },
  {
    "id": "1055",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "怀远大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1055/1055-怀远大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1055/1055-怀远大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1147",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠万方新都汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1147/1147-蚌埠万方新都汇店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1147/1147-蚌埠万方新都汇店-经营许可证.png"
    }
  },
  {
    "id": "1628",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠五河人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1628/1628-蚌埠五河人民医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1628/1628-蚌埠五河人民医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1028",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠五河彩虹时代广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1028/1028-蚌埠五河彩虹时代广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1028/1028-蚌埠五河彩虹时代广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1056",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠兴业街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1056/1056-蚌埠兴业街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1056/1056-蚌埠兴业街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1341",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠合家福店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1341/1341-蚌埠合家福店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1341/1341-蚌埠合家福店-经营许可证.jpeg"
    }
  },
  {
    "id": "1844",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1844/1844-蚌埠吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1844/1844-蚌埠吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1068",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠和顺名都城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1068/1068-蚌埠和顺名都城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1068/1068-蚌埠和顺名都城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1052",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠商之都餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1052/1052-蚌埠商之都餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1052/1052-蚌埠商之都餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1138",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠固镇新天地店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1138/1138-蚌埠固镇新天地店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1138/1138-蚌埠固镇新天地店-经营许可证.jpg"
    }
  },
  {
    "id": "1131",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠固镇百大购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1131/1131-蚌埠固镇百大购物中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1131/1131-蚌埠固镇百大购物中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1137",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠国购广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1137/1137-蚌埠国购广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1137/1137-蚌埠国购广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1965",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠大学城融实购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1965/1965-蚌埠大学城融实购物中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1965/1965-蚌埠大学城融实购物中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1062",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠工农路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1062/1062-蚌埠工农路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1062/1062-蚌埠工农路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1057",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠怀远新河路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1057/1057-蚌埠怀远新河路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1057/1057-蚌埠怀远新河路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1127",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠怀远泰谷玖街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1127/1127-蚌埠怀远泰谷玖街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1127/1127-蚌埠怀远泰谷玖街店-经营许可证.jpeg"
    }
  },
  {
    "id": "1044",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠拓基餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1044/1044-蚌埠拓基餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1044/1044-蚌埠拓基餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1047",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠文化广场大润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1047/1047-蚌埠文化广场大润发餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1047/1047-蚌埠文化广场大润发餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1133",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠明珠大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1133/1133-蚌埠明珠大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1133/1133-蚌埠明珠大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1132",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠永昌国际大厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1132/1132-蚌埠永昌国际大厦店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1132/1132-蚌埠永昌国际大厦店-经营许可证.jpeg"
    }
  },
  {
    "id": "1051",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠淮上万达餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1051/1051-蚌埠淮上万达餐厅-营业执照.jpg",
      "food": ""
    }
  },
  {
    "id": "1107",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠绿地珠峰店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1107/1107-蚌埠绿地珠峰店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1107/1107-蚌埠绿地珠峰店-经营许可证.jpg"
    }
  },
  {
    "id": "1896",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠蚌医一附院东门店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1896/1896-蚌埠蚌医一附院东门店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1896/1896-蚌埠蚌医一附院东门店-经营许可证.jpg"
    }
  },
  {
    "id": "1111",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠蚌山万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1111/1111-蚌埠蚌山万达店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1111/1111-蚌埠蚌山万达店-经营许可证.jpeg"
    }
  },
  {
    "id": "1165",
    "province": "安徽省",
    "city": "蚌埠市",
    "name": "蚌埠鼎元府邸餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1165/1165-蚌埠鼎元府邸餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1165/1165-蚌埠鼎元府邸餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1717",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵五环店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1717/1717-铜陵五环店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1717/1717-铜陵五环店-经营许可证.jpeg"
    }
  },
  {
    "id": "1690",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵县店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1690/1690-铜陵县店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1690/1690-铜陵县店-经营许可证.jpg"
    }
  },
  {
    "id": "1813",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1813/1813-铜陵吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1813/1813-铜陵吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1823",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵嘉华国际广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1823/1823-铜陵嘉华国际广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1823/1823-铜陵嘉华国际广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1693",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵嘉禾广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1693/1693-铜陵嘉禾广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1693/1693-铜陵嘉禾广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1692",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵天润嘉园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1692/1692-铜陵天润嘉园店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1692/1692-铜陵天润嘉园店-经营许可证.jpeg"
    }
  },
  {
    "id": "1696",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵市万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1696/1696-铜陵市万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1696/1696-铜陵市万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1689",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵枞阳港城广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1689/1689-铜陵枞阳港城广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1689/1689-铜陵枞阳港城广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1806",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵枞阳逸龙山庄店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1806/1806-铜陵枞阳逸龙山庄店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1806/1806-铜陵枞阳逸龙山庄店-经营许可证.jpg"
    }
  },
  {
    "id": "1909",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵柏庄香域店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1909/1909-铜陵柏庄香域店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1909/1909-铜陵柏庄香域店-经营许可证.jpg"
    }
  },
  {
    "id": "1694",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵石城大道店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1694/1694-铜陵石城大道店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1694/1694-铜陵石城大道店-经营许可证.jpeg"
    }
  },
  {
    "id": "1698",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵西湖春城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1698/1698-铜陵西湖春城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1698/1698-铜陵西湖春城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1697",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵长江中路乐都店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1697/1697-铜陵长江中路乐都店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1697/1697-铜陵长江中路乐都店-经营许可证.jpeg"
    }
  },
  {
    "id": "1695",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵长江二路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1695/1695-铜陵长江二路店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1695/1695-铜陵长江二路店-经营许可证.jpeg"
    }
  },
  {
    "id": "1763",
    "province": "安徽省",
    "city": "铜陵市",
    "name": "铜陵顺安镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1763/1763-铜陵顺安镇店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1763/1763-铜陵顺安镇店-经营许可证.jpeg"
    }
  },
  {
    "id": "1488",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳万象城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1488/1488-阜阳万象城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1488/1488-阜阳万象城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1487",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳万达广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1487/1487-阜阳万达广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1487/1487-阜阳万达广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1617",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳临沂商城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1617/1617-阜阳临沂商城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1617/1617-阜阳临沂商城餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1066",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳临泉佳源东方餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1066/1066-阜阳临泉佳源东方餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1066/1066-阜阳临泉佳源东方餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1084",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳临泉华安城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1084/1084-阜阳临泉华安城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1084/1084-阜阳临泉华安城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1067",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳临泉大润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1067/1067-阜阳临泉大润发餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1067/1067-阜阳临泉大润发餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1064",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳临泉御园财富广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1064/1064-阜阳临泉御园财富广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1064/1064-阜阳临泉御园财富广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1492",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳丽丰一品餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1492/1492-阜阳丽丰一品餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1492/1492-阜阳丽丰一品餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1071",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳五洲万汇餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1071/1071-阜阳五洲万汇餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1071/1071-阜阳五洲万汇餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1707",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳双清湾水街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1707/1707-阜阳双清湾水街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1707/1707-阜阳双清湾水街店-经营许可证.jpg"
    }
  },
  {
    "id": "1500",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳商厦时代广场新餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1500/1500-阜阳商厦时代广场新餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1500/1500-阜阳商厦时代广场新餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1495",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳国贸商城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1495/1495-阜阳国贸商城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1495/1495-阜阳国贸商城餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1494",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳天瑞名城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1494/1494-阜阳天瑞名城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1494/1494-阜阳天瑞名城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1537",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳太和万达餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1537/1537-阜阳太和万达餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1537/1537-阜阳太和万达餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1681",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳太和印象城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1681/1681-阜阳太和印象城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1681/1681-阜阳太和印象城店-经营许可证.jpg"
    }
  },
  {
    "id": "1571",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳太和永辉店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1571/1571-阜阳太和永辉店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1571/1571-阜阳太和永辉店-经营许可证.jpeg"
    }
  },
  {
    "id": "1502",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳太和第五街区餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1502/1502-阜阳太和第五街区餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1502/1502-阜阳太和第五街区餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1501",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳太和长征北路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1501/1501-阜阳太和长征北路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1501/1501-阜阳太和长征北路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1745",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳如意豪庭店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1745/1745-阜阳如意豪庭店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1745/1745-阜阳如意豪庭店-经营许可证.jpeg"
    }
  },
  {
    "id": "1499",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳安医餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1499/1499-阜阳安医餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1499/1499-阜阳安医餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1498",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳市幸福公馆店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1498/1498-阜阳市幸福公馆店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1498/1498-阜阳市幸福公馆店-经营许可证.jpg"
    }
  },
  {
    "id": "1497",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳开乐广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1497/1497-阜阳开乐广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1497/1497-阜阳开乐广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1535",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳怡和广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1535/1535-阜阳怡和广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1535/1535-阜阳怡和广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1496",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳新五院餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1496/1496-阜阳新五院餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1496/1496-阜阳新五院餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1538",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳正基首府餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1538/1538-阜阳正基首府餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1538/1538-阜阳正基首府餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1493",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳汇美城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1493/1493-阜阳汇美城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1493/1493-阜阳汇美城餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1069",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳沣泽悦城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1069/1069-阜阳沣泽悦城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1069/1069-阜阳沣泽悦城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1490",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳清河东路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1490/1490-阜阳清河东路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1490/1490-阜阳清河东路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1129",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳界首万吉广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1129/1129-阜阳界首万吉广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1129/1129-阜阳界首万吉广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1088",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳界首人民路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1088/1088-阜阳界首人民路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1088/1088-阜阳界首人民路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1065",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳界首国祯广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1065/1065-阜阳界首国祯广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1065/1065-阜阳界首国祯广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1953",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳界首漫乐城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1953/1953-阜阳界首漫乐城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1953/1953-阜阳界首漫乐城店-经营许可证.jpg"
    }
  },
  {
    "id": "1485",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳祥源城公园餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1485/1485-阜阳祥源城公园餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1485/1485-阜阳祥源城公园餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1489",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳站前广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1489/1489-阜阳站前广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1489/1489-阜阳站前广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1562",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳聚隆美墅店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1562/1562-阜阳聚隆美墅店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1562/1562-阜阳聚隆美墅店-经营许可证.jpg"
    }
  },
  {
    "id": "1491",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳金悦广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1491/1491-阜阳金悦广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1491/1491-阜阳金悦广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1950",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳阜南万宇步行街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1950/1950-阜阳阜南万宇步行街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1950/1950-阜阳阜南万宇步行街店-经营许可证.jpeg"
    }
  },
  {
    "id": "1779",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳阜南中医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1779/1779-阜阳阜南中医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1779/1779-阜阳阜南中医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1536",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳阜南天筑广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1536/1536-阜阳阜南天筑广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1536/1536-阜阳阜南天筑广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1486",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳阜南曹集路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1486/1486-阜阳阜南曹集路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1486/1486-阜阳阜南曹集路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1081",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳颍上前进路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1081/1081-阜阳颍上前进路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1081/1081-阜阳颍上前进路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1082",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳颍上县人民路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1082/1082-阜阳颍上县人民路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1082/1082-阜阳颍上县人民路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1079",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳颍上县颖阳路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1079/1079-阜阳颍上县颖阳路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1079/1079-阜阳颍上县颖阳路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1070",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳颍上太平洋广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1070/1070-阜阳颍上太平洋广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1070/1070-阜阳颍上太平洋广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1582",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳颍上服务区北区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1582/1582-阜阳颍上服务区北区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1582/1582-阜阳颍上服务区北区店-经营许可证.jpg"
    }
  },
  {
    "id": "1581",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳颍上服务区南区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1581/1581-阜阳颍上服务区南区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1581/1581-阜阳颍上服务区南区店-经营许可证.jpg"
    }
  },
  {
    "id": "1876",
    "province": "安徽省",
    "city": "阜阳市",
    "name": "阜阳颍东发到家店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1876/1876-阜阳颍东发到家店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1876/1876-阜阳颍东发到家店-经营许可证.jpg"
    }
  },
  {
    "id": "1480",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山万达广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1480/1480-马鞍山万达广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1480/1480-马鞍山万达广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1539",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山东方明珠餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1539/1539-马鞍山东方明珠餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1539/1539-马鞍山东方明珠餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1540",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山伟星广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1540/1540-马鞍山伟星广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1540/1540-马鞍山伟星广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1479",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山军民路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1479/1479-马鞍山军民路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1479/1479-马鞍山军民路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1626",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山博望店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1626/1626-马鞍山博望店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1626/1626-马鞍山博望店-经营许可证.jpg"
    }
  },
  {
    "id": "1884",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山向山镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1884/1884-马鞍山向山镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1884/1884-马鞍山向山镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1842",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山含山县鑫乐广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1842/1842-马鞍山含山县鑫乐广场店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1842/1842-马鞍山含山县鑫乐广场店-经营许可证.png"
    }
  },
  {
    "id": "1482",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山含山天润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1482/1482-马鞍山含山天润发餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1482/1482-马鞍山含山天润发餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1570",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山含山昭关东路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1570/1570-马鞍山含山昭关东路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1570/1570-马鞍山含山昭关东路店-经营许可证.jpg"
    }
  },
  {
    "id": "1483",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山含山环峰西路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1483/1483-马鞍山含山环峰西路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1483/1483-马鞍山含山环峰西路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1285",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山和县乌江店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1285/1285-马鞍山和县乌江店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1285/1285-马鞍山和县乌江店-经营许可证.jpg"
    }
  },
  {
    "id": "1569",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山和县和州路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1569/1569-马鞍山和县和州路店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1569/1569-马鞍山和县和州路店-经营许可证.jpeg"
    }
  },
  {
    "id": "1741",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山和县安德利购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1741/1741-马鞍山和县安德利购物中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1741/1741-马鞍山和县安德利购物中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1598",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山哥伦布广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1598/1598-马鞍山哥伦布广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1598/1598-马鞍山哥伦布广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1714",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山大学城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1714/1714-马鞍山大学城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1714/1714-马鞍山大学城店-经营许可证.jpg"
    }
  },
  {
    "id": "1608",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山大润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1608/1608-马鞍山大润发餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1608/1608-马鞍山大润发餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1759",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山太白滨江汇商业中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1759/1759-马鞍山太白滨江汇商业中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1759/1759-马鞍山太白滨江汇商业中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1475",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山市和泰国际花园餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1475/1475-马鞍山市和泰国际花园餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1475/1475-马鞍山市和泰国际花园餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1476",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山市金色新天地餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1476/1476-马鞍山市金色新天地餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1476/1476-马鞍山市金色新天地餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1780",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山康泰佳苑店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1780/1780-马鞍山康泰佳苑店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1780/1780-马鞍山康泰佳苑店-经营许可证.jpg"
    }
  },
  {
    "id": "1881",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山当涂君悦华庭店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1881/1881-马鞍山当涂君悦华庭店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1881/1881-马鞍山当涂君悦华庭店-经营许可证.jpg"
    }
  },
  {
    "id": "1607",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山当涂大润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1607/1607-马鞍山当涂大润发餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1607/1607-马鞍山当涂大润发餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1954",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山恒大御景湾店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1954/1954-马鞍山恒大御景湾店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1954/1954-马鞍山恒大御景湾店-经营许可证.jpeg"
    }
  },
  {
    "id": "1478",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山欣明国际餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1478/1478-马鞍山欣明国际餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1478/1478-马鞍山欣明国际餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1962",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山绿洲花园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1962/1962-马鞍山绿洲花园店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1962/1962-马鞍山绿洲花园店-经营许可证.jpeg"
    }
  },
  {
    "id": "1568",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山郑蒲港店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1568/1568-马鞍山郑蒲港店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1568/1568-马鞍山郑蒲港店-经营许可证.jpg"
    }
  },
  {
    "id": "1477",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山金鹰购物中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1477/1477-马鞍山金鹰购物中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1477/1477-马鞍山金鹰购物中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1817",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山银河湾店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1817/1817-马鞍山银河湾店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1817/1817-马鞍山银河湾店-经营许可证.jpeg"
    }
  },
  {
    "id": "1649",
    "province": "安徽省",
    "city": "马鞍山市",
    "name": "马鞍山雨山路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1649/1649-马鞍山雨山路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1649/1649-马鞍山雨山路店-经营许可证.jpg"
    }
  },
  {
    "id": "1146",
    "province": "安徽省",
    "city": "黄山市",
    "name": "休宁萝宁新天地店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1146/1146-休宁萝宁新天地店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1146/1146-休宁萝宁新天地店-经营许可证.jpg"
    }
  },
  {
    "id": "1640",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山休宁服务区东区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1640/1640-黄山休宁服务区东区店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1640/1640-黄山休宁服务区东区店-经营许可证.jpeg"
    }
  },
  {
    "id": "1641",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山休宁服务区西区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1641/1641-黄山休宁服务区西区店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1641/1641-黄山休宁服务区西区店-经营许可证.jpeg"
    }
  },
  {
    "id": "1982",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山凫峰服务区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1982/1982-黄山凫峰服务区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1982/1982-黄山凫峰服务区店-经营许可证.jpg"
    }
  },
  {
    "id": "1544",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山呈坎服务区东区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1544/1544-黄山呈坎服务区东区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1544/1544-黄山呈坎服务区东区店-经营许可证.jpg"
    }
  },
  {
    "id": "1706",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1706/1706-黄山大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1706/1706-黄山大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1824",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山大观店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1824/1824-黄山大观店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1824/1824-黄山大观店-经营许可证.jpeg"
    }
  },
  {
    "id": "1688",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山太平洋购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1688/1688-黄山太平洋购物中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1688/1688-黄山太平洋购物中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1255",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山宏村店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1255/1255-黄山宏村店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1255/1255-黄山宏村店-经营许可证.jpg"
    }
  },
  {
    "id": "1687",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山岩寺老街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1687/1687-黄山岩寺老街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1687/1687-黄山岩寺老街店-经营许可证.jpeg"
    }
  },
  {
    "id": "1421",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山歙县紫阳广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1421/1421-黄山歙县紫阳广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1421/1421-黄山歙县紫阳广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1898",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山汤口镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1898/1898-黄山汤口镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1898/1898-黄山汤口镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1890",
    "province": "安徽省",
    "city": "黄山市",
    "name": "黄山浩创城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1890/1890-黄山浩创城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1890/1890-黄山浩创城店-经营许可证.jpg"
    }
  },
  {
    "id": "1937",
    "province": "广东省",
    "city": "中山市",
    "name": "中山石岐万象汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1937/1937-中山石岐万象汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1937/1937-中山石岐万象汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1721",
    "province": "广东省",
    "city": "佛山市",
    "name": "佛山万科金融中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1721/1721-佛山万科金融中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1721/1721-佛山万科金融中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1989",
    "province": "广东省",
    "city": "佛山市",
    "name": "佛山千灯湖环宇城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1989/1989-佛山千灯湖环宇城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1989/1989-佛山千灯湖环宇城店-经营许可证.jpg"
    }
  },
  {
    "id": "1946",
    "province": "广东省",
    "city": "佛山市",
    "name": "佛山南海万科广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1946/1946-佛山南海万科广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1946/1946-佛山南海万科广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1995",
    "province": "广东省",
    "city": "佛山市",
    "name": "佛山悦然广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1995/1995-佛山悦然广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1995/1995-佛山悦然广场店-经营许可证.png"
    }
  },
  {
    "id": "1839",
    "province": "广东省",
    "city": "佛山市",
    "name": "佛山金铂中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1839/1839-佛山金铂中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1839/1839-佛山金铂中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1711",
    "province": "广东省",
    "city": "佛山市",
    "name": "佛山顺德万象汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1711/1711-佛山顺德万象汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1711/1711-佛山顺德万象汇店-经营许可证.jpg"
    }
  },
  {
    "id": "A016",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳HALO PLACE店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a72e40aa-67ae-4790-b96e-901f597a1ab21662621879817.png",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/25fb4b9a-2b0a-42f3-acef-f8b396dcda1d1662621880151.JPG"
    }
  },
  {
    "id": "A007",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳TCL大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f8fd44de-d6ef-4a64-a843-ccb8ae25221b1778951278586.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dde9d52d-f8b2-4849-8b29-32caa0ecc6f61778951278883.jpg"
    }
  },
  {
    "id": "A043",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳佳华领悦店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/IMG_20260108_1055551768213303903.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/IMG_20260108_1055171768213311860.jpg"
    }
  },
  {
    "id": "A019",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳先健科技大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/42356c1c-1664-41e6-9b6e-7fb60c4f13d21668013665555.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e0b46b8d-26bb-4b99-aae5-e7fd7fe179fa1668013665945.pdf"
    }
  },
  {
    "id": "A025",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳共乐智慧城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/31941093-0ad6-4e5b-99f2-75fd78168ab91696871208655.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/be495cd4-de01-4363-b38b-39ba9caafc8c1696871219484.pdf"
    }
  },
  {
    "id": "A008",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳前海卓越店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6924661a-acf4-4662-905a-97f2a58aa1aa1662621870833.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dd34a0af-ff90-4497-98a2-566d17e2dd8c1662621871207.jpg"
    }
  },
  {
    "id": "A014",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳后海卓悦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/621f378b-79c8-4a7c-88b6-8d4839a76e9c1662621877065.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a80fd1e2-4557-488b-90f7-e57a24e0227f1662621877429.jpg"
    }
  },
  {
    "id": "A013",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳壹方天地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f3311902-391f-45fc-908a-0e60c8f9b8d61662621875985.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bb25b34c-7410-4535-94cb-56c8c2e9ae131662621876318.jpg"
    }
  },
  {
    "id": "A011",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳平安金融中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/595231df-0fa5-487e-9d5f-72fde649a1121662621874718.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cb70df94-d7cc-447f-973c-2db48f9c9dbc1662621875090.jpg"
    }
  },
  {
    "id": "A004",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳方大城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b3864844-3070-4b41-8567-033a09756f9f1780160880362.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8df858df-dacf-463e-bb19-8064b5514eba1780160880760.jpg"
    }
  },
  {
    "id": "A022",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳星河WORLD店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/33d91b53-42b1-44d7-95d4-18a4382a50da1675443765630.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8501958e-5e4c-4ddc-91a9-c2c541bc33ee1675443765943.jpg"
    }
  },
  {
    "id": "A015",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳汉京中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8f3a51ec-00ac-4411-8250-0bb4895ec6801662621878402.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2de004d2-638a-46b3-8e1a-9258bc905dc41662621878826.pdf"
    }
  },
  {
    "id": "A017",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳深业上城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bd6f62d7-f77b-4110-80e4-333a0216e8271668013664367.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3ce366d0-b0c5-422d-b565-7f457a92cc2d1668013664720.jpg"
    }
  },
  {
    "id": "A005",
    "province": "广东省",
    "city": "深圳市",
    "name": "深圳龙岗万达广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bbe522da-cc4a-40b1-a6ce-57136773348c1662621867139.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c26bdede-108c-4cba-95bd-5a84182eb8e51662621867504.jpg"
    }
  },
  {
    "id": "1722",
    "province": "广东省",
    "city": "珠海市",
    "name": "珠海万象汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1722/1722-珠海万象汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1722/1722-珠海万象汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1242",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京万寿购物中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1242/1242-南京万寿购物中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1242/1242-南京万寿购物中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1913",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京万科都荟天地城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1913/1913-南京万科都荟天地城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1913/1913-南京万科都荟天地城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1809",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京万谷慧生活广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1809/1809-南京万谷慧生活广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1809/1809-南京万谷慧生活广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1783",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京万达永辉超市店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1783/1783-南京万达永辉超市店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1783/1783-南京万达永辉超市店-经营许可证.jpg"
    }
  },
  {
    "id": "1405",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京万达茂店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1405/1405-南京万达茂店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1405/1405-南京万达茂店-经营许可证.jpg"
    }
  },
  {
    "id": "1319",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京世纪雅苑店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1319/1319-南京世纪雅苑店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1319/1319-南京世纪雅苑店-经营许可证.jpeg"
    }
  },
  {
    "id": "1652",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京东善桥店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1652/1652-南京东善桥店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1652/1652-南京东善桥店-经营许可证.jpg"
    }
  },
  {
    "id": "1370",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京东来荟邻生活广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1370/1370-南京东来荟邻生活广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1370/1370-南京东来荟邻生活广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1106",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京中南棉花糖餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1106/1106-南京中南棉花糖餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1106/1106-南京中南棉花糖餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1190",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京中山东路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1190/1190-南京中山东路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1190/1190-南京中山东路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1758",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京中海龙湾财富中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1758/1758-南京中海龙湾财富中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1758/1758-南京中海龙湾财富中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1406",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京中粮鸿云坊店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1406/1406-南京中粮鸿云坊店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1406/1406-南京中粮鸿云坊店-经营许可证.jpg"
    }
  },
  {
    "id": "1326",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京仁恒置地广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1326/1326-南京仁恒置地广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1326/1326-南京仁恒置地广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1546",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京仙林金鹰店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1546/1546-南京仙林金鹰店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1546/1546-南京仙林金鹰店-经营许可证.jpg"
    }
  },
  {
    "id": "1552",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京伟星万科店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1552/1552-南京伟星万科店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1552/1552-南京伟星万科店-经营许可证.jpeg"
    }
  },
  {
    "id": "1635",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京传媒学院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1635/1635-南京传媒学院店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1635/1635-南京传媒学院店-经营许可证.jpeg"
    }
  },
  {
    "id": "1196",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京俊杰大厦餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1196/1196-南京俊杰大厦餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1196/1196-南京俊杰大厦餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1216",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京保利樾广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1216/1216-南京保利樾广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1216/1216-南京保利樾广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1243",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京六合服务区东餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1243/1243-南京六合服务区东餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1243/1243-南京六合服务区东餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1244",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京六合服务区西餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1244/1244-南京六合服务区西餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1244/1244-南京六合服务区西餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1894",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京共享大厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1894/1894-南京共享大厦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1894/1894-南京共享大厦店-经营许可证.jpg"
    }
  },
  {
    "id": "1212",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京凤展路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1212/1212-南京凤展路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1212/1212-南京凤展路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1240",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京升龙汇金餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1240/1240-南京升龙汇金餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1240/1240-南京升龙汇金餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1359",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京华贸中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1359/1359-南京华贸中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1359/1359-南京华贸中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1746",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京南理工科技园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1746/1746-南京南理工科技园店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1746/1746-南京南理工科技园店-经营许可证.jpg"
    }
  },
  {
    "id": "1233",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京吴侯街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1233/1233-南京吴侯街餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1233/1233-南京吴侯街餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1184",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京喜玛拉雅餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1184/1184-南京喜玛拉雅餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1184/1184-南京喜玛拉雅餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1395",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京城南茂店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1395/1395-南京城南茂店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1395/1395-南京城南茂店-经营许可证.jpg"
    }
  },
  {
    "id": "1265",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京大田悦生活店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1265/1265-南京大田悦生活店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1265/1265-南京大田悦生活店-经营许可证.jpg"
    }
  },
  {
    "id": "1109",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京天元路店餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1109/1109-南京天元路店餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1109/1109-南京天元路店餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1238",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京天印大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1238/1238-南京天印大道餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1238/1238-南京天印大道餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1727",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京夫子庙地铁站店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1727/1727-南京夫子庙地铁站店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1727/1727-南京夫子庙地铁站店-经营许可证.jpg"
    }
  },
  {
    "id": "1380",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京尧佳路苏果店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1380/1380-南京尧佳路苏果店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1380/1380-南京尧佳路苏果店-经营许可证.jpg"
    }
  },
  {
    "id": "1209",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京市三山街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1209/1209-南京市三山街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1209/1209-南京市三山街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1189",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京市保利餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1189/1189-南京市保利餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1189/1189-南京市保利餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1103",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京市华侨城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1103/1103-南京市华侨城餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1103/1103-南京市华侨城餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1211",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京市殷华街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1211/1211-南京市殷华街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1211/1211-南京市殷华街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1213",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京市观竹苑餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1213/1213-南京市观竹苑餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1213/1213-南京市观竹苑餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1108",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京市郁金香路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1108/1108-南京市郁金香路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1108/1108-南京市郁金香路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1967",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京常发广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1967/1967-南京常发广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1967/1967-南京常发广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1317",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京幸福里商业街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1317/1317-南京幸福里商业街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1317/1317-南京幸福里商业街店-经营许可证.jpg"
    }
  },
  {
    "id": "1648",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京建邺万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1648/1648-南京建邺万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1648/1648-南京建邺万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1659",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京建邺区海峡城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1659/1659-南京建邺区海峡城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1659/1659-南京建邺区海峡城店-经营许可证.jpg"
    }
  },
  {
    "id": "1100",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京弘阳大道餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1100/1100-南京弘阳大道餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1100/1100-南京弘阳大道餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1169",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京弘阳广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1169/1169-南京弘阳广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1169/1169-南京弘阳广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1425",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京招商花园城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1425/1425-南京招商花园城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1425/1425-南京招商花园城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1178",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京文博路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1178/1178-南京文博路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1178/1178-南京文博路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1905",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京新世纪广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1905/1905-南京新世纪广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1905/1905-南京新世纪广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1114",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京新城总部餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1114/1114-南京新城总部餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1114/1114-南京新城总部餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1228",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京新城汇餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1228/1228-南京新城汇餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1228/1228-南京新城汇餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1231",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京新城科技园餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1231/1231-南京新城科技园餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1231/1231-南京新城科技园餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1181",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京新尧金地广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1181/1181-南京新尧金地广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1181/1181-南京新尧金地广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1660",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京新街口大洋百货店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1660/1660-南京新街口大洋百货店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1660/1660-南京新街口大洋百货店-经营许可证.jpg"
    }
  },
  {
    "id": "1191",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京新街口金鹰餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1191/1191-南京新街口金鹰餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1191/1191-南京新街口金鹰餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1186",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京明发国际中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1186/1186-南京明发国际中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1186/1186-南京明发国际中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1224",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京欢乐港餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1224/1224-南京欢乐港餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1224/1224-南京欢乐港餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1220",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京正丰街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1220/1220-南京正丰街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1220/1220-南京正丰街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1201",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京汉中门地铁口餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1201/1201-南京汉中门地铁口餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1201/1201-南京汉中门地铁口餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1226",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京江东中路三餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1226/1226-南京江东中路三餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1226/1226-南京江东中路三餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1115",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京江浦市民中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1115/1115-南京江浦市民中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1115/1115-南京江浦市民中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1180",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京浦口侨康路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1180/1180-南京浦口侨康路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1180/1180-南京浦口侨康路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1994",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京浦口江月府店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1994/1994-南京浦口江月府店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1994/1994-南京浦口江月府店-经营许可证.jpg"
    }
  },
  {
    "id": "1901",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京淳化德购超市店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1901/1901-南京淳化德购超市店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1901/1901-南京淳化德购超市店-经营许可证.jpeg"
    }
  },
  {
    "id": "1316",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京清江苏宁广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1316/1316-南京清江苏宁广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1316/1316-南京清江苏宁广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1234",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京湖东路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1234/1234-南京湖东路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1234/1234-南京湖东路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1236",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京溧水乐活城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1236/1236-南京溧水乐活城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1236/1236-南京溧水乐活城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1101",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京溧水珍珠路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1101/1101-南京溧水珍珠路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1101/1101-南京溧水珍珠路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1720",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京澳林广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1720/1720-南京澳林广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1720/1720-南京澳林广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1960",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京熙乐汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1960/1960-南京熙乐汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1960/1960-南京熙乐汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1247",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京燕亭路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1247/1247-南京燕亭路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1247/1247-南京燕亭路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1225",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京珠江路金鹰餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1225/1225-南京珠江路金鹰餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1225/1225-南京珠江路金鹰餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1098",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京甬利广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1098/1098-南京甬利广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1098/1098-南京甬利广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1195",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京百利广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1195/1195-南京百利广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1195/1195-南京百利广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1335",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京百家湖金鹰店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1335/1335-南京百家湖金鹰店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1335/1335-南京百家湖金鹰店-经营许可证.jpg"
    }
  },
  {
    "id": "1917",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京盘金华府店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1917/1917-南京盘金华府店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1917/1917-南京盘金华府店-经营许可证.jpg"
    }
  },
  {
    "id": "1327",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京砂之船店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1327/1327-南京砂之船店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1327/1327-南京砂之船店-经营许可证.png"
    }
  },
  {
    "id": "1215",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京紫东创意园餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1215/1215-南京紫东创意园餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1215/1215-南京紫东创意园餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1182",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京紫峰大厦餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1182/1182-南京紫峰大厦餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1182/1182-南京紫峰大厦餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1208",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京绿地之窗二餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1208/1208-南京绿地之窗二餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1208/1208-南京绿地之窗二餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1120",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京绿地花茂餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1120/1120-南京绿地花茂餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1120/1120-南京绿地花茂餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1204",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京胜利路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1204/1204-南京胜利路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1204/1204-南京胜利路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1207",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京能仁里餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1207/1207-南京能仁里餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1207/1207-南京能仁里餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1875",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京花园路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1875/1875-南京花园路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1875/1875-南京花园路店-经营许可证.jpg"
    }
  },
  {
    "id": "1199",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京花生唐购物广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1199/1199-南京花生唐购物广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1199/1199-南京花生唐购物广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1321",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京苏宁睿城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1321/1321-南京苏宁睿城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1321/1321-南京苏宁睿城店-经营许可证.jpg"
    }
  },
  {
    "id": "1179",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京苏宁青创园餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1179/1179-南京苏宁青创园餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1179/1179-南京苏宁青创园餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1096",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京莲池路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1096/1096-南京莲池路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1096/1096-南京莲池路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1246",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京虹桥中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1246/1246-南京虹桥中心餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1246/1246-南京虹桥中心餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1980",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京谷里U购U生活超市店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1980/1980-南京谷里U购U生活超市店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1980/1980-南京谷里U购U生活超市店-经营许可证.jpg"
    }
  },
  {
    "id": "1113",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京财富中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1113/1113-南京财富中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1113/1113-南京财富中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1119",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京财智广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1119/1119-南京财智广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1119/1119-南京财智广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1740",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京软件谷云密城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1740/1740-南京软件谷云密城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1740/1740-南京软件谷云密城店-经营许可证.jpg"
    }
  },
  {
    "id": "1217",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京远洋国际中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1217/1217-南京远洋国际中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1217/1217-南京远洋国际中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1121",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京通淮街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1121/1121-南京通淮街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1121/1121-南京通淮街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1227",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京金奥国际中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1227/1227-南京金奥国际中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1227/1227-南京金奥国际中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1230",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京金威广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1230/1230-南京金威广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1230/1230-南京金威广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1102",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京金盛路二店.",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1102/1102-南京金盛路二店.-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1102/1102-南京金盛路二店.-经营许可证.jpg"
    }
  },
  {
    "id": "1239",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京金融城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1239/1239-南京金融城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1239/1239-南京金融城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1318",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京金象城购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1318/1318-南京金象城购物中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1318/1318-南京金象城购物中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1117",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京金马路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1117/1117-南京金马路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1117/1117-南京金马路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1874",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京鑫乐生活广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1874/1874-南京鑫乐生活广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1874/1874-南京鑫乐生活广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1863",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京钟鼎名悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1863/1863-南京钟鼎名悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1863/1863-南京钟鼎名悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1734",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京铂丽大厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1734/1734-南京铂丽大厦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1734/1734-南京铂丽大厦店-经营许可证.jpg"
    }
  },
  {
    "id": "1237",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京银城东苑餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1237/1237-南京银城东苑餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1237/1237-南京银城东苑餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1185",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京银城中心餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1185/1185-南京银城中心餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1185/1185-南京银城中心餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1202",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京锁金村餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1202/1202-南京锁金村餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1202/1202-南京锁金村餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1781",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京陈沛桥商业邻里中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1781/1781-南京陈沛桥商业邻里中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1781/1781-南京陈沛桥商业邻里中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1843",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京雨山龙湖天街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1843/1843-南京雨山龙湖天街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1843/1843-南京雨山龙湖天街店-经营许可证.jpg"
    }
  },
  {
    "id": "1095",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京雨润大街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1095/1095-南京雨润大街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1095/1095-南京雨润大街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1210",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京雨花世茂餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1210/1210-南京雨花世茂餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1210/1210-南京雨花世茂餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1219",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京顾家营路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1219/1219-南京顾家营路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1219/1219-南京顾家营路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1118",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京高淳人民医院餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1118/1118-南京高淳人民医院餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1118/1118-南京高淳人民医院餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1985",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京高淳医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1985/1985-南京高淳医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1985/1985-南京高淳医院店-经营许可证.png"
    }
  },
  {
    "id": "1738",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京高淳宝龙店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1738/1738-南京高淳宝龙店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1738/1738-南京高淳宝龙店-经营许可证.jpg"
    }
  },
  {
    "id": "1269",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京高科荣境店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1269/1269-南京高科荣境店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1269/1269-南京高科荣境店-经营许可证.jpg"
    }
  },
  {
    "id": "1235",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京高铁网谷餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1235/1235-南京高铁网谷餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1235/1235-南京高铁网谷餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1175",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京鸿利广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1175/1175-南京鸿利广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1175/1175-南京鸿利广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1192",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京麒麟东路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1192/1192-南京麒麟东路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1192/1192-南京麒麟东路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1322",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京麒麟街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1322/1322-南京麒麟街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1322/1322-南京麒麟街店-经营许可证.jpg"
    }
  },
  {
    "id": "1173",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京龙池餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1173/1173-南京龙池餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1173/1173-南京龙池餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1116",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京龙湖文景路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1116/1116-南京龙湖文景路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1116/1116-南京龙湖文景路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1857",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京龙湖河西天街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1857/1857-南京龙湖河西天街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1857/1857-南京龙湖河西天街店-经营许可证.jpeg"
    }
  },
  {
    "id": "1218",
    "province": "江苏省",
    "city": "南京市",
    "name": "南京龙蟠汇餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1218/1218-南京龙蟠汇餐厅-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1218/1218-南京龙蟠汇餐厅-经营许可证.png"
    }
  },
  {
    "id": "1804",
    "province": "江苏省",
    "city": "南京市",
    "name": "高淳八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1804/1804-高淳八佰伴店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1804/1804-高淳八佰伴店-经营许可证.jpg"
    }
  },
  {
    "id": "1354",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通七彩城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1354/1354-南通七彩城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1354/1354-南通七彩城店-经营许可证.jpg"
    }
  },
  {
    "id": "1812",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通万佑广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1812/1812-南通万佑广场店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1812/1812-南通万佑广场店-经营许可证.png"
    }
  },
  {
    "id": "1399",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通万象城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1399/1399-南通万象城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1399/1399-南通万象城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1926",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通世茂广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1926/1926-南通世茂广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1926/1926-南通世茂广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1417",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通中南城彩虹漾店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1417/1417-南通中南城彩虹漾店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1417/1417-南通中南城彩虹漾店-经营许可证.jpeg"
    }
  },
  {
    "id": "1654",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通利群时代龙王桥店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1654/1654-南通利群时代龙王桥店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1654/1654-南通利群时代龙王桥店-经营许可证.jpg"
    }
  },
  {
    "id": "1718",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通南大街步行街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1718/1718-南通南大街步行街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1718/1718-南通南大街步行街店-经营许可证.jpg"
    }
  },
  {
    "id": "1374",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通印象城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1374/1374-南通印象城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1374/1374-南通印象城店-经营许可证.jpg"
    }
  },
  {
    "id": "1627",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通叠石桥罗缦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1627/1627-南通叠石桥罗缦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1627/1627-南通叠石桥罗缦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1762",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通启东文峰大世界店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1762/1762-南通启东文峰大世界店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1762/1762-南通启东文峰大世界店-经营许可证.jpeg"
    }
  },
  {
    "id": "1748",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通启东申港城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1748/1748-南通启东申港城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1748/1748-南通启东申港城店-经营许可证.jpg"
    }
  },
  {
    "id": "1545",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通圆融嘉悦汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1545/1545-南通圆融嘉悦汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1545/1545-南通圆融嘉悦汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1408",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通如东大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1408/1408-南通如东大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1408/1408-南通如东大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1221",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通如皋万达餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1221/1221-南通如皋万达餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1221/1221-南通如皋万达餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1576",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通如皋文峰大世界店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1576/1576-南通如皋文峰大世界店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1576/1576-南通如皋文峰大世界店-经营许可证.jpeg"
    }
  },
  {
    "id": "1267",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通崇川大有境店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1267/1267-南通崇川大有境店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1267/1267-南通崇川大有境店-经营许可证.jpg"
    }
  },
  {
    "id": "1936",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通文峰广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1936/1936-南通文峰广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1936/1936-南通文峰广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1355",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通海安万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1355/1355-南通海安万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1355/1355-南通海安万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1927",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通海安大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1927/1927-南通海安大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1927/1927-南通海安大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1162",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通海安盛世名门餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1162/1162-南通海安盛世名门餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1162/1162-南通海安盛世名门餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1861",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通海门利群时代店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1861/1861-南通海门利群时代店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1861/1861-南通海门利群时代店-经营许可证.jpg"
    }
  },
  {
    "id": "1301",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通海门龙信广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1301/1301-南通海门龙信广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1301/1301-南通海门龙信广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1130",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通深南路大润发餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1130/1130-南通深南路大润发餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1130/1130-南通深南路大润发餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1360",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通财富广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1360/1360-南通财富广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1360/1360-南通财富广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1381",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通通州万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1381/1381-南通通州万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1381/1381-南通通州万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1866",
    "province": "江苏省",
    "city": "南通市",
    "name": "南通金鹰餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1866/1866-南通金鹰餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1866/1866-南通金鹰餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1333",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1333/1333-宿迁人民医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1333/1333-宿迁人民医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1164",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1164/1164-宿迁宝龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1164/1164-宿迁宝龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1261",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁宿城吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1261/1261-宿迁宿城吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1261/1261-宿迁宿城吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1790",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁市泗阳哥伦布店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1790/1790-宿迁市泗阳哥伦布店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1790/1790-宿迁市泗阳哥伦布店-经营许可证.jpeg"
    }
  },
  {
    "id": "1270",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁沭阳中央广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1270/1270-宿迁沭阳中央广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1270/1270-宿迁沭阳中央广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1325",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁沭阳浙江商城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1325/1325-宿迁沭阳浙江商城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1325/1325-宿迁沭阳浙江商城店-经营许可证.jpg"
    }
  },
  {
    "id": "1529",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁泗洪泗州吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1529/1529-宿迁泗洪泗州吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1529/1529-宿迁泗洪泗州吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1159",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁泗洪花园口店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1159/1159-宿迁泗洪花园口店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1159/1159-宿迁泗洪花园口店-经营许可证.jpeg"
    }
  },
  {
    "id": "1558",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁泗阳吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1558/1558-宿迁泗阳吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1558/1558-宿迁泗阳吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1313",
    "province": "江苏省",
    "city": "宿迁市",
    "name": "宿迁泗阳大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1313/1313-宿迁泗阳大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1313/1313-宿迁泗阳大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1358",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州T-PARK蓝海豚店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1358/1358-常州T-PARK蓝海豚店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1358/1358-常州T-PARK蓝海豚店-经营许可证.jpg"
    }
  },
  {
    "id": "1300",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州万和城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1300/1300-常州万和城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1300/1300-常州万和城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1868",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州世茂广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1868/1868-常州世茂广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1868/1868-常州世茂广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1141",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州九洲新世界店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1141/1141-常州九洲新世界店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1141/1141-常州九洲新世界店-经营许可证.jpg"
    }
  },
  {
    "id": "1150",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州关河大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1150/1150-常州关河大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1150/1150-常州关河大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1298",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州典雅广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1298/1298-常州典雅广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1298/1298-常州典雅广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1348",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州天宁吾悦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1348/1348-常州天宁吾悦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1348/1348-常州天宁吾悦店-经营许可证.jpg"
    }
  },
  {
    "id": "1382",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州好利广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1382/1382-常州好利广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1382/1382-常州好利广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1879",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州弘阳广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1879/1879-常州弘阳广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1879/1879-常州弘阳广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1A09",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州怀德大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A09/1A09-常州怀德大润发店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A09/1A09-常州怀德大润发店-经营许可证.png"
    }
  },
  {
    "id": "1775",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州星湖荟店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1775/1775-常州星湖荟店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1775/1775-常州星湖荟店-经营许可证.jpeg"
    }
  },
  {
    "id": "1966",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州星耀吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1966/1966-常州星耀吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1966/1966-常州星耀吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1356",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州曼哈顿广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1356/1356-常州曼哈顿广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1356/1356-常州曼哈顿广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1865",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州歌林公园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1865/1865-常州歌林公园店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1865/1865-常州歌林公园店-经营许可证.jpeg"
    }
  },
  {
    "id": "1264",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州武进万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1264/1264-常州武进万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1264/1264-常州武进万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1148",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州武进大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1148/1148-常州武进大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1148/1148-常州武进大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1601",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州湖塘万博店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1601/1601-常州湖塘万博店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1601/1601-常州湖塘万博店-经营许可证.jpg"
    }
  },
  {
    "id": "1160",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州溧阳大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1160/1160-常州溧阳大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1160/1160-常州溧阳大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1153",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州火车站店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1153/1153-常州火车站店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1153/1153-常州火车站店-经营许可证.jpg"
    }
  },
  {
    "id": "1384",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州环球港店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1384/1384-常州环球港店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1384/1384-常州环球港店-经营许可证.jpeg"
    }
  },
  {
    "id": "1342",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州科教城北店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1342/1342-常州科教城北店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1342/1342-常州科教城北店-经营许可证.jpg"
    }
  },
  {
    "id": "1259",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1259/1259-常州购物中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1259/1259-常州购物中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1556",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州遥观大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1556/1556-常州遥观大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1556/1556-常州遥观大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1268",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州邹区泰富时代广场",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1268/1268-常州邹区泰富时代广场-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1268/1268-常州邹区泰富时代广场-经营许可证.jpg"
    }
  },
  {
    "id": "1123",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州金坛万科店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1123/1123-常州金坛万科店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1123/1123-常州金坛万科店-经营许可证.jpg"
    }
  },
  {
    "id": "1657",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州金坛八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1657/1657-常州金坛八佰伴店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1657/1657-常州金坛八佰伴店-经营许可证.jpg"
    }
  },
  {
    "id": "1377",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州金坛吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1377/1377-常州金坛吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1377/1377-常州金坛吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1349",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州金坛新天地店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1349/1349-常州金坛新天地店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1349/1349-常州金坛新天地店-经营许可证.jpg"
    }
  },
  {
    "id": "1149",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州马杭店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1149/1149-常州马杭店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1149/1149-常州马杭店-经营许可证.jpg"
    }
  },
  {
    "id": "1891",
    "province": "江苏省",
    "city": "常州市",
    "name": "常州高力国际汽博城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1891/1891-常州高力国际汽博城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1891/1891-常州高力国际汽博城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1012",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州万福街餐厅（儿童医院）",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1012/1012-徐州万福街餐厅（儿童医院）-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1012/1012-徐州万福街餐厅（儿童医院）-经营许可证.jpg"
    }
  },
  {
    "id": "1981",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州万科新淮印象城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1981/1981-徐州万科新淮印象城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1981/1981-徐州万科新淮印象城店-经营许可证.jpg"
    }
  },
  {
    "id": "1A05",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州东关大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A05/1A05-徐州东关大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A05/1A05-徐州东关大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1560",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州云龙万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1560/1560-徐州云龙万达店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1560/1560-徐州云龙万达店-经营许可证.jpeg"
    }
  },
  {
    "id": "1010",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州和信宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1010/1010-徐州和信宝龙广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1010/1010-徐州和信宝龙广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1869",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州市中医院新院区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1869/1869-徐州市中医院新院区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1869/1869-徐州市中医院新院区店-经营许可证.jpg"
    }
  },
  {
    "id": "1661",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州市建国路大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1661/1661-徐州市建国路大润发店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1661/1661-徐州市建国路大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1646",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州市铜山万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1646/1646-徐州市铜山万达广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1646/1646-徐州市铜山万达广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1852",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州彭城苏宁广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1852/1852-徐州彭城苏宁广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1852/1852-徐州彭城苏宁广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1747",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州招商花园城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1747/1747-徐州招商花园城店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1747/1747-徐州招商花园城店-经营许可证.png"
    }
  },
  {
    "id": "1105",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州新沂金桥国际店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1105/1105-徐州新沂金桥国际店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1105/1105-徐州新沂金桥国际店-经营许可证.jpg"
    }
  },
  {
    "id": "1037",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州泰隆商业街餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1037/1037-徐州泰隆商业街餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1037/1037-徐州泰隆商业街餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1911",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州淮海环球港店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1911/1911-徐州淮海环球港店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1911/1911-徐州淮海环球港店-经营许可证.jpg"
    }
  },
  {
    "id": "1655",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州第一人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1655/1655-徐州第一人民医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1655/1655-徐州第一人民医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1035",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州美的广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1035/1035-徐州美的广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1035/1035-徐州美的广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1620",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州苏宁广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1620/1620-徐州苏宁广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1620/1620-徐州苏宁广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1045",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州邳州大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1045/1045-徐州邳州大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1045/1045-徐州邳州大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1013",
    "province": "江苏省",
    "city": "徐州市",
    "name": "徐州颐和汇邻湾餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1013/1013-徐州颐和汇邻湾餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1013/1013-徐州颐和汇邻湾餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1838",
    "province": "江苏省",
    "city": "徐州市",
    "name": "邳州通城欢乐买店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1838/1838-邳州通城欢乐买店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1838/1838-邳州通城欢乐买店-经营许可证.jpg"
    }
  },
  {
    "id": "1038",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州三盛国际餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1038/1038-扬州三盛国际餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1038/1038-扬州三盛国际餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1042",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州仪征宝能餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1042/1042-扬州仪征宝能餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1042/1042-扬州仪征宝能餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1072",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州国庆路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1072/1072-扬州国庆路餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1072/1072-扬州国庆路餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1036",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州广陵新城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1036/1036-扬州广陵新城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1036/1036-扬州广陵新城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1063",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州昌建广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1063/1063-扬州昌建广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1063/1063-扬州昌建广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1074",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州江都东方红路餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1074/1074-扬州江都东方红路餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1074/1074-扬州江都东方红路餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1089",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州江都大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1089/1089-扬州江都大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1089/1089-扬州江都大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1723",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州邗江吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1723/1723-扬州邗江吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1723/1723-扬州邗江吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1308",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州邗江宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1308/1308-扬州邗江宝龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1308/1308-扬州邗江宝龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1078",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州顺达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1078/1078-扬州顺达广场店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1078/1078-扬州顺达广场店-经营许可证.png"
    }
  },
  {
    "id": "1073",
    "province": "江苏省",
    "city": "扬州市",
    "name": "扬州高邮世贸广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1073/1073-扬州高邮世贸广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1073/1073-扬州高邮世贸广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1080",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡11广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1080/1080-无锡11广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1080/1080-无锡11广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1085",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡Kpark商务中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1085/1085-无锡Kpark商务中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1085/1085-无锡Kpark商务中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1058",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡中山路红豆店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1058/1058-无锡中山路红豆店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1058/1058-无锡中山路红豆店-经营许可证.jpeg"
    }
  },
  {
    "id": "1805",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡九里东韵店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1805/1805-无锡九里东韵店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1805/1805-无锡九里东韵店-经营许可证.jpg"
    }
  },
  {
    "id": "1997",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡五洲国际工业博览城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1997/1997-无锡五洲国际工业博览城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1997/1997-无锡五洲国际工业博览城店-经营许可证.jpg"
    }
  },
  {
    "id": "1629",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡佛奥天佑城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1629/1629-无锡佛奥天佑城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1629/1629-无锡佛奥天佑城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1847",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡和畅睦邻广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1847/1847-无锡和畅睦邻广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1847/1847-无锡和畅睦邻广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1928",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡国际招商城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1928/1928-无锡国际招商城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1928/1928-无锡国际招商城店-经营许可证.jpg"
    }
  },
  {
    "id": "1719",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡圆融嘉悦汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1719/1719-无锡圆融嘉悦汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1719/1719-无锡圆融嘉悦汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1559",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡圆融广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1559/1559-无锡圆融广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1559/1559-无锡圆融广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1094",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡太平洋城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1094/1094-无锡太平洋城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1094/1094-无锡太平洋城店-经营许可证.jpg"
    }
  },
  {
    "id": "1292",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡太湖智谷店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1292/1292-无锡太湖智谷店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1292/1292-无锡太湖智谷店-经营许可证.jpeg"
    }
  },
  {
    "id": "1603",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡宜兴东郊店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1603/1603-无锡宜兴东郊店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1603/1603-无锡宜兴东郊店-经营许可证.jpg"
    }
  },
  {
    "id": "1296",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡宜兴八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1296/1296-无锡宜兴八佰伴店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1296/1296-无锡宜兴八佰伴店-经营许可证.jpeg"
    }
  },
  {
    "id": "1554",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡宜兴和信广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1554/1554-无锡宜兴和信广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1554/1554-无锡宜兴和信广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1549",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡宜兴大统华店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1549/1549-无锡宜兴大统华店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1549/1549-无锡宜兴大统华店-经营许可证.jpg"
    }
  },
  {
    "id": "1251",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡宝龙城市广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1251/1251-无锡宝龙城市广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1251/1251-无锡宝龙城市广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1373",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡市民中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1373/1373-无锡市民中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1373/1373-无锡市民中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1931",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡市阳光广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1931/1931-无锡市阳光广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1931/1931-无锡市阳光广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1634",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡惠山万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1634/1634-无锡惠山万达店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1634/1634-无锡惠山万达店-经营许可证.jpeg"
    }
  },
  {
    "id": "1343",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡新之城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1343/1343-无锡新之城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1343/1343-无锡新之城店-经营许可证.jpg"
    }
  },
  {
    "id": "1787",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡方圆荟购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1787/1787-无锡方圆荟购物中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1787/1787-无锡方圆荟购物中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1372",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡梦享城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1372/1372-无锡梦享城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1372/1372-无锡梦享城店-经营许可证.jpg"
    }
  },
  {
    "id": "1166",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡永乐万悦集店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1166/1166-无锡永乐万悦集店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1166/1166-无锡永乐万悦集店-经营许可证.jpg"
    }
  },
  {
    "id": "1156",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡江阴万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1156/1156-无锡江阴万达广场店-营业执照.pdf",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1156/1156-无锡江阴万达广场店-经营许可证.pdf"
    }
  },
  {
    "id": "1134",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡江阴临港红豆店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1134/1134-无锡江阴临港红豆店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1134/1134-无锡江阴临港红豆店-经营许可证.jpeg"
    }
  },
  {
    "id": "1394",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡江阴八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1394/1394-无锡江阴八佰伴店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1394/1394-无锡江阴八佰伴店-经营许可证.jpg"
    }
  },
  {
    "id": "1819",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡江阴华士镇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1819/1819-无锡江阴华士镇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1819/1819-无锡江阴华士镇店-经营许可证.jpg"
    }
  },
  {
    "id": "1550",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡江阴大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1550/1550-无锡江阴大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1550/1550-无锡江阴大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1254",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡江阴小湖新村店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1254/1254-无锡江阴小湖新村店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1254/1254-无锡江阴小湖新村店-经营许可证.jpg"
    }
  },
  {
    "id": "1893",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡江阴海岸城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1893/1893-无锡江阴海岸城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1893/1893-无锡江阴海岸城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1751",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡清扬茂业店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1751/1751-无锡清扬茂业店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1751/1751-无锡清扬茂业店-经营许可证.jpeg"
    }
  },
  {
    "id": "1061",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡清扬路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1061/1061-无锡清扬路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1061/1061-无锡清扬路店-经营许可证.jpg"
    }
  },
  {
    "id": "1647",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡滨湖万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1647/1647-无锡滨湖万达店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1647/1647-无锡滨湖万达店-经营许可证.jpeg"
    }
  },
  {
    "id": "1404",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡璟隆广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1404/1404-无锡璟隆广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1404/1404-无锡璟隆广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1046",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡红豆万花城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1046/1046-无锡红豆万花城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1046/1046-无锡红豆万花城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1386",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡苏宁广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1870/1870-无锡苏宁广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1870/1870-无锡苏宁广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1315",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡金悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1315/1315-无锡金悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1315/1315-无锡金悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1555",
    "province": "江苏省",
    "city": "无锡市",
    "name": "无锡钱桥大街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1555/1555-无锡钱桥大街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1555/1555-无锡钱桥大街店-经营许可证.jpeg"
    }
  },
  {
    "id": "1726",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰兴新城吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1726/1726-泰兴新城吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1726/1726-泰兴新城吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1A06",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州万象城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A06/1A06-泰州万象城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A06/1A06-泰州万象城店-经营许可证.jpg"
    }
  },
  {
    "id": "1605",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州中骏世界城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1605/1605-泰州中骏世界城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1605/1605-泰州中骏世界城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1290",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1290/1290-泰州人民医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1290/1290-泰州人民医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1385",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州兴化吾悦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1385/1385-泰州兴化吾悦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1385/1385-泰州兴化吾悦店-经营许可证.jpg"
    }
  },
  {
    "id": "1158",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州天虹广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1158/1158-泰州天虹广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1158/1158-泰州天虹广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1415",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州市靖江吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1415/1415-泰州市靖江吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1415/1415-泰州市靖江吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1135",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州时代商业广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1135/1135-泰州时代商业广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1135/1135-泰州时代商业广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1347",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州泰兴中南悠曼里店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1347/1347-泰州泰兴中南悠曼里店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1347/1347-泰州泰兴中南悠曼里店-经营许可证.jpg"
    }
  },
  {
    "id": "1350",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州海陵万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1350/1350-泰州海陵万达店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1350/1350-泰州海陵万达店-经营许可证.jpeg"
    }
  },
  {
    "id": "1968",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州靖江荟品仓店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1968/1968-泰州靖江荟品仓店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1968/1968-泰州靖江荟品仓店-经营许可证.jpg"
    }
  },
  {
    "id": "1599",
    "province": "江苏省",
    "city": "泰州市",
    "name": "泰州龙河城投店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1599/1599-泰州龙河城投店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1599/1599-泰州龙河城投店-经营许可证.jpg"
    }
  },
  {
    "id": "1398",
    "province": "江苏省",
    "city": "淮安市",
    "name": "淮安幸福城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1398/1398-淮安幸福城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1398/1398-淮安幸福城店-经营许可证.jpg"
    }
  },
  {
    "id": "1840",
    "province": "江苏省",
    "city": "淮安市",
    "name": "淮安楚州万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1840/1840-淮安楚州万达广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1840/1840-淮安楚州万达广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1314",
    "province": "江苏省",
    "city": "淮安市",
    "name": "淮安水渡口万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1314/1314-淮安水渡口万达店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1314/1314-淮安水渡口万达店-经营许可证.jpeg"
    }
  },
  {
    "id": "1376",
    "province": "江苏省",
    "city": "淮安市",
    "name": "淮安涟水大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1376/1376-淮安涟水大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1376/1376-淮安涟水大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1334",
    "province": "江苏省",
    "city": "淮安市",
    "name": "淮安盱眙凤凰广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1334/1334-淮安盱眙凤凰广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1334/1334-淮安盱眙凤凰广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1636",
    "province": "江苏省",
    "city": "淮安市",
    "name": "淮安第一人民医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1636/1636-淮安第一人民医院店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1636/1636-淮安第一人民医院店-经营许可证.jpeg"
    }
  },
  {
    "id": "1250",
    "province": "江苏省",
    "city": "淮安市",
    "name": "淮安苏宁广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1250/1250-淮安苏宁广场餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1250/1250-淮安苏宁广场餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1413",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1413/1413-盐城万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1413/1413-盐城万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1915",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城东台市大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1915/1915-盐城东台市大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1915/1915-盐城东台市大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1528",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城中韩未来科技城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1528/1528-盐城中韩未来科技城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1528/1528-盐城中韩未来科技城店-经营许可证.jpg"
    }
  },
  {
    "id": "1604",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城大丰吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1604/1604-盐城大丰吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1604/1604-盐城大丰吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1371",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城射阳吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1371/1371-盐城射阳吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1371/1371-盐城射阳吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1A04",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城市大有境店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A04/1A04-盐城市大有境店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A04/1A04-盐城市大有境店-经营许可证.jpg"
    }
  },
  {
    "id": "1414",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城新龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1414/1414-盐城新龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1414/1414-盐城新龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1791",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城聚龙湖金鹰购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1791/1791-盐城聚龙湖金鹰购物中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1791/1791-盐城聚龙湖金鹰购物中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1311",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城金融城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1311/1311-盐城金融城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1311/1311-盐城金融城店-经营许可证.jpg"
    }
  },
  {
    "id": "1379",
    "province": "江苏省",
    "city": "盐城市",
    "name": "盐城金鹰购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1379/1379-盐城金鹰购物中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1379/1379-盐城金鹰购物中心店-经营许可证.jpg"
    }
  },
  {
    "id": "5219",
    "province": "江苏省",
    "city": "苏州市",
    "name": "吴中龙湖天街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b05624cf-7a44-4056-9f0f-aabe15da439d1675443722170.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0734f3db-4aa7-4feb-b015-bfb3c1e487ca1675443722723.jpg"
    }
  },
  {
    "id": "5163",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州久光百货店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fdd15829-d3e7-4ec7-8bc3-a24ae5ba3bd11662621530461.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e6de706f-b0fd-4ba4-abc5-f82610b237741662621530968.jpg"
    }
  },
  {
    "id": "5233",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州人工智能产业园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/82be611d-cb94-4f57-8d11-42f58b59af4d1701968739428.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/213e4110-27e7-4bd8-9495-10542616e7921701968739930.jpg"
    }
  },
  {
    "id": "5199",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州印象城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6eae166e-a608-415f-aa82-a708f44945771662621564780.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/192bfad3-247f-4a05-8a3f-eaeded3770bd1662621565117.jpg"
    }
  },
  {
    "id": "5150",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州双湖广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/34f4d31e-ad27-40bc-8779-0c44f01c2b721778346314938.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3c13b338-cd01-4489-929b-cd89fa65c32e1778346315638.jpg"
    }
  },
  {
    "id": "5205",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州吴江万宝店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/38d59991-d7ec-4078-a44b-f5a85241c86a1668618326310.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/62ec35df-9fa4-4a52-b9f2-7ef9b32af0c61668618326659.jpg"
    }
  },
  {
    "id": "5182",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州大运城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/18f98fb6-17a8-45bf-846c-09ac9af19b5a1662621547931.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aed2c85a-db6f-4f97-88e9-20830b792a701662621548258.jpg"
    }
  },
  {
    "id": "5142",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州平江万达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2ce8717d-f09f-4d67-9ff2-cff88b3ec4201662621517443.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品许可证1768827802397.jpg"
    }
  },
  {
    "id": "5231",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州张家港曼巴特店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2e0473a8-8892-4f70-97ac-39827010c64d1701968736509.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/025ae3f5-cba2-4b01-a28a-f7608492daf31701968736891.jpg"
    }
  },
  {
    "id": "5234",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州昆山中医院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eeca302c-56cb-41ce-8e88-343af8a9141a1701968740926.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/26fd17d2-0500-4e88-b342-920137bc96761701968741425.jpg"
    }
  },
  {
    "id": "5230",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州昆山金鹰店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f1c70e66-1538-4abc-9829-289083a4238f1701968734859.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c51e84cb-3b02-41b4-b8d6-6f049f1a57b11701968735456.jpg"
    }
  },
  {
    "id": "5159",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州星塘欧尚店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bc31a458-c79d-4df3-9cb9-f624e10bcfa01662621527518.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/扫描全能王 2026-5-18 09.161779067522232.pdf"
    }
  },
  {
    "id": "5191",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州星汇广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6a0f7cf4-b9ba-4777-8fc3-2f0f9981880a1662621556634.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3a9f9017-a574-48f9-94ad-ce01fd4900721662621556970.jpg"
    }
  },
  {
    "id": "5208",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州晋合广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bf2ecf18-477b-443c-a5da-647a9baa3fbb1662621573406.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c78caf92-79b4-47e8-b67e-8640617f705d1662621573743.jpg"
    }
  },
  {
    "id": "5203",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州景城邻里中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5fcdee89-a58c-49d8-acd7-481c44a4f1951662621568505.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5a72c0b7-b9a6-48d6-a990-8e88968b5aca1662621568822.jpg"
    }
  },
  {
    "id": "5190",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州汇邻广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8d3d6fd0-ccb7-432c-a565-9ce7f4d32fc91662621555439.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/27672977-7667-40fd-b14f-3d5bc933e9851662621555792.jpg"
    }
  },
  {
    "id": "5156",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州狮山路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a165011b-a113-4548-8aa8-e95201a79a751662621526356.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6551189c-199e-493a-b51b-2cff01d6b5e21662621526727.jpg"
    }
  },
  {
    "id": "5186",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州狮山金鹰店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/23b85c3a-539a-44e2-8122-67d84c6d23251662621552796.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fbaa30c3-8076-417c-8b3b-e6eafd5ef1231662621553137.jpg"
    }
  },
  {
    "id": "5195",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州现代传媒店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/990b3337-49e6-44ed-8198-5d72f15011301662621561423.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/15157f04-1818-41a2-845a-2992739030a01662621561723.jpg"
    }
  },
  {
    "id": "5206",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州盛泽欧尚店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c70ef1bc-44e5-4142-8d63-71f70b9ebc6f1662621570908.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e2cf97ac-7eb8-4a6b-b700-09d8e7ff19161662621571409.jpg"
    }
  },
  {
    "id": "5154",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州绿地复客店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bc22b070-7a67-4831-adf3-a168027e22831662621524964.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b333b7b3-cea2-4b2b-ad1d-25ef3e6628941662621525462.jpg"
    }
  },
  {
    "id": "5185",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州金鸡湖欧尚店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c3f71e08-6841-4d32-a775-7b70a76f24dc1662621551626.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/08372abc-b1f0-4ec7-8e64-aaa6497679e11662621551966.jpg"
    }
  },
  {
    "id": "5210",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州锦峰广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d836189a-b68d-47ad-8a1c-154614de88b31662621576097.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/00bafa67-3b59-4000-a608-f49d16ff3feb1662621576433.jpg"
    }
  },
  {
    "id": "5217",
    "province": "江苏省",
    "city": "苏州市",
    "name": "苏州鲜橙广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f6589f19-8950-452c-88ad-dd702c01c60b1689267708294.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1cf6a33c-d2f9-401e-8b74-5e15c76cd0681689267708621.jpg"
    }
  },
  {
    "id": "1930",
    "province": "江苏省",
    "city": "连云港市",
    "name": "连云港中山西路大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1930/1930-连云港中山西路大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1930/1930-连云港中山西路大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1430",
    "province": "江苏省",
    "city": "连云港市",
    "name": "连云港利群京东电器店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1430/1430-连云港利群京东电器店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1430/1430-连云港利群京东电器店-经营许可证.jpg"
    }
  },
  {
    "id": "1363",
    "province": "江苏省",
    "city": "连云港市",
    "name": "连云港嘉瑞宝商业广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1363/1363-连云港嘉瑞宝商业广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1363/1363-连云港嘉瑞宝商业广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1728",
    "province": "江苏省",
    "city": "连云港市",
    "name": "连云港海州大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1728/1728-连云港海州大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1728/1728-连云港海州大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1887",
    "province": "江苏省",
    "city": "镇江市",
    "name": "丹阳梦想城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1887/1887-丹阳梦想城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1887/1887-丹阳梦想城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1155",
    "province": "江苏省",
    "city": "镇江市",
    "name": "江苏镇江八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1155/1155-江苏镇江八佰伴店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1155/1155-江苏镇江八佰伴店-经营许可证.jpeg"
    }
  },
  {
    "id": "1289",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1289/1289-镇江万达店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1289/1289-镇江万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1972",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江丹阳丹曻邻里中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1972/1972-镇江丹阳丹曻邻里中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1972/1972-镇江丹阳丹曻邻里中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1368",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江丹阳吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1368/1368-镇江丹阳吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1368/1368-镇江丹阳吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1969",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江丹阳大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1969/1969-镇江丹阳大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1969/1969-镇江丹阳大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1803",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江丹阳金鹰天地广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1803/1803-镇江丹阳金鹰天地广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1803/1803-镇江丹阳金鹰天地广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1294",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江凤凰广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1294/1294-镇江凤凰广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1294/1294-镇江凤凰广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1383",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江句容吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1383/1383-镇江句容吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1383/1383-镇江句容吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1367",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江句容御东国际店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1367/1367-镇江句容御东国际店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1367/1367-镇江句容御东国际店-经营许可证.jpg"
    }
  },
  {
    "id": "1302",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1302/1302-镇江吾悦广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1302/1302-镇江吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1307",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1307/1307-镇江宝龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1307/1307-镇江宝龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1815",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江市大港大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1815/1815-镇江市大港大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1815/1815-镇江市大港大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1674",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江悦然广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1674/1674-镇江悦然广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1674/1674-镇江悦然广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1816",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江永隆城市广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1816/1816-镇江永隆城市广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1816/1816-镇江永隆城市广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1A26",
    "province": "江苏省",
    "city": "镇江市",
    "name": "镇江红豆万花城购物广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A26/1A26-镇江红豆万花城购物广场店-营业执照.jpg",
      "food": ""
    }
  },
  {
    "id": "1987",
    "province": "江西省",
    "city": "上饶市",
    "name": "上饶万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1987/1987-上饶万达广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1987/1987-上饶万达广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1A07",
    "province": "江西省",
    "city": "上饶市",
    "name": "上饶余干鹏颂百货店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A07/1A07-上饶余干鹏颂百货店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A07/1A07-上饶余干鹏颂百货店-经营许可证.jpg"
    }
  },
  {
    "id": "1614",
    "province": "江西省",
    "city": "九江市",
    "name": "江西九江联盛快乐城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1614/1614-江西九江联盛快乐城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1614/1614-江西九江联盛快乐城店-经营许可证.jpg"
    }
  },
  {
    "id": "1615",
    "province": "江西省",
    "city": "景德镇市",
    "name": "景德镇华润万家店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1615/1615-景德镇华润万家店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1615/1615-景德镇华润万家店-经营许可证.jpeg"
    }
  },
  {
    "id": "1771",
    "province": "江西省",
    "city": "景德镇市",
    "name": "景德镇金鼎广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1771/1771-景德镇金鼎广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1771/1771-景德镇金鼎广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1323",
    "province": "河南省",
    "city": "信阳市",
    "name": "信阳光山西亚广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1323/1323-信阳光山西亚广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1323/1323-信阳光山西亚广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1595",
    "province": "河南省",
    "city": "信阳市",
    "name": "信阳固始万佳店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1595/1595-信阳固始万佳店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1595/1595-信阳固始万佳店-经营许可证.jpg"
    }
  },
  {
    "id": "1087",
    "province": "河南省",
    "city": "信阳市",
    "name": "信阳固始古城路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1087/1087-信阳固始古城路店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1087/1087-信阳固始古城路店-经营许可证.jpeg"
    }
  },
  {
    "id": "1086",
    "province": "河南省",
    "city": "信阳市",
    "name": "信阳市固始县店.",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1086/1086-信阳市固始县店.-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1086/1086-信阳市固始县店.-经营许可证.jpeg"
    }
  },
  {
    "id": "1709",
    "province": "河南省",
    "city": "信阳市",
    "name": "信阳息县荣誉购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1709/1709-信阳息县荣誉购物中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1709/1709-信阳息县荣誉购物中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1672",
    "province": "河南省",
    "city": "信阳市",
    "name": "信阳新天地广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1672/1672-信阳新天地广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1672/1672-信阳新天地广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1596",
    "province": "河南省",
    "city": "信阳市",
    "name": "信阳美邻荟店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1596/1596-信阳美邻荟店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1596/1596-信阳美邻荟店-经营许可证.jpg"
    }
  },
  {
    "id": "1A00",
    "province": "河南省",
    "city": "信阳市",
    "name": "光山县慧泉广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A00/1A00-光山县慧泉广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A00/1A00-光山县慧泉广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1975",
    "province": "河南省",
    "city": "信阳市",
    "name": "河南信阳西亚城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1975/1975-河南信阳西亚城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1975/1975-河南信阳西亚城店-经营许可证.jpg"
    }
  },
  {
    "id": "1914",
    "province": "河南省",
    "city": "南阳市",
    "name": "河南南阳摩根吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1914/1914-河南南阳摩根吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1914/1914-河南南阳摩根吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1883",
    "province": "河南省",
    "city": "周口市",
    "name": "周口川汇万顺达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1883/1883-周口川汇万顺达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1883/1883-周口川汇万顺达店-经营许可证.jpg"
    }
  },
  {
    "id": "1375",
    "province": "河南省",
    "city": "周口市",
    "name": "周口沈丘新田360广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1375/1375-周口沈丘新田360广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1375/1375-周口沈丘新田360广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1594",
    "province": "河南省",
    "city": "商丘市",
    "name": "商丘夏邑家庭茂店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1594/1594-商丘夏邑家庭茂店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1594/1594-商丘夏邑家庭茂店-经营许可证.jpg"
    }
  },
  {
    "id": "1679",
    "province": "河南省",
    "city": "商丘市",
    "name": "商丘柘城欣泰百货店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1679/1679-商丘柘城欣泰百货店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1679/1679-商丘柘城欣泰百货店-经营许可证.jpg"
    }
  },
  {
    "id": "1951",
    "province": "河南省",
    "city": "商丘市",
    "name": "商丘梁园正弘汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1951/1951-商丘梁园正弘汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1951/1951-商丘梁园正弘汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1632",
    "province": "河南省",
    "city": "商丘市",
    "name": "商丘民权县圣晖广场餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1632/1632-商丘民权县圣晖广场餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1632/1632-商丘民权县圣晖广场餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1773",
    "province": "河南省",
    "city": "商丘市",
    "name": "商丘金博大商场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1773/1773-商丘金博大商场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1773/1773-商丘金博大商场店-经营许可证.jpg"
    }
  },
  {
    "id": "1873",
    "province": "河南省",
    "city": "郑州市",
    "name": "河南农业大学龙子湖校区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1873/1873-河南农业大学龙子湖校区店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1873/1873-河南农业大学龙子湖校区店-经营许可证.jpeg"
    }
  },
  {
    "id": "1543",
    "province": "浙江省",
    "city": "丽水市",
    "name": "丽水市万地广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1543/1543-丽水市万地广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1543/1543-丽水市万地广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1578",
    "province": "浙江省",
    "city": "丽水市",
    "name": "丽水绿谷信息产业园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1578/1578-丽水绿谷信息产业园店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1578/1578-丽水绿谷信息产业园店-经营许可证.jpg"
    }
  },
  {
    "id": "1366",
    "province": "浙江省",
    "city": "丽水市",
    "name": "丽水银泰城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1366/1366-丽水银泰城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1366/1366-丽水银泰城店-经营许可证.jpg"
    }
  },
  {
    "id": "1391",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州中盛广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1391/1391-台州中盛广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1391/1391-台州中盛广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1344",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州临海靖江中心广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1344/1344-台州临海靖江中心广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1344/1344-台州临海靖江中心广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1258",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州市立医院餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1258/1258-台州市立医院餐厅-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1258/1258-台州市立医院餐厅-经营许可证.jpeg"
    }
  },
  {
    "id": "1295",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州意得广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1295/1295-台州意得广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1295/1295-台州意得广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1736",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州新世纪广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1736/1736-台州新世纪广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1736/1736-台州新世纪广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1733",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州温岭CC广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1733/1733-台州温岭CC广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1733/1733-台州温岭CC广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1925",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州温岭九龙店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1925/1925-台州温岭九龙店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1925/1925-台州温岭九龙店-经营许可证.jpg"
    }
  },
  {
    "id": "1750",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州温岭宝龙店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1750/1750-台州温岭宝龙店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1750/1750-台州温岭宝龙店-经营许可证.jpeg"
    }
  },
  {
    "id": "1841",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州玉环吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1841/1841-台州玉环吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1841/1841-台州玉环吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1262",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州经开万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1262/1262-台州经开万达广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1262/1262-台州经开万达广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1871",
    "province": "浙江省",
    "city": "台州市",
    "name": "台州黄岩吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1871/1871-台州黄岩吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1871/1871-台州黄岩吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1418",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴中关村广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1418/1418-嘉兴中关村广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1418/1418-嘉兴中关村广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1713",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴中山路八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1713/1713-嘉兴中山路八佰伴店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1713/1713-嘉兴中山路八佰伴店-经营许可证.png"
    }
  },
  {
    "id": "1976",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴华府八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1976/1976-嘉兴华府八佰伴店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1976/1976-嘉兴华府八佰伴店-经营许可证.jpg"
    }
  },
  {
    "id": "1731",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴合乐城广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1731/1731-嘉兴合乐城广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1731/1731-嘉兴合乐城广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1846",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴嘉善中山西路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1846/1846-嘉兴嘉善中山西路店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1846/1846-嘉兴嘉善中山西路店-经营许可证.png"
    }
  },
  {
    "id": "1854",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴嘉善江南邻里中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1854/1854-嘉兴嘉善江南邻里中心店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1854/1854-嘉兴嘉善江南邻里中心店-经营许可证.png"
    }
  },
  {
    "id": "1855",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴市嘉善县星座标盒马店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1855/1855-嘉兴市嘉善县星座标盒马店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1855/1855-嘉兴市嘉善县星座标盒马店-经营许可证.jpg"
    }
  },
  {
    "id": "1810",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴平湖八佰伴店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1810/1810-嘉兴平湖八佰伴店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1810/1810-嘉兴平湖八佰伴店-经营许可证.jpg"
    }
  },
  {
    "id": "1419",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴平湖吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1419/1419-嘉兴平湖吾悦广场店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1419/1419-嘉兴平湖吾悦广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1899",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴星河COCOCity购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1899/1899-嘉兴星河COCOCity购物中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1899/1899-嘉兴星河COCOCity购物中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1387",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴杉杉in象店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1387/1387-嘉兴杉杉in象店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1387/1387-嘉兴杉杉in象店-经营许可证.jpeg"
    }
  },
  {
    "id": "1895",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴海宁许巷新大街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1895/1895-嘉兴海宁许巷新大街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1895/1895-嘉兴海宁许巷新大街店-经营许可证.jpg"
    }
  },
  {
    "id": "1935",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴海盐县云兴商业中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1935/1935-嘉兴海盐县云兴商业中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1935/1935-嘉兴海盐县云兴商业中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1393",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴海盐吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1393/1393-嘉兴海盐吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1393/1393-嘉兴海盐吾悦广场店-经营许可证.png"
    }
  },
  {
    "id": "1260",
    "province": "浙江省",
    "city": "嘉兴市",
    "name": "嘉兴龙鼎万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1260/1260-嘉兴龙鼎万达广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1260/1260-嘉兴龙鼎万达广场店-经营许可证.png"
    }
  },
  {
    "id": "1886",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波世纪东方店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1886/1886-宁波世纪东方店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1886/1886-宁波世纪东方店-经营许可证.jpg"
    }
  },
  {
    "id": "1830",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波世纪金源店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1830/1830-宁波世纪金源店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1830/1830-宁波世纪金源店-经营许可证.jpg"
    }
  },
  {
    "id": "1923",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波余姚万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1923/1923-宁波余姚万达广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1923/1923-宁波余姚万达广场店-经营许可证.png"
    }
  },
  {
    "id": "1918",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波余姚五彩城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1918/1918-宁波余姚五彩城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1918/1918-宁波余姚五彩城店-经营许可证.jpg"
    }
  },
  {
    "id": "1835",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波北仑大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1835/1835-宁波北仑大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1835/1835-宁波北仑大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1938",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波北仑银泰城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1938/1938-宁波北仑银泰城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1938/1938-宁波北仑银泰城店-经营许可证.jpg"
    }
  },
  {
    "id": "1827",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波南部商务区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1827/1827-宁波南部商务区店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1827/1827-宁波南部商务区店-经营许可证.jpg"
    }
  },
  {
    "id": "1944",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波和丰创意广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1944/1944-宁波和丰创意广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1944/1944-宁波和丰创意广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1933",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波天一广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1933/1933-宁波天一广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1933/1933-宁波天一广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1831",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波宁兴财富广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1831/1831-宁波宁兴财富广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1831/1831-宁波宁兴财富广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1977",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波宁海桃源广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1977/1977-宁波宁海桃源广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1977/1977-宁波宁海桃源广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1849",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波山丘市集店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1849/1849-宁波山丘市集店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1849/1849-宁波山丘市集店-经营许可证.jpeg"
    }
  },
  {
    "id": "1877",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波慈溪利时广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1877/1877-宁波慈溪利时广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1877/1877-宁波慈溪利时广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1848",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波慈溪吾悦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1848/1848-宁波慈溪吾悦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1848/1848-宁波慈溪吾悦店-经营许可证.jpg"
    }
  },
  {
    "id": "1845",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波慈溪天鸿大厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1845/1845-宁波慈溪天鸿大厦店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1845/1845-宁波慈溪天鸿大厦店-经营许可证.jpeg"
    }
  },
  {
    "id": "1880",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波慈溪银泰城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1880/1880-宁波慈溪银泰城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1880/1880-宁波慈溪银泰城店-经营许可证.jpg"
    }
  },
  {
    "id": "1828",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波来福士广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1828/1828-宁波来福士广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1828/1828-宁波来福士广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1834",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波江北万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1834/1834-宁波江北万达广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1834/1834-宁波江北万达广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1833",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波海曙龙湖天街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1833/1833-宁波海曙龙湖天街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1833/1833-宁波海曙龙湖天街店-经营许可证.jpeg"
    }
  },
  {
    "id": "1A13",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波观海卫悦美广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A13/1A13-宁波观海卫悦美广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A13/1A13-宁波观海卫悦美广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1872",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波镇海万科广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1872/1872-宁波镇海万科广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1872/1872-宁波镇海万科广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1A14",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波骆驼印象汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A14/1A14-宁波骆驼印象汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A14/1A14-宁波骆驼印象汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1829",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波高新宝龙店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1829/1829-宁波高新宝龙店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1829/1829-宁波高新宝龙店-经营许可证.jpeg"
    }
  },
  {
    "id": "1832",
    "province": "浙江省",
    "city": "宁波市",
    "name": "宁波高鑫广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1832/1832-宁波高鑫广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1832/1832-宁波高鑫广场店-经营许可证.jpg"
    }
  },
  {
    "id": "C021",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州MCC店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/38a1c4a0-ba5b-45c5-88da-18f069fbd0651662621925808.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/17f7a675-72c0-42b5-bae1-b10548dcebf51662621926136.jpg"
    }
  },
  {
    "id": "C037",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州丁兰广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3466dcf5-605d-41d1-982e-5884368d7bb11701968873087.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fdcd3e78-f2c9-447f-b72a-1aa00c1290211701968874631.pdf"
    }
  },
  {
    "id": "C018",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州三坝地铁站店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a6362fbc-0a9d-4494-bc32-d3b6a3a091bf1662621921981.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3b8db5d0-4561-41c7-8a74-f6b5d26bc0ee1662621922634.jpg"
    }
  },
  {
    "id": "C005",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州中大银泰店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aee1f9b7-24e0-4185-966b-c9275c0727ad1778346507575.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c92e76bc-7d15-4ba0-a1ba-3d809fb414191778346515695.pdf"
    }
  },
  {
    "id": "C028",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州临平银泰店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5a292238-ee7a-4e59-8b65-cd975ce6b9821701968864867.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4321c60a-9710-476e-be77-5c26d82bffb61701968865408.jpg"
    }
  },
  {
    "id": "C022",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州乐天城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e1649ecd-d4a3-4d53-90fa-f757bb86576f1662621927130.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/716c7968-31ca-4cad-8f54-ae702fa1beb91662621927543.jpg"
    }
  },
  {
    "id": "C003",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州华星广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b1f5ebfe-a25d-49bf-9e0e-9cd8e7170b601777136879088.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/780fc342-183b-4e44-bc66-bab1b5fbf7321777136879465.jpg"
    }
  },
  {
    "id": "C007",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州奥体印象城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1631bd56-9da9-411f-a8a2-bab482c095e31662621908031.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dd3deddf-17f3-4040-8fe0-4717e3d78fe21662621908376.jpg"
    }
  },
  {
    "id": "C029",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州尊宝大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ebfb04b2-4200-4dbe-90d3-5141b22e46301701968866356.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ee3ce946-0f3d-48de-92a6-37cc7a82a6451701968867188.pdf"
    }
  },
  {
    "id": "C039",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州庆春万家店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a3c5a384-62e9-4311-9de5-58e691a8b35c1704215396966.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a8657bd0-0808-4bac-8558-e99bf62a477e1704215397535.jpg"
    }
  },
  {
    "id": "C012",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州庆春联华店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fa9e399b-44bc-409b-9145-44e6cd06d11b1662621914057.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e5b575a5-0f3f-40bd-a67a-f03f5ca298741662621914410.jpg"
    }
  },
  {
    "id": "1396",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州建德恒太城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1396/1396-杭州建德恒太城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1396/1396-杭州建德恒太城店-经营许可证.png"
    }
  },
  {
    "id": "C020",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州支付宝店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ff506d94-7914-4925-9fb8-d10c06e3ea371662621924623.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a11805a8-c67a-4ce9-adad-e8f8bb718f251662621924948.jpg"
    }
  },
  {
    "id": "C014",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州朝龙汇店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/680dd554-ffb9-4beb-8c3c-efe46660edb01662621916589.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e09a1c47-f343-42d0-bfbc-8deeddeb8e211662621917067.jpg"
    }
  },
  {
    "id": "C024",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州杭行荟店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e9b50fcb-dbb4-4ee1-bddf-8216562e9b071662621929798.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c9013007-c3aa-42d5-8247-e7fccfe190ef1662621930143.jpg"
    }
  },
  {
    "id": "1401",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州桐庐大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1401/1401-杭州桐庐大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1401/1401-杭州桐庐大润发店-经营许可证.png"
    }
  },
  {
    "id": "C006",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州欧美金融城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5743812d-43af-4e51-8c9b-a90a281d9f371662621906787.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/浙江老乡鸡餐饮有限公司杭州欧美店_01(1)1777076079939.png"
    }
  },
  {
    "id": "C031",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州浦沿店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b6d2a419-5001-4148-b964-76b441ed30ac1701968868901.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/05b880e1-ecbd-44e6-81bc-09113cc5ccc71701968869447.jpg"
    }
  },
  {
    "id": "C004",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州海创大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/658f2f94-b037-4467-9f20-962acc2e2c1e1778346490259.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7e2932c6-2a1a-4871-883d-f1f2bbf327621778346490936.jpg"
    }
  },
  {
    "id": "C008",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州滨江儿童医院店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3e7d6a85-928b-4434-9c54-7242715e872d1662621909295.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a9c3db7c-e766-4123-998b-def22a4163aa1662621909623.jpg"
    }
  },
  {
    "id": "C001",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州滨江天街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6423ddfb-ee68-4a75-b027-69a9279ece9b1765300336837.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9d2e04a6-e4ce-4491-836a-ea328d0d7f801765300345789.pdf"
    }
  },
  {
    "id": "C038",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州滨江宝龙店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f1fe3818-2a13-42fe-918b-94f735e781d91701968876044.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6bb30744-e271-4483-a28a-b1756902ef0f1701968876487.jpg"
    }
  },
  {
    "id": "C010",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州绿谷店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f8297871-dc2a-4a7d-b124-9d072fc4d93f1662621911713.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fb7fd682-0032-48a9-8479-c0eab4317fbb1662621912049.jpg"
    }
  },
  {
    "id": "C013",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州联合中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7b444047-727e-499c-b209-4e68b27e33d01662621915272.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cbe37faa-c5ea-4936-bc45-f580ae6f58d41662621915691.jpg"
    }
  },
  {
    "id": "C009",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州花园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/527eb31f-7bca-4bfc-a280-3cc0ae6d5f871662621910539.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8193644b-68e9-48cd-8649-1666d0ef90c61662621910865.jpg"
    }
  },
  {
    "id": "C011",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州西港新界店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/47fb8040-3e9f-4821-bb4c-770fc71f578a1687021350848.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/326e4726-5404-49e1-8ab2-13f910434c941687021351183.jpg"
    }
  },
  {
    "id": "C052",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州西湖文化广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/store-center/images/营业执照，042，040，，044043052_81720420050349.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/store-center/images/微信图片_202407081431421720420318431.png"
    }
  },
  {
    "id": "C023",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州西溪谷店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3663eaac-b51a-4d74-9cb3-9d97e8cad4201662621928527.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3ccf628f-5828-4e6b-8d1a-bfac58043ecb1662621928926.jpg"
    }
  },
  {
    "id": "C019",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州西田城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a22c1399-7471-4af3-93f2-cfb8e44919761662621923486.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/11c76c85-5ff1-40d0-8e8d-14437933b5cf1662621923810.jpg"
    }
  },
  {
    "id": "C016",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州西联广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9817a1fc-038b-45ad-bd10-0e3e326c7f101662621919507.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8508b78b-111d-46c0-9021-4e5a147842521662621919835.jpg"
    }
  },
  {
    "id": "C032",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州运河上街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/42e112f7-8762-493c-9491-c6180ca36e8c1701968870376.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7428cf97-b80c-412e-a593-ae45e7f518181701968871226.pdf"
    }
  },
  {
    "id": "C030",
    "province": "浙江省",
    "city": "杭州市",
    "name": "杭州金街美地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f5bceef1-0b07-4995-9ed0-836d16e1bbb71685811745806.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7603fb5d-1b29-4317-a479-5e884e2c8ca11685811746622.pdf"
    }
  },
  {
    "id": "C025",
    "province": "浙江省",
    "city": "杭州市",
    "name": "滨江银泰店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8f837fee-7506-47e7-bec8-d58cc9aac0b01662621931012.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d8ab8cf4-b984-4eaa-8ddc-6cc6d47574cb1662621931544.jpg"
    }
  },
  {
    "id": "1288",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州东吴银泰店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1288/1288-湖州东吴银泰店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1288/1288-湖州东吴银泰店-经营许可证.jpg"
    }
  },
  {
    "id": "1988",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1988/1988-湖州吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1988/1988-湖州吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1328",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州安吉万象城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1328/1328-湖州安吉万象城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1328/1328-湖州安吉万象城店-经营许可证.png"
    }
  },
  {
    "id": "1613",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州安吉春天尚居景尚街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1613/1613-湖州安吉春天尚居景尚街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1613/1613-湖州安吉春天尚居景尚街店-经营许可证.jpeg"
    }
  },
  {
    "id": "1427",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州德清儿童医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1427/1427-湖州德清儿童医院店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1427/1427-湖州德清儿童医院店-经营许可证.jpg"
    }
  },
  {
    "id": "1577",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州德清正翔广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1577/1577-湖州德清正翔广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1577/1577-湖州德清正翔广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1859",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州星火外滩广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1859/1859-湖州星火外滩广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1859/1859-湖州星火外滩广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1397",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州浙北大厦东迁店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1397/1397-湖州浙北大厦东迁店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1397/1397-湖州浙北大厦东迁店-经营许可证.jpg"
    }
  },
  {
    "id": "1293",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州浙北大厦运动中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1293/1293-湖州浙北大厦运动中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1293/1293-湖州浙北大厦运动中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1971",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州爱山广场步行街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1971/1971-湖州爱山广场步行街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1971/1971-湖州爱山广场步行街店-经营许可证.jpg"
    }
  },
  {
    "id": "1756",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州爱山浙北大厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1756/1756-湖州爱山浙北大厦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1756/1756-湖州爱山浙北大厦店-经营许可证.jpg"
    }
  },
  {
    "id": "1715",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州织里春风长住盒马店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1715/1715-湖州织里春风长住盒马店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1715/1715-湖州织里春风长住盒马店-经营许可证.jpg"
    }
  },
  {
    "id": "1802",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州织里财富大厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1802/1802-湖州织里财富大厦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1802/1802-湖州织里财富大厦店-经营许可证.jpg"
    }
  },
  {
    "id": "1309",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州织里长安路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1309/1309-湖州织里长安路店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1309/1309-湖州织里长安路店-经营许可证.jpeg"
    }
  },
  {
    "id": "1642",
    "province": "浙江省",
    "city": "湖州市",
    "name": "湖州长兴九汇城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1642/1642-湖州长兴九汇城店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1642/1642-湖州长兴九汇城店-经营许可证.jpeg"
    }
  },
  {
    "id": "1768",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴万达广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1768/1768-绍兴万达广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1768/1768-绍兴万达广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1286",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴上虞大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1286/1286-绍兴上虞大润发店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1286/1286-绍兴上虞大润发店-经营许可证.jpeg"
    }
  },
  {
    "id": "1912",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴世茂大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1912/1912-绍兴世茂大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1912/1912-绍兴世茂大润发店-经营许可证.png"
    }
  },
  {
    "id": "1724",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴利佰家广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1724/1724-绍兴利佰家广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1724/1724-绍兴利佰家广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1257",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴大悦城餐厅",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1257/1257-绍兴大悦城餐厅-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1257/1257-绍兴大悦城餐厅-经营许可证.jpg"
    }
  },
  {
    "id": "1853",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴天悦城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1853/1853-绍兴天悦城店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1853/1853-绍兴天悦城店-经营许可证.png"
    }
  },
  {
    "id": "1671",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴宝龙天地店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1671/1671-绍兴宝龙天地店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1671/1671-绍兴宝龙天地店-经营许可证.jpeg"
    }
  },
  {
    "id": "1410",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1410/1410-绍兴宝龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1410/1410-绍兴宝龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1755",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴嵊州吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1755/1755-绍兴嵊州吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1755/1755-绍兴嵊州吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1867",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴嵊州西皮茂店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1867/1867-绍兴嵊州西皮茂店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1867/1867-绍兴嵊州西皮茂店-经营许可证.jpg"
    }
  },
  {
    "id": "1579",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴永利中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1579/1579-绍兴永利中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1579/1579-绍兴永利中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1749",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴港越路店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1749/1749-绍兴港越路店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1749/1749-绍兴港越路店-经营许可证.jpeg"
    }
  },
  {
    "id": "1907",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴滨海万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1907/1907-绍兴滨海万达店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1907/1907-绍兴滨海万达店-经营许可证.jpeg"
    }
  },
  {
    "id": "1800",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴滨海商业中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1800/1800-绍兴滨海商业中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1800/1800-绍兴滨海商业中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1631",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴祥源广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1631/1631-绍兴祥源广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1631/1631-绍兴祥源广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1310",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴紫金广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1310/1310-绍兴紫金广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1310/1310-绍兴紫金广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1357",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴自在天地店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1357/1357-绍兴自在天地店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1357/1357-绍兴自在天地店-经营许可证.jpeg"
    }
  },
  {
    "id": "1929",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴诸暨大唐大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1929/1929-绍兴诸暨大唐大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1929/1929-绍兴诸暨大唐大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1666",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴诸暨宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1666/1666-绍兴诸暨宝龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1666/1666-绍兴诸暨宝龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1531",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴都市春天店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1531/1531-绍兴都市春天店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1531/1531-绍兴都市春天店-经营许可证.jpeg"
    }
  },
  {
    "id": "1407",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴银泰城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1407/1407-绍兴银泰城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1407/1407-绍兴银泰城店-经营许可证.jpg"
    }
  },
  {
    "id": "1402",
    "province": "浙江省",
    "city": "绍兴市",
    "name": "绍兴颐高广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1402/1402-绍兴颐高广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1402/1402-绍兴颐高广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1792",
    "province": "浙江省",
    "city": "舟山市",
    "name": "舟山宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1792/1792-舟山宝龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1792/1792-舟山宝龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1784",
    "province": "浙江省",
    "city": "舟山市",
    "name": "舟山普陀大润发店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1784/1784-舟山普陀大润发店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1784/1784-舟山普陀大润发店-经营许可证.jpg"
    }
  },
  {
    "id": "1993",
    "province": "浙江省",
    "city": "舟山市",
    "name": "舟山海山广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1993/1993-舟山海山广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1993/1993-舟山海山广场店-经营许可证.png"
    }
  },
  {
    "id": "1429",
    "province": "浙江省",
    "city": "衢州市",
    "name": "衢州南湖东方商厦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1429/1429-衢州南湖东方商厦店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1429/1429-衢州南湖东方商厦店-经营许可证.jpeg"
    }
  },
  {
    "id": "1339",
    "province": "浙江省",
    "city": "金华市",
    "name": "义乌万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1339/1339-义乌万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1339/1339-义乌万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1412",
    "province": "浙江省",
    "city": "金华市",
    "name": "义乌之心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1412/1412-义乌之心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1412/1412-义乌之心店-经营许可证.jpg"
    }
  },
  {
    "id": "1878",
    "province": "浙江省",
    "city": "金华市",
    "name": "义乌吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1878/1878-义乌吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1878/1878-义乌吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1665",
    "province": "浙江省",
    "city": "金华市",
    "name": "义乌市银海店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1665/1665-义乌市银海店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1665/1665-义乌市银海店-经营许可证.jpg"
    }
  },
  {
    "id": "1324",
    "province": "浙江省",
    "city": "金华市",
    "name": "义乌新光汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1324/1324-义乌新光汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1324/1324-义乌新光汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1656",
    "province": "浙江省",
    "city": "金华市",
    "name": "义乌江东广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1656/1656-义乌江东广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1656/1656-义乌江东广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1860",
    "province": "浙江省",
    "city": "金华市",
    "name": "横店万盛南街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1860/1860-横店万盛南街店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1860/1860-横店万盛南街店-经营许可证.jpg"
    }
  },
  {
    "id": "1284",
    "province": "浙江省",
    "city": "金华市",
    "name": "浙江金华万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1284/1284-浙江金华万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1284/1284-浙江金华万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1774",
    "province": "浙江省",
    "city": "金华市",
    "name": "金华义乌北苑商贸区店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1774/1774-金华义乌北苑商贸区店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1774/1774-金华义乌北苑商贸区店-经营许可证.jpeg"
    }
  },
  {
    "id": "1630",
    "province": "浙江省",
    "city": "金华市",
    "name": "金华之心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1630/1630-金华之心店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1630/1630-金华之心店-经营许可证.png"
    }
  },
  {
    "id": "1730",
    "province": "浙江省",
    "city": "金华市",
    "name": "金华兰溪宝龙广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1730/1730-金华兰溪宝龙广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1730/1730-金华兰溪宝龙广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1532",
    "province": "浙江省",
    "city": "金华市",
    "name": "金华和悦邻里店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1532/1532-金华和悦邻里店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1532/1532-金华和悦邻里店-经营许可证.jpeg"
    }
  },
  {
    "id": "1664",
    "province": "浙江省",
    "city": "金华市",
    "name": "金华武义金湖中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1664/1664-金华武义金湖中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1664/1664-金华武义金湖中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1557",
    "province": "浙江省",
    "city": "金华市",
    "name": "金华永康世贸中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1557/1557-金华永康世贸中心店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1557/1557-金华永康世贸中心店-经营许可证.jpg"
    }
  },
  {
    "id": "1423",
    "province": "浙江省",
    "city": "金华市",
    "name": "金华永盛购物广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1423/1423-金华永盛购物广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1423/1423-金华永盛购物广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1369",
    "province": "湖北省",
    "city": "仙桃市",
    "name": "仙桃吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1369/1369-仙桃吾悦广场店-营业执照.png",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1369/1369-仙桃吾悦广场店-经营许可证.png"
    }
  },
  {
    "id": "1403",
    "province": "湖北省",
    "city": "仙桃市",
    "name": "仙桃武商MAll店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1403/1403-仙桃武商MAll店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1403/1403-仙桃武商MAll店-经营许可证.jpg"
    }
  },
  {
    "id": "1361",
    "province": "湖北省",
    "city": "孝感市",
    "name": "孝感万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1361/1361-孝感万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1361/1361-孝感万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1948",
    "province": "湖北省",
    "city": "孝感市",
    "name": "孝感乾坤大道店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1948/1948-孝感乾坤大道店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1948/1948-孝感乾坤大道店-经营许可证.jpeg"
    }
  },
  {
    "id": "1785",
    "province": "湖北省",
    "city": "孝感市",
    "name": "孝感保丽广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1785/1785-孝感保丽广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1785/1785-孝感保丽广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1685",
    "province": "湖北省",
    "city": "孝感市",
    "name": "孝感吾悦店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1685/1685-孝感吾悦店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1685/1685-孝感吾悦店-经营许可证.jpg"
    }
  },
  {
    "id": "1616",
    "province": "湖北省",
    "city": "孝感市",
    "name": "孝感恒泰购物广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1616/1616-孝感恒泰购物广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1616/1616-孝感恒泰购物广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1958",
    "province": "湖北省",
    "city": "宜昌市",
    "name": "宜昌国贸店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1958/1958-宜昌国贸店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1958/1958-宜昌国贸店-经营许可证.png"
    }
  },
  {
    "id": "1422",
    "province": "湖北省",
    "city": "宜昌市",
    "name": "宜昌宜都万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1422/1422-宜昌宜都万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1422/1422-宜昌宜都万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1625",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武商MALL·众圆",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1625/1625-武商MALL·众圆-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1625/1625-武商MALL·众圆-经营许可证.jpg"
    }
  },
  {
    "id": "1A08",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武商超市常青花园店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A08/1A08-武商超市常青花园店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A08/1A08-武商超市常青花园店-经营许可证.jpg"
    }
  },
  {
    "id": "6001",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉一店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/fe0d8708-168e-4d84-840d-a53b9c2a75b11764349536783.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bd3486f3-08c8-4877-957a-63551be6f96b1764349537050.jpg"
    }
  },
  {
    "id": "6120",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉万科嘉园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/537a8fe3-fd7b-4a9d-9f2b-7ec241061ff81744823124741.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/eef35d7f-95cd-44c4-aeba-54fe3288482d1744823125280.jpg"
    }
  },
  {
    "id": "6014",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉万科汉阳国际店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5db60a1c-ec4c-414c-8d9a-b7e386d3ce641662621601148.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/02f0a939-0f7c-45f5-a4ce-9affe24aac4d1662621601516.pdf"
    }
  },
  {
    "id": "6151",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉东合中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c80df199-dd44-442a-a3ea-69fa09ffb9a71768583183860.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/868c0b34-2e50-4b7d-9a89-33bf9b830c961768583184176.jpg"
    }
  },
  {
    "id": "6006",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉东湖景园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f09b4160-8c83-4fac-ba22-b90bd83f838b1662621589593.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/40395edd-1f39-4bc5-99d1-4751b8c258571662621590053.jpg"
    }
  },
  {
    "id": "6200",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉东湖网谷店",
    "licenses": {
      "business": "",
      "food": ""
    }
  },
  {
    "id": "6079",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉东澜岸店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/81210fa8-efb0-469e-9254-595154f57cd81701190897223.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f03d511b-2678-4a24-8870-f8c2235299821701190897872.jpg"
    }
  },
  {
    "id": "6045",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉中南路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5f97058a-31c5-44d7-a931-30a8ea0d50921698858210592.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b36a5e61-96b8-44d1-a028-a46749095ae91698858211092.JPG"
    }
  },
  {
    "id": "6010",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉中央公馆店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/08f44c69-ccd3-4ec8-aa35-abb08cbabc9e1662621595741.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/82b16a1a-06fa-4f5e-b325-afbd732bf64b1662621596130.pdf"
    }
  },
  {
    "id": "6194",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉中建御景星城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a5eb2717-310b-4682-aef4-dd868caed1e51701968801313.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a11870fc-cb92-40c8-baef-db48f57c7a121701968801906.pdf"
    }
  },
  {
    "id": "6078",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉中部名居店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/97d84ebf-06a5-4281-9e18-2a50a01c3fe81702659957643.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ba027e78-cfe1-47fe-bedd-6751694fcce51702659958131.png"
    }
  },
  {
    "id": "6094",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉五环广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2d140149-01d8-4646-9c52-881da75923cf1719507911441.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/0a2c491d-bd22-46d0-a027-c48153fc09c61719507911913.jpg"
    }
  },
  {
    "id": "6189",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉佳园路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/89e546d0-f6d2-4b12-9cf0-11ad4ab02c071662621741180.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/288d6a41-ec77-4477-a35f-59c0a0a2aa311662621741672.png"
    }
  },
  {
    "id": "6190",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉保利公园九里店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a730d693-79a1-4f25-9668-df9893ed4c251662621742647.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/52bde417-e127-4287-9a8a-4218b0b111621662621743010.jpg"
    }
  },
  {
    "id": "6004",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉保利心语店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6dbcf32c-f65c-48b5-9aeb-0108f1b5a4e21662621587080.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/dcd421b7-35a5-4ffa-8546-5177270ec2d41662621587484.jpg"
    }
  },
  {
    "id": "6127",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉保利时代天悦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/083fd246-3a89-4b3d-8557-71d735c382391752771954571.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/68247ec9-c424-4f16-b8cc-af6561b90fc71752771954842.jpg"
    }
  },
  {
    "id": "6149",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉关山坊店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/964b65c7-7610-43cc-9fcb-a795dd777f3b1763658390154.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6f6912bc-22d8-484e-9de1-b8409a7ba60f1763658390715.pdf"
    }
  },
  {
    "id": "6188",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉关山大道紫菘店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ed43ed02-fc99-4c7c-bf45-0eab9ab065f71662621739766.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f81de1fc-226b-43de-a3ca-2c0879c4d1a41662621740277.png"
    }
  },
  {
    "id": "6049",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉凌云广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9e1ca84e-11c4-4ef6-bc74-b772be730ceb1691946113879.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5ec497ef-4f2a-46c7-89e9-002bb28866b21691946114284.jpg"
    }
  },
  {
    "id": "6230",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉创新天地230街店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/iwEcAqNqcGcDAQTRCsEF0Qd3BrArJUJ2NkfKiwnoyVn9O3AAB9IHn-mICAAJomltCgAL0gAs2lk.jpg_720x720q901779671463799.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/iwEdAqNqcGcDAQTRCM4F0QZiBrAW6LXmjGjKRAnoyV_DPWIAB9IHn-mICAAJomltCgAL0gAiX4w.jpg_720x720q901779671485007.jpg"
    }
  },
  {
    "id": "6182",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉华林广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4e875448-45f9-4d77-ad5d-558231d8a86c1662621732460.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/78f0fe05-2e16-4773-b961-80589cecc7c31662621732811.jpg"
    }
  },
  {
    "id": "6174",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉华鼎丽都店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2fb7c28b-5656-447a-902a-ebe2eaeaa4871662621723475.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/10896259-82b5-456d-9c61-e814f0ce46b31662621723822.jpg"
    }
  },
  {
    "id": "6011",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉南湖大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/85eff97a-0a5f-49a3-8f49-fb02b5ccafe81662621597049.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1f950f0e-a3e4-4d45-b70e-48b85b19a5271662621597774.jpg"
    }
  },
  {
    "id": "6031",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉古田四路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c405b2d7-c127-4c21-b4af-04f25297b1fd1679101346553.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9402efa0-762a-4688-866c-60d86fa09bfc1679101346811.jpg"
    }
  },
  {
    "id": "6048",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉吴家山中心广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ae64a20b-e133-4e4f-a2db-ed081cc592301691427789715.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f254a063-dc9d-4208-b444-4f133ee43d261691427800924.jpg"
    }
  },
  {
    "id": "6191",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉国采中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d627178e-15ca-40bf-81f6-4ad3c80a672f1675443738052.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/254d07bb-4e01-4817-b8fb-dedab24730881675443738415.pdf"
    }
  },
  {
    "id": "6162",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉复地东湖国际店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/aa3b00c7-1cdf-4fe1-b085-1ce9d45c24f71662621713244.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/湖北老乡鸡餐饮有限公司复地东湖国际分公司食品经营许可证（正本）_11772598572359.jpg"
    }
  },
  {
    "id": "6089",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉奥山世纪城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/93d6cbdd-760e-4381-a387-5a338afdea4a1715187807324.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/778c68b2-8013-43eb-a734-6e562d33813b1715187807668.jpg"
    }
  },
  {
    "id": "6175",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉客厅店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a5c7b9ad-d202-4862-9734-0364c5d190591662621724649.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a7c7d0c5-b9dc-4597-a59a-0fcbfbe910e31662621724949.jpg"
    }
  },
  {
    "id": "6167",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉家盛时代店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8b6015af-8ea9-4b31-a1fb-33a13d529d3a1662621717760.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品经营许可证（副本） (1)1777119084716.pdf"
    }
  },
  {
    "id": "6021",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉市光谷世界城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/249f21fd-2530-4d04-9349-f7ea683744221662656729395.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e2b0d7c4-a300-4ac3-948e-3802e6849be71662656730207.jpg"
    }
  },
  {
    "id": "6019",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉市梅苑地铁站店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/048aa994-9899-43c6-8d00-022cab967faa1662621602447.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2996aaa6-a50a-48d6-8141-6bba5c7261c11662621602811.JPG"
    }
  },
  {
    "id": "6027",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉市清江山水店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b75badae-a25b-4214-8f72-9e44a6b8887f1668618342677.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/55541ae1-cff6-4520-9174-706fda2a4d191668618343133.jpg"
    }
  },
  {
    "id": "6104",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉广电兰亭时代店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/53d9453c-4e81-406d-9e2c-89b3419b2ef71726765541407.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ce4faa4d-d343-434a-8c03-126971f582231726765541701.jpg"
    }
  },
  {
    "id": "6002",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉座标城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/97198561-c5ef-42ad-b597-5c4dc41d5e341687280602232.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4c031d63-aed7-4c01-b558-79a76f26127d1687280615989.jpg"
    }
  },
  {
    "id": "6130",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉建设大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/33cd9f3f-c43c-4471-8f25-a80989ac3bc41753117536832.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8d166d1c-1878-4700-a259-e37fb81d4cf71753117537470.pdf"
    }
  },
  {
    "id": "6176",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉当代梦工厂店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/faf9cec7-8b95-46d0-b897-47850632d8311662621725777.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c6e668c2-7ace-4513-af71-3117ae373a451662621726068.jpg"
    }
  },
  {
    "id": "6033",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉御华园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/131cb088-5cc0-4efe-95c9-4dcd170415d41681350711960.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ebfb0aa5-a421-4e77-9c7e-2c17003bcd831681350712382.jpg"
    }
  },
  {
    "id": "6025",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉循礼门店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3a020e38-2281-4e1d-ac0f-d08bac77af971669050060436.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/78bc5691-165b-4be4-b8fa-29886bc7d3171669050060734.jpg"
    }
  },
  {
    "id": "6144",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉恒大城店",
    "licenses": {
      "business": "",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/许可证1765279785920.png"
    }
  },
  {
    "id": "6008",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉拓创大厦店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f5a56bea-af4a-413f-b8f5-990e5c15281c1727456707251.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e23f100e-940c-446d-b6e5-8186bb5c70631727456707558.jpg"
    }
  },
  {
    "id": "6178",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉文体汇店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b7f4bc96-1554-4efb-8b8c-c84037cf705e1662621728007.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/417df124-fe6b-43e8-832a-0d3678711a271662621728302.jpg"
    }
  },
  {
    "id": "6156",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉文化大道店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2130292e-ba76-4136-90f1-42f7917e43801773680787249.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c038a977-3307-4c17-8deb-f2cbd395507e1773680787554.jpg"
    }
  },
  {
    "id": "6192",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉新佳丽广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9f4aed9d-f74c-4355-a5fd-79ecd973bafb1708967305704.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ad7e7fa7-512b-4f4b-a780-a53429ca06dd1708967317235.pdf"
    }
  },
  {
    "id": "6179",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉星光时代店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f8f1b5de-4132-4053-ac93-90d371a71ee61662621729113.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3301b04e-d812-4f8d-9abe-77b32db6c1701662621729419.jpg"
    }
  },
  {
    "id": "6193",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉星汇维港店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e7bbf4e6-d64b-4569-8ee0-73225d41bee71701968799923.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/3b2fa893-9a63-47b6-a675-df94a48e94c31701968800360.pdf"
    }
  },
  {
    "id": "6135",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉未来科技城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e6d24a44-870e-48c9-9c0f-e43a4d183c6e1763658381128.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e4bc5d29-d75c-4ae5-8f6d-834024bbdc891763658381703.pdf"
    }
  },
  {
    "id": "6113",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉杨园南路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d357c677-ace3-497b-bfcc-411e82274ff21729703097645.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/43f122f4-3575-46ad-938a-6400e6834dc51729703098060.jpg"
    }
  },
  {
    "id": "6163",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉杨家湾店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bbf309c4-11ce-4ce6-ba5f-610fe3d16d471777741594142.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/291ce2a0-8cc7-4644-8742-93342541a0361777741594436.jpg"
    }
  },
  {
    "id": "1673",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉武商梦时代店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1673/1673-武汉武商梦时代店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1673/1673-武汉武商梦时代店-经营许可证.jpeg"
    }
  },
  {
    "id": "6128",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉武大科技园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e908dc8b-603a-4956-ab01-2cbd1c364f041751941306335.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9de6743e-2cb9-4c6b-a2a3-0e9badb520441751941306675.jpg"
    }
  },
  {
    "id": "6129",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉汉口融科天城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ca213d82-e37a-4778-a714-f498df6939531752771955926.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/17fc62ee-a0dd-434b-98c3-98ca1cd735dd1752771956257.jpg"
    }
  },
  {
    "id": "6154",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉汉街店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9ebe249f-0230-49bb-8de8-b1ffc0bacd401772471184472.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/74f5887a-e7fe-4358-9c19-23af1009c4091772471185079.pdf"
    }
  },
  {
    "id": "6105",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉汉阳欧亚达店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a580478e-dca9-4c06-af3e-b64f76ef5d551724778332597.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d3641115-e1ca-4870-81c2-4b39cd405a9b1724778332897.jpg"
    }
  },
  {
    "id": "6115",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉汉阳江腾荟店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1e5888f0-944e-4751-907e-0d77575b45b51732554317616.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/86ed4cfe-847a-4777-b98d-1e9ea519a9331732554317932.jpg"
    }
  },
  {
    "id": "6039",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉江汉区航空路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/17e3ce88-9d1c-4b51-9b56-ffbff821388e1679883052721.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c0bc0cec-0e19-416c-826e-d18c6dc82f3b1679883053021.jpg"
    }
  },
  {
    "id": "6217",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉沙港品悦荟店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/0cd37588d017fc9665daee24517a0fee1774005165633.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/1bf032e21e71687335aab7f62d069bca1774005174113.jpg"
    }
  },
  {
    "id": "6157",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉泛海国际店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/85757719-eba0-4ca8-94e1-08f76041d2311768669569483.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8a33060d-ff07-4e59-9352-98c13d065b061768669569763.jpg"
    }
  },
  {
    "id": "6058",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉洪山区保利城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/25592035-0a94-4935-a63d-99f396857ab41701882199705.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/97141505-942f-4fcc-9b8a-fcd7fe2c1d881701882201552.pdf"
    }
  },
  {
    "id": "6231",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉洪山广场地铁站店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/营业执照1778732871367.png",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/经营许可证1778732894597.png"
    }
  },
  {
    "id": "6143",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉海尔广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/178e528f-a994-4d85-9efd-1d862148b1501763658386027.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/42b43cad-9b4d-461b-bc81-ce6d4749e0771763658386590.pdf"
    }
  },
  {
    "id": "6153",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉现代世贸中心店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/ce3eaac6-1923-4c52-b2e2-b934b5ac8aa11772039174199.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bcc87566-cadd-467b-ade3-20425ef7c0191772039174589.png"
    }
  },
  {
    "id": "6141",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉白沙居然之家店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d9647385-7b1c-4917-a7fb-9fff0e85662a1662621689433.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d8b9bd89-f948-4fa3-a1f5-873817e215401662621689728.jpg"
    }
  },
  {
    "id": "6145",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉白沙洲烽火店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8443bd68-5794-48ab-81be-0285a440418e1763571979806.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/IMG_20251121_143051(3)1767152290769.jpg"
    }
  },
  {
    "id": "6148",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉百瑞景店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/19122ae3-aac3-4544-8feb-26ec620633a91763053586888.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d5d42518-9c98-4cf4-854b-2da76578a8621763053587721.pdf"
    }
  },
  {
    "id": "6116",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉盘龙城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/8648b37f-cd23-4a5e-9134-5dd3a20429cc1730567141653.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/89e68706-441f-4b30-8814-5abdb1442b281730567141958.jpg"
    }
  },
  {
    "id": "6065",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉红坊里店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/9b33e27a-9f95-4170-ae75-bd12f76702ea1698858223733.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f56a3f8e-1b3b-40d1-83e3-a3349ee6facb1698858224334.pdf"
    }
  },
  {
    "id": "6012",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉绿地新都会店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cde92f84-6378-4b8f-af67-81cd91bd1b571662621598689.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/d36f8615-2191-494e-8c19-ade629bac7aa1662621599058.jpg"
    }
  },
  {
    "id": "6219",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉绿地汉口中心店",
    "licenses": {
      "business": "",
      "food": ""
    }
  },
  {
    "id": "6131",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉花城大道二店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/bfe64573-fe4a-4fea-8a63-dd74406772f61756573533005.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/66197afe-fad8-4202-ae7e-a45ace7848ca1756573533330.jpg"
    }
  },
  {
    "id": "6044",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉花桥店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5652bd53-d060-4618-b31c-e2f17d69e6a91685811712799.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b35f3576-2a05-43c7-aedf-a49e44d9c7a71685811713559.jpg"
    }
  },
  {
    "id": "1A10",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉蔡甸服务区西店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A10/1A10-武汉蔡甸服务区西店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1A10/1A10-武汉蔡甸服务区西店-经营许可证.jpg"
    }
  },
  {
    "id": "6180",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉融创智谷店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e36c22c2-4b9f-40b5-be98-63c3acf1605e1662621730228.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/45f622fb-e8cd-4270-b12f-f0ce9c66c23e1662621730521.jpg"
    }
  },
  {
    "id": "6140",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉街道口未来城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f138d264-64f1-455d-b44b-8eaeddc641b61764781580483.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/cc3d4a5a-abc3-4c87-8b6b-67040dabc98e1764781581033.jpg"
    }
  },
  {
    "id": "6003",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉软件园店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/48df1670-e2e3-4716-b4dc-e3f5c5c07d151662621585872.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b91a2aa7-8feb-4746-aed6-d289a466a4dd1662621586188.jpg"
    }
  },
  {
    "id": "6032",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉软件新城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/16eb5f3c-23c8-422c-a604-2301898a01cf1683766097825.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/da3169b7-8cb0-46b9-99c4-6b76debfe8c81683766098204.jpg"
    }
  },
  {
    "id": "6102",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉金地广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2c1f4770-1aa2-4b94-9ff2-fd85a2d5f5d11721927122596.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4885f234-c8dd-4994-9296-119545f933051721927123140.jpg"
    }
  },
  {
    "id": "6119",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉金地艺境店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/32552275-3b6d-41e8-b7f0-92314381eee91763140055626.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/68356757-c70d-446a-bcaa-6406a7dae8ab1763140056043.jpg"
    }
  },
  {
    "id": "6020",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉金银湖店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/c3870ad0-b988-4bc1-8741-bfe8bceb05d31662621603671.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/42f0c5d9-4163-49a6-b9df-c4fb9ca7eb3b1662621604043.pdf"
    }
  },
  {
    "id": "6146",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉锦绣龙城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f9889f25-1f59-4646-85a2-c248d564d20b1763140067896.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/5074c986-1ce3-493d-836a-cc05fddfa5d11763140068202.jpg"
    }
  },
  {
    "id": "6170",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉长城汇店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/259c8b98-203b-4782-b6a8-36a4632b49351662621720055.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a7774fd6-591a-4f7a-9ccf-d00f05ebec151662621720347.jpg"
    }
  },
  {
    "id": "1825",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉长江存储CB大楼店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1825/1825-武汉长江存储CB大楼店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1825/1825-武汉长江存储CB大楼店-经营许可证.png"
    }
  },
  {
    "id": "6061",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉静安上城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/b308afd1-4fc8-4d14-ab9e-39b6c98fe6621699030910194.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f7ebaf2f-8b9b-4941-a598-c54a6a7b1c661699030911390.pdf"
    }
  },
  {
    "id": "6168",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉韵湖星光店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/2915c229-bdc9-46a3-9cdc-c37e6370bd041662621718905.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/a3e618d2-11fc-42ef-996f-292d263946781662621719204.jpg"
    }
  },
  {
    "id": "6218",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉马应龙医药产业园店",
    "licenses": {
      "business": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/马应龙营业执照1764586250129.pdf",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/马应龙食品经营许可证1764586298029.pdf"
    }
  },
  {
    "id": "6158",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉马湖创意天地店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/43efc867-5c9c-4b6a-a10f-849f5e5c778d1662621709645.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/IMG_20260327_1025581774578833396.jpg"
    }
  },
  {
    "id": "6132",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉鲁磨路店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/f9e0f5dc-1c9c-4d32-a1c5-eb64d8089dcd1752771957781.jpg",
      "food": "https://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/store-center/pro/images/食品经营许可证（正本）1769139818815.pdf"
    }
  },
  {
    "id": "6138",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉黄陂万象城店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/e4fd14db-b1b4-4164-94f3-a580446ebd091760115972730.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/6b79ec9d-1cc9-4c7b-aaa0-96eee50b34791760115972997.jpg"
    }
  },
  {
    "id": "6047",
    "province": "湖北省",
    "city": "武汉市",
    "name": "武汉黄陂广场店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/73b85220-3948-4acf-86ea-6f31933601831698858213280.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/7fbe82b3-d663-4c0b-80d6-e190a44d79f61698858213778.jpg"
    }
  },
  {
    "id": "6155",
    "province": "湖北省",
    "city": "武汉市",
    "name": "老乡鸡武汉绿地606店",
    "licenses": {
      "business": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/1b112dcf-7237-4c7f-bc32-161af6dfa3211773680786148.jpg",
      "food": "http://lxj-business-center-file.oss-cn-hangzhou.aliyuncs.com/archives-service/pro/images/4c9a6169-8771-4c0d-86b3-bbba72e4d3dd1773680786485.jpg"
    }
  },
  {
    "id": "1574",
    "province": "湖北省",
    "city": "荆州市",
    "name": "荆州人信汇店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1574/1574-荆州人信汇店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1574/1574-荆州人信汇店-经营许可证.jpg"
    }
  },
  {
    "id": "1573",
    "province": "湖北省",
    "city": "襄阳市",
    "name": "襄阳东津民发世纪广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1573/1573-襄阳东津民发世纪广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1573/1573-襄阳东津民发世纪广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1304",
    "province": "湖北省",
    "city": "襄阳市",
    "name": "襄阳天元四季城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1304/1304-襄阳天元四季城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1304/1304-襄阳天元四季城店-经营许可证.jpg"
    }
  },
  {
    "id": "1330",
    "province": "湖北省",
    "city": "襄阳市",
    "name": "襄阳民发广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1330/1330-襄阳民发广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1330/1330-襄阳民发广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1303",
    "province": "湖北省",
    "city": "襄阳市",
    "name": "襄阳环球金融城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1303/1303-襄阳环球金融城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1303/1303-襄阳环球金融城店-经营许可证.jpg"
    }
  },
  {
    "id": "1764",
    "province": "湖北省",
    "city": "鄂州市",
    "name": "鄂州吾悦广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1764/1764-鄂州吾悦广场店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1764/1764-鄂州吾悦广场店-经营许可证.jpg"
    }
  },
  {
    "id": "1761",
    "province": "湖北省",
    "city": "鄂州市",
    "name": "鄂州银泰百货店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1761/1761-鄂州银泰百货店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1761/1761-鄂州银泰百货店-经营许可证.jpeg"
    }
  },
  {
    "id": "1770",
    "province": "湖北省",
    "city": "随州市",
    "name": "随州中心医院店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1770/1770-随州中心医院店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1770/1770-随州中心医院店-经营许可证.png"
    }
  },
  {
    "id": "1682",
    "province": "湖北省",
    "city": "黄冈市",
    "name": "麻城黄商购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1682/1682-麻城黄商购物中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1682/1682-麻城黄商购物中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1331",
    "province": "湖北省",
    "city": "黄冈市",
    "name": "黄冈万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1331/1331-黄冈万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1331/1331-黄冈万达店-经营许可证.jpg"
    }
  },
  {
    "id": "1992",
    "province": "湖北省",
    "city": "黄冈市",
    "name": "黄冈武穴武商购物中心店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1992/1992-黄冈武穴武商购物中心店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1992/1992-黄冈武穴武商购物中心店-经营许可证.jpeg"
    }
  },
  {
    "id": "1024",
    "province": "湖北省",
    "city": "黄冈市",
    "name": "黄冈罗田义水外滩步行街店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1024/1024-黄冈罗田义水外滩步行街店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1024/1024-黄冈罗田义水外滩步行街店-经营许可证.png"
    }
  },
  {
    "id": "1735",
    "province": "湖北省",
    "city": "黄冈市",
    "name": "黄冈英山中央城店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1735/1735-黄冈英山中央城店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1735/1735-黄冈英山中央城店-经营许可证.jpg"
    }
  },
  {
    "id": "1807",
    "province": "湖北省",
    "city": "黄冈市",
    "name": "黄冈蕲春大中华购物广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1807/1807-黄冈蕲春大中华购物广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1807/1807-黄冈蕲春大中华购物广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1797",
    "province": "湖北省",
    "city": "黄冈市",
    "name": "黄冈黄梅鹏泰百货店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1797/1797-黄冈黄梅鹏泰百货店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1797/1797-黄冈黄梅鹏泰百货店-经营许可证.jpg"
    }
  },
  {
    "id": "1332",
    "province": "湖北省",
    "city": "黄石市",
    "name": "黄石大冶雨润广场店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1332/1332-黄石大冶雨润广场店-营业执照.jpeg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1332/1332-黄石大冶雨润广场店-经营许可证.jpeg"
    }
  },
  {
    "id": "1908",
    "province": "湖北省",
    "city": "黄石市",
    "name": "黄石武商MALL店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1908/1908-黄石武商MALL店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1908/1908-黄石武商MALL店-经营许可证.jpg"
    }
  },
  {
    "id": "1337",
    "province": "湖北省",
    "city": "黄石市",
    "name": "黄石港万达店",
    "licenses": {
      "business": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1337/1337-黄石港万达店-营业执照.jpg",
      "food": "https://lxj-sso-file.oss-cn-hangzhou.aliyuncs.com/sso-memo/file-jiameng/1337/1337-黄石港万达店-经营许可证.jpg"
    }
  }
];
