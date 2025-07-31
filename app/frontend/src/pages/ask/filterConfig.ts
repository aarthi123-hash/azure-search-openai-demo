// filterConfig.ts
export interface IDropdownOption {
  key: string;
  text: string;
}

export interface FilterConfig {
  programs: {
    [key: string]: {
      label: string;
      taskOrders?: string[];
    };
  };
}

export const filterConfig: FilterConfig = {
  programs: {
    ngansgsi: {
      label: "NGA NSG/SI"
    },
    ngaesmarts: {
      label: "NGA ESMARTS",
      taskOrders: ["TO01 X4", "TO02 XK", "TO03 Enterprise Operations", "TO04 OCIO CS"]
    },
    ngardas: {
      label: "NGA RDAS",
      taskOrders: [
  "TO04 Innovision",
  "TO11 International",
  "TO13 ITEMS",
  "TO17 Office of GEOINT Management",
  "TO18 Tactical Data Program Management",
  "TO21 ITEMS PMO",
  "TO24 NSG Expeditionary Architecture",
  "TO26 CAE",
  "TO27 Online GEOINT Services"
],
    },
    ngasein: {
      label: "NGA SEIN",
    },
    ngaseinasisb: {
      label: "NGA SEIN ASB",
    },
    odnicases: {
      label: "ODNI CASES",
      taskOrders: ["TO01 AT/CCT"]
    },
    ngaemerald: {
      label: "NGA EMERALD",
      taskOrders: [
  "TO02 National Technical Means",
  "TO06 Source Content Conveyance",
  "TO08 Research",
  "TO17 Source",
  "TO19 N2W",
  "TO20 Office of Content Solutions",
  "TO23 GEOINT Services",
  "TO30 GEOINT Enterprise",
  "TO31 Open IT Solutions",
  "TO32 IC Enterprise Management",
  "TO34 ITEMS and IC ITE",
  "TO43 OVI",
  "TO48 IPF"
]
    },
    ngamogave: {
      label: "NGA MOJAVE",
      taskOrders: ["TO02 RFO", "TO04 ATP", "TO12 SI"]
    },
    nganse: {
      label: "NGA NSE",
      taskOrders: ["TO05 IPA"]
    },
    npocolossus: {
      label: "NGA COLOSSUS",
    },
    nctctpi: {
      label: "NCTC TPI"
    },
    ngawstamp: {
      label: "NNGA WSTAMP"
    },
    ngartss: {
      label: "NGA RTSS",
    },
    ngas3: {
      label: "NGA S3",  
  },
  usaailmlfacialrecognition: {
    label: "USA AI/ML Facial Recognition",
  },
  nrolandmarkaos: {
    label: "NRO Landmark AOS",
  },
  nroispo: {
    label: "NRO ISPO",
  },
    ngaarso : {
      label: "NGA ARSO",
    }
  }
  };