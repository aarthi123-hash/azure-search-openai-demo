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
    war: {
      label: "War",
      taskOrders: ["1", "2"]
    },
    exile: {
      label: "Exile",
      taskOrders: ["1"]
    },
    mitre: {
      label: "Mitre"
      // No taskOrders - this program has no task order options
    },
    rfp: {
      label: "program RFP",
       taskOrders: ["1",'2', '3']
    }
  }
};