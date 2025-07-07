// filterUtils.ts
import { filterConfig, IDropdownOption, FilterConfig } from './filterConfig';

// Helper function to get dropdown options for programs
export const getProgramOptions = (config: FilterConfig = filterConfig): IDropdownOption[] => {
  return Object.keys(config.programs).map(key => ({
    key,
    text: config.programs[key].label
  }));
};

// Helper function to get task order options for a specific program
export const getTaskOrderOptions = (programKey: string, config: FilterConfig = filterConfig): IDropdownOption[] => {
  const program = config.programs[programKey];
  if (!program || !program.taskOrders) {
    return [];
  }
  
  return program.taskOrders.map(taskOrder => ({
    key: taskOrder,
    text: `Task Order ${taskOrder}`
  }));
};

// Helper function to check if a program has task orders
export const hasTaskOrders = (programKey: string, config: FilterConfig = filterConfig): boolean => {
  const program = config.programs[programKey];
  return !!(program && program.taskOrders && program.taskOrders.length > 0);
};