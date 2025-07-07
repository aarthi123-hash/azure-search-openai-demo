// filterUtils.ts
import { filterConfig, IDropdownOption } from './filterConfig';

// Helper function to get dropdown options for programs
export const getProgramOptions = (): IDropdownOption[] => {
  return Object.keys(filterConfig.programs).map(key => ({
    key,
    text: filterConfig.programs[key].label
  }));
};

// Helper function to get task order options for a specific program
export const getTaskOrderOptions = (programKey: string): IDropdownOption[] => {
  const program = filterConfig.programs[programKey];
  if (!program || !program.taskOrders) {
    return [];
  }
  
  return program.taskOrders.map(taskOrder => ({
    key: taskOrder,
    text: `Task Order ${taskOrder}`
  }));
};

// Helper function to check if a program has task orders
export const hasTaskOrders = (programKey: string): boolean => {
  const program = filterConfig.programs[programKey];
  return !!(program && program.taskOrders && program.taskOrders.length > 0);
};