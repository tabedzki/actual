// @ts-strict-ignore
import { send } from 'loot-core/platform/client/connection';
import * as monthUtils from 'loot-core/shared/months';
import type {
  CategoryEntity,
  RuleConditionEntity,
} from 'loot-core/types/models';

import type { useSpreadsheet } from '@desktop-client/hooks/useSpreadsheet';

type BudgetAnalysisIntervalData = {
  date: string;
  budgeted: number;
  spent: number;
  balance: number;
  overspendingAdjustment: number;
};

type BudgetAnalysisData = {
  intervalData: BudgetAnalysisIntervalData[];
  startDate: string;
  endDate: string;
  totalBudgeted: number;
  totalSpent: number;
  totalOverspendingAdjustment: number;
  finalOverspendingAdjustment: number;
};

type createBudgetAnalysisSpreadsheetProps = {
  conditions?: RuleConditionEntity[];
  conditionsOp?: 'and' | 'or';
  startDate: string;
  endDate: string;
};

export function createBudgetAnalysisSpreadsheet({
  conditions = [],
  conditionsOp = 'and',
  startDate,
  endDate,
}: createBudgetAnalysisSpreadsheetProps) {
  return async (
    spreadsheet: ReturnType<typeof useSpreadsheet>,
    setData: (data: BudgetAnalysisData) => void,
  ) => {
    const categoryScopeConditions = conditions.filter(
      condition =>
        !condition.customName &&
        (condition.field === 'category' ||
          condition.field === 'category_group'),
    );

    // Get all categories
    const { list: allCategories } = await send('get-categories');

    // Budget Analysis only supports category-scoped filters.
    const { categoryIds } = await send('budget/conditions-to-category-ids', {
      conditions: categoryScopeConditions,
      conditionsOp,
    });

    // Base set: expense categories only (exclude income and hidden)
    const baseCategories = allCategories.filter(
      (cat: CategoryEntity) => !cat.is_income && !cat.hidden,
    );

    // null means "no filter was applied" → use all expense categories
    const categoriesToInclude: CategoryEntity[] =
      categoryIds === null
        ? baseCategories
        : baseCategories.filter(cat => categoryIds.includes(cat.id));

    // Get monthly intervals (Budget Analysis only supports monthly)
    const intervals = monthUtils.rangeInclusive(
      monthUtils.getMonth(startDate),
      monthUtils.getMonth(endDate),
    );

    const intervalData: BudgetAnalysisIntervalData[] = [];

    // Track running balance that respects carryover flags
    // Get the balance from the month before the start period to initialize properly
    let runningBalance = 0;
    const monthBeforeStart = monthUtils.subMonths(
      monthUtils.getMonth(startDate),
      1,
    );
    const prevMonthData = await send('envelope-budget-month', {
      month: monthBeforeStart,
    });

    // Calculate the carryover from the previous month
    for (const cat of categoriesToInclude) {
      const balanceCell = prevMonthData.find((cell: { name: string }) =>
        cell.name.endsWith(`leftover-${cat.id}`),
      );
      const carryoverCell = prevMonthData.find((cell: { name: string }) =>
        cell.name.endsWith(`carryover-${cat.id}`),
      );

      const catBalance = (balanceCell?.value as number) || 0;
      const hasCarryover = Boolean(carryoverCell?.value);

      // Add to running balance if it would carry over
      if (catBalance > 0 || (catBalance < 0 && hasCarryover)) {
        runningBalance += catBalance;
      }
    }

    // Track totals across all months
    let totalBudgeted = 0;
    let totalSpent = 0;
    let totalOverspendingAdjustment = 0;

    // Track overspending from previous month to apply in next month
    let overspendingFromPrevMonth = 0;

    // Process each month
    for (const month of intervals) {
      // Get budget values from the server for this month
      // This uses the same calculations as the budget page
      const monthData = await send('envelope-budget-month', { month });

      let budgeted = 0;
      let spent = 0;
      let overspendingThisMonth = 0;

      // Track what will carry over to next month
      let carryoverToNextMonth = 0;

      // Sum up values for categories we're interested in
      for (const cat of categoriesToInclude) {
        // Find the budget, spent, balance, and carryover flag for this category
        const budgetCell = monthData.find((cell: { name: string }) =>
          cell.name.endsWith(`budget-${cat.id}`),
        );
        const spentCell = monthData.find((cell: { name: string }) =>
          cell.name.endsWith(`sum-amount-${cat.id}`),
        );
        const balanceCell = monthData.find((cell: { name: string }) =>
          cell.name.endsWith(`leftover-${cat.id}`),
        );
        const carryoverCell = monthData.find((cell: { name: string }) =>
          cell.name.endsWith(`carryover-${cat.id}`),
        );

        const catBudgeted = (budgetCell?.value as number) || 0;
        const catSpent = (spentCell?.value as number) || 0;
        const catBalance = (balanceCell?.value as number) || 0;
        const hasCarryover = Boolean(carryoverCell?.value);

        budgeted += catBudgeted;
        spent += catSpent;

        // Add to next month's carryover if:
        // - Balance is positive (always carries over), OR
        // - Balance is negative AND carryover is enabled
        if (catBalance > 0 || (catBalance < 0 && hasCarryover)) {
          carryoverToNextMonth += catBalance;
        } else if (catBalance < 0 && !hasCarryover) {
          // If balance is negative and carryover is NOT enabled,
          // this will be zeroed out and becomes next month's overspending adjustment
          overspendingThisMonth += catBalance; // Keep as negative
        }
      }

      // Apply overspending adjustment from previous month (negative value)
      const overspendingAdjustment = overspendingFromPrevMonth;

      // This month's balance = budgeted + spent + running balance + overspending adjustment
      const monthBalance = budgeted + spent + runningBalance;

      // Update totals
      totalBudgeted += budgeted;
      totalSpent += spent;
      totalOverspendingAdjustment += Math.abs(overspendingAdjustment);

      intervalData.push({
        date: month,
        budgeted,
        spent, // Display as positive
        balance: monthBalance,
        overspendingAdjustment: Math.abs(overspendingAdjustment), // Display as positive
      });

      // Update running balance for next month
      runningBalance = carryoverToNextMonth;
      // Save this month's overspending to apply in next month
      overspendingFromPrevMonth = overspendingThisMonth;
    }

    setData({
      intervalData,
      startDate,
      endDate,
      totalBudgeted,
      totalSpent,
      totalOverspendingAdjustment,
      finalOverspendingAdjustment: overspendingFromPrevMonth,
    });
  };
}
