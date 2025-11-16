import * as d from 'date-fns';

import { send } from 'loot-core/platform/client/fetch';
import * as monthUtils from 'loot-core/shared/months';
import { q } from 'loot-core/shared/query';
import {
  type AccountEntity,
  type PayeeEntity,
  type CategoryEntity,
  type RuleConditionEntity,
  type CategoryGroupEntity,
  type balanceTypeOpType,
  type sortByOpType,
  type DataEntity,
  type GroupedEntity,
  type IntervalEntity,
} from 'loot-core/types/models';
import { type SyncedPrefs } from 'loot-core/types/prefs';

import { calculateLegend } from './calculateLegend';
import { filterEmptyRows } from './filterEmptyRows';
import { filterHiddenItems } from './filterHiddenItems';
import { makeQuery } from './makeQuery';
import { recalculate } from './recalculate';
import { sortData } from './sortData';
import {
  determineIntervalRange,
  trimIntervalDataToRange,
  trimIntervalsToRange,
} from './trimIntervals';

import {
  categoryLists,
  groupBySelections,
  type QueryDataEntity,
  ReportOptions,
  type UncategorizedEntity,
} from '@desktop-client/components/reports/ReportOptions';
import { type useSpreadsheet } from '@desktop-client/hooks/useSpreadsheet';
import { aqlQuery } from '@desktop-client/queries/aqlQuery';

export type createCustomSpreadsheetProps = {
  startDate: string;
  endDate: string;
  interval: string;
  categories: { list: CategoryEntity[]; grouped: CategoryGroupEntity[] };
  conditions: RuleConditionEntity[];
  conditionsOp: string;
  showEmpty: boolean;
  showOffBudget: boolean;
  showHiddenCategories: boolean;
  showUncategorized: boolean;
  trimIntervals: boolean;
  groupBy?: string;
  balanceTypeOp?: balanceTypeOpType;
  sortByOp?: sortByOpType;
  payees?: PayeeEntity[];
  accounts?: AccountEntity[];
  graphType?: string;
  firstDayOfWeekIdx?: SyncedPrefs['firstDayOfWeekIdx'];
  setDataCheck?: (value: boolean) => void;
};

export function createCustomSpreadsheet({
  startDate,
  endDate,
  interval,
  categories,
  conditions = [],
  conditionsOp,
  showEmpty,
  showOffBudget,
  showHiddenCategories,
  showUncategorized,
  trimIntervals,
  groupBy = '',
  balanceTypeOp = 'totalDebts',
  sortByOp = 'desc',
  payees = [],
  accounts = [],
  graphType,
  firstDayOfWeekIdx,
  setDataCheck,
}: createCustomSpreadsheetProps) {
  const [categoryList, categoryGroup] = categoryLists(categories);

  const [groupByList, groupByLabel]: [
    groupByList: UncategorizedEntity[],
    groupByLabel: 'category' | 'categoryGroup' | 'payee' | 'account',
  ] = groupBySelections(groupBy, categoryList, categoryGroup, payees, accounts);

  return async (
    spreadsheet: ReturnType<typeof useSpreadsheet>,
    setData: (data: DataEntity) => void,
  ) => {
    if (groupByList.length === 0) {
      return;
    }

    const { filters } = await send('make-filters-from-conditions', {
      conditions: conditions.filter(cond => !cond.customName),
    });
    const conditionsOpKey = conditionsOp === 'or' ? '$or' : '$and';

    // Prepare budget filters (category-only filters for budget queries)
    const { filters: budgetFilters } = await send(
      'make-filters-from-conditions',
      {
        conditions: conditions.filter(
          cond => !cond.customName && cond.field === 'category',
        ),
        applySpecialCases: false,
      },
    );

    let assets: QueryDataEntity[];
    let debts: QueryDataEntity[];
    [assets, debts] = await Promise.all([
      aqlQuery(
        makeQuery(
          'assets',
          startDate,
          endDate,
          interval,
          conditionsOpKey,
          filters,
        ),
      ).then(({ data }) => data),
      aqlQuery(
        makeQuery(
          'debts',
          startDate,
          endDate,
          interval,
          conditionsOpKey,
          filters,
        ),
      ).then(({ data }) => data),
    ]);

    if (interval === 'Weekly') {
      debts = debts.map(d => {
        return {
          ...d,
          date: monthUtils.weekFromDate(d.date, firstDayOfWeekIdx),
        };
      });
      assets = assets.map(d => {
        return {
          ...d,
          date: monthUtils.weekFromDate(d.date, firstDayOfWeekIdx),
        };
      });
    }

    const intervals =
      interval === 'Weekly'
        ? monthUtils.weekRangeInclusive(startDate, endDate, firstDayOfWeekIdx)
        : monthUtils[
            ReportOptions.intervalRange.get(interval) || 'rangeInclusive'
          ](startDate, endDate);

    // Fetch budget data for budget-related balance types
    // Budget data is only meaningful for category-based grouping
    const needsBudgetData =
      balanceTypeOp === 'budgeted' || balanceTypeOp === 'budgetBalance';
    const isCategoryBased =
      groupByLabel === 'category' || groupByLabel === 'categoryGroup';

    type BudgetDataEntity = {
      month: number;
      category: string;
      amount: number;
    };

    let budgetData: BudgetDataEntity[] = [];
    if (needsBudgetData && isCategoryBased) {
      // Get all months covered by the intervals
      const monthIntervals =
        interval === 'Monthly'
          ? intervals
          : interval === 'Yearly'
            ? intervals.map(y => `${y}-01`)
            : monthUtils.rangeInclusive(
                monthUtils.monthFromDate(d.parseISO(intervals[0])),
                monthUtils.monthFromDate(
                  d.parseISO(intervals[intervals.length - 1]),
                ),
              );

      // Fetch budget data for all months
      const monthNumbers = monthIntervals.map(m =>
        parseInt(m.replace('-', '')),
      );

      budgetData = await aqlQuery(
        q('zero_budgets')
          .filter({
            month: { $oneof: monthNumbers },
          })
          .filter({
            [conditionsOpKey]: budgetFilters,
          })
          .select(['month', 'category', 'amount']),
      ).then(({ data }) => data);
    }

    let totalAssets = 0;
    let totalDebts = 0;
    let netAssets = 0;
    let netDebts = 0;
    let totalBudgeted = 0;
    let totalBudgetBalance = 0;

    const groupsByCategory =
      groupByLabel === 'category' || groupByLabel === 'categoryGroup';

    const intervalData = intervals.reduce(
      (arr: IntervalEntity[], intervalItem, index) => {
        let perIntervalAssets = 0;
        let perIntervalDebts = 0;
        let perIntervalNetAssets = 0;
        let perIntervalNetDebts = 0;
        let perIntervalTotals = 0;
        let perIntervalBudgeted = 0;
        let perIntervalBudgetBalance = 0;
        const stacked: Record<string, number> = {};

        groupByList.map(item => {
          let stackAmounts = 0;

          const intervalAssets = filterHiddenItems(
            item,
            assets,
            showOffBudget,
            showHiddenCategories,
            showUncategorized,
            groupsByCategory,
          )
            .filter(
              asset =>
                asset.date === intervalItem &&
                (asset[groupByLabel] === (item.id ?? null) ||
                  (item.uncategorized_id && groupsByCategory)),
            )
            .reduce((a, v) => (a = a + v.amount), 0);
          perIntervalAssets += intervalAssets;

          const intervalDebts = filterHiddenItems(
            item,
            debts,
            showOffBudget,
            showHiddenCategories,
            showUncategorized,
            groupsByCategory,
          )
            .filter(
              debt =>
                debt.date === intervalItem &&
                (debt[groupByLabel] === (item.id ?? null) ||
                  (item.uncategorized_id && groupsByCategory)),
            )
            .reduce((a, v) => (a = a + v.amount), 0);
          perIntervalDebts += intervalDebts;

          const netAmounts = intervalAssets + intervalDebts;

          // Calculate budget amounts for this item in this interval
          let budgetAmount = 0;
          if (needsBudgetData && isCategoryBased && item.id) {
            // Determine which month(s) this interval covers
            let intervalMonths: number[] = [];
            if (interval === 'Monthly') {
              intervalMonths = [parseInt(intervalItem.replace('-', ''))];
            } else if (interval === 'Yearly') {
              // For yearly, sum all months in that year
              const year = parseInt(intervalItem);
              intervalMonths = Array.from(
                { length: 12 },
                (_, i) => year * 100 + i + 1,
              );
            } else {
              // For Daily/Weekly, get the month of this interval
              const month = monthUtils.monthFromDate(d.parseISO(intervalItem));
              intervalMonths = [parseInt(month.replace('-', ''))];
            }

            // Sum budget for this category/group in these months
            budgetAmount = budgetData
              .filter(
                b =>
                  intervalMonths.includes(b.month) &&
                  (groupByLabel === 'category'
                    ? b.category === item.id
                    : // For groups, we need to check if the category belongs to this group
                      categories.list.some(
                        cat => cat.id === b.category && cat.group === item.id,
                      )),
              )
              .reduce((sum, b) => sum + b.amount, 0);
          }

          const budgetBalanceAmount = budgetAmount - Math.abs(intervalDebts);

          if (balanceTypeOp === 'totalAssets') {
            stackAmounts += intervalAssets;
          }
          if (balanceTypeOp === 'totalDebts') {
            stackAmounts += Math.abs(intervalDebts);
          }
          if (balanceTypeOp === 'netAssets') {
            stackAmounts += netAmounts > 0 ? netAmounts : 0;
          }
          if (balanceTypeOp === 'netDebts') {
            stackAmounts = netAmounts < 0 ? Math.abs(netAmounts) : 0;
          }
          if (balanceTypeOp === 'totalTotals') {
            stackAmounts += netAmounts;
          }
          if (balanceTypeOp === 'budgeted') {
            stackAmounts += Math.abs(budgetAmount);
          }
          if (balanceTypeOp === 'budgetBalance') {
            stackAmounts += budgetBalanceAmount;
          }

          stacked[item.name] = stackAmounts;

          perIntervalBudgeted += Math.abs(budgetAmount);
          perIntervalBudgetBalance += budgetBalanceAmount;

          perIntervalNetAssets =
            netAmounts > 0
              ? perIntervalNetAssets + netAmounts
              : perIntervalNetAssets;
          perIntervalNetDebts =
            netAmounts < 0
              ? perIntervalNetDebts + netAmounts
              : perIntervalNetDebts;
          perIntervalTotals += netAmounts;

          return null;
        });
        totalAssets += perIntervalAssets;
        totalDebts += perIntervalDebts;
        netAssets += perIntervalNetAssets;
        netDebts += perIntervalNetDebts;
        totalBudgeted += perIntervalBudgeted;
        totalBudgetBalance += perIntervalBudgetBalance;

        arr.push({
          date: d.format(
            d.parseISO(intervalItem),
            ReportOptions.intervalFormat.get(interval) || '',
          ),
          ...stacked,
          intervalStartDate: index === 0 ? startDate : intervalItem,
          intervalEndDate:
            index + 1 === intervals.length
              ? endDate
              : monthUtils.subDays(intervals[index + 1], 1),
          totalAssets: perIntervalAssets,
          totalDebts: perIntervalDebts,
          netAssets: perIntervalNetAssets,
          netDebts: perIntervalNetDebts,
          totalTotals: perIntervalTotals,
          budgeted: needsBudgetData ? perIntervalBudgeted : undefined,
          budgetBalance: needsBudgetData ? perIntervalBudgetBalance : undefined,
        });

        return arr;
      },
      [],
    );

    const calcData: GroupedEntity[] = groupByList.map(item => {
      const calc = recalculate({
        item,
        intervals,
        assets,
        debts,
        groupByLabel,
        showOffBudget,
        showHiddenCategories,
        showUncategorized,
        startDate,
        endDate,
      });
      return { ...calc };
    });

    // First, filter rows so trimming reflects the visible dataset
    const calcDataFiltered = calcData.filter(i =>
      filterEmptyRows({ showEmpty, data: i, balanceTypeOp }),
    );

    // Determine interval range across filtered groups and main intervalData
    const { startIndex, endIndex } = determineIntervalRange(
      calcDataFiltered,
      intervalData,
      trimIntervals,
      balanceTypeOp,
    );

    // Trim only if enabled
    const trimmedIntervalData = trimIntervals
      ? trimIntervalDataToRange(intervalData, startIndex, endIndex)
      : intervalData;

    if (trimIntervals) {
      // Keep group data in sync with the trimmed range
      trimIntervalsToRange(calcDataFiltered, startIndex, endIndex);
    }

    const sortedCalcDataFiltered = [...calcDataFiltered].sort(
      sortData({ balanceTypeOp, sortByOp }),
    );

    const legend = calculateLegend(
      trimmedIntervalData,
      sortedCalcDataFiltered,
      groupBy,
      graphType,
      balanceTypeOp,
    );

    setData({
      data: sortedCalcDataFiltered,
      intervalData: trimmedIntervalData,
      legend,
      startDate,
      endDate,
      totalAssets,
      totalDebts,
      netAssets,
      netDebts,
      totalTotals: totalAssets + totalDebts,
      budgeted: needsBudgetData ? totalBudgeted : undefined,
      budgetBalance: needsBudgetData ? totalBudgetBalance : undefined,
    });
    setDataCheck?.(true);
  };
}
