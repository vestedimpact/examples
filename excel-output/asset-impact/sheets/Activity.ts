import { Workbook } from 'exceljs';

import { Asset, AssetImpact } from '../../../api/types';
import { SDGUtils } from '../../../utils/sdg';
import { OutputWorksheet } from '../../common/worksheet';

export class ActivityDataSheet extends OutputWorksheet {
  constructor(
    private readonly impact: AssetImpact['impactBreakdown'][0],
    private readonly date: string,
    private readonly logoId: number,
    private readonly asset: Asset,
    workbook: Workbook,
    index: number,
  ) {
    super(`Activity ${index + 1} Data`, workbook);
  }

  private addHeader() {
    this.addBlankRows(7, 1);
    this.addImage(this.logoId, {
      tl: { col: 1, row: 1 },
      ext: { width: 500, height: 50 },
    });
    this.getRow(2).height = 50;
    this.addText('B3', 'Asset Name', { bold: true });
    this.addText('B4', 'Industry', { bold: true });
    this.addText('B5', 'Assessment date', { bold: true });
    this.getColumn('B').width = 25;
    this.addText('C3', this.asset.name, { wrapText: false });
    this.addText('C4', this.asset.industry || '', { wrapText: false });
    this.addText('C5', this.date, { wrapText: false });
    this.getColumn('C').width = 25;
    this.getRow(6).height = 5;
    this.getRow(6).border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
  }

  private activitySummaryTable() {
    this.addText('B11', 'Business Activity');
    this.addBorders('B11');
    this.setColor('B11', 'FFCCCCCC');
    this.addText('C11', 'Proportion of Organization');
    this.addBorders('C11');
    this.setColor('C11', 'FFCCCCCC');
    this.addText('D11', 'Positive Impact');
    this.addBorders('D11');
    this.setColor('D11', 'FFCCCCCC');
    this.addText('E11', 'Negative Impact');
    this.addBorders('E11');
    this.setColor('E11', 'FFCCCCCC');
    this.addText('B12', this.impact.activity);
    this.addBorders('B12');
    this.addNumber('C12', this.impact.weight, { numFmt: '0.00%' });
    this.addBorders('C12');
    if (this.impact.positiveImpact !== 0) {
      this.addNumber('D12', this.impact.positiveImpact, { numFmt: '#' });
    } else {
      this.addText('D12', 'None');
    }
    this.addBorders('D12');
    if (this.impact.negativeImpact !== 0) {
      this.addNumber('E12', this.impact.negativeImpact, { numFmt: '#' });
    } else {
      this.addText('E12', 'None');
    }
    this.addBorders('E12');
  }

  private benchmarksTable(startIndex: number) {
    this.addText(`B${startIndex}`, `Goal-based Benchmarks for ${this.impact.activity} Activity`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    if (this.impact.benchmarks.length > 0) {
      this.addText(`B${startIndex + 1}`, 'SDG Target');
      this.addBorders(`B${startIndex + 1}`);
      this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
      this.addText(`C${startIndex + 1}`, 'Country');
      this.addBorders(`C${startIndex + 1}`);
      this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
      this.merge(`D${startIndex + 1}:F${startIndex + 1}`);
      this.addText(`D${startIndex + 1}`, 'Indicator Used');
      this.addBorders(`D${startIndex + 1}`);
      this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
      this.merge(`G${startIndex + 1}:J${startIndex + 1}`);
      this.addText(`G${startIndex + 1}`, 'Annual Required % Change To Meet SDG');
      this.addBorders(`G${startIndex + 1}`);
      this.setColor(`G${startIndex + 1}`, 'FFCCCCCC');
      this.merge(`K${startIndex + 1}:M${startIndex + 1}`);
      this.addText(`K${startIndex + 1}`, 'Organization Pace Of Change');
      this.addBorders(`K${startIndex + 1}`);
      this.setColor(`K${startIndex + 1}`, 'FFCCCCCC');
      let benchmarkIndex = startIndex + 2;
      this.impact.benchmarks.forEach((benchmark) => {
        benchmark.countries.forEach((country) => {
          this.addText(`B${benchmarkIndex}`, benchmark.sdgTarget);
          this.addBorders(`B${benchmarkIndex}`);
          this.addText(`C${benchmarkIndex}`, country.country);
          this.addBorders(`C${benchmarkIndex}`);
          this.merge(`D${benchmarkIndex}:F${benchmarkIndex}`);
          this.addText(`D${benchmarkIndex}`, benchmark.indicator);
          this.addBorders(`D${benchmarkIndex}`);
          this.merge(`G${benchmarkIndex}:J${benchmarkIndex}`);
          this.addNumber(`G${benchmarkIndex}`, country.paceOfChange, { numFmt: '0.00%' });
          this.addBorders(`G${benchmarkIndex}`);
          this.merge(`K${benchmarkIndex}:M${benchmarkIndex}`);
          this.addNumber(`K${benchmarkIndex}`, country.assetGrowth / 100, { numFmt: '0.00%' });
          this.addBorders(`K${benchmarkIndex}`);
          benchmarkIndex += 1;
        });
      });
      return benchmarkIndex + 1;
    } else {
      this.merge(`B${startIndex + 1}:M${startIndex + 1}`);
      this.addText(
        `B${startIndex + 1}`,
        'No goal-based benchmarks have been identified at the time of the assessment',
        { italic: true },
      );
      return startIndex + 3;
    }
  }

  private countrySummaryTable() {
    this.addText('B14', `Countries Impacted by ${this.impact.activity} Activity`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    this.addText('B15', 'Country');
    this.addBorders('B15');
    this.setColor('B15', 'FFCCCCCC');
    this.addText('C15', 'Proportion of Activity');
    this.addBorders('C15');
    this.setColor('C15', 'FFCCCCCC');
    this.addText('D15', 'Positive Impact');
    this.addBorders('D15');
    this.setColor('D15', 'FFCCCCCC');
    this.addText('E15', 'Negative Impact');
    this.addBorders('E15');
    this.setColor('E15', 'FFCCCCCC');
    this.impact.countries.forEach((country, index) => {
      this.addText(`B${16 + index}`, country.country);
      this.addBorders(`B${16 + index}`);
      this.addNumber(`C${16 + index}`, country.weight / this.impact.weight, { numFmt: '0.00%' });
      this.addBorders(`C${16 + index}`);
      if (country.positiveImpact !== 0) {
        this.addNumber(`D${16 + index}`, country.positiveImpact, { numFmt: '#' });
      } else {
        this.addText(`D${16 + index}`, 'None');
      }
      this.addBorders(`D${16 + index}`);
      if (country.negativeImpact !== 0) {
        this.addNumber(`E${16 + index}`, country.negativeImpact, { numFmt: '#' });
      } else {
        this.addText(`E${16 + index}`, 'None');
      }
      this.addBorders(`E${16 + index}`);
    });
    return 17 + this.impact.countries.length;
  }

  private flagsTables(startIndex: number) {
    const hasFlaggedOpportunities = this.impact.flaggedOpportunities.length > 0;
    const hasFlaggedRisks = this.impact.flaggedRisks.length > 0;
    let opportunitiesStart = startIndex + 3;
    this.addText(`B${startIndex}`, `Flagged Risks for ${this.impact.activity} Activity`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    if (hasFlaggedRisks) {
      this.addText(`B${startIndex + 1}`, 'Flag Type');
      this.addBorders(`B${startIndex + 1}`);
      this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
      this.addText(`C${startIndex + 1}`, 'Status');
      this.addBorders(`C${startIndex + 1}`);
      this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
      this.addText(`D${startIndex + 1}`, 'Countries');
      this.addBorders(`D${startIndex + 1}`);
      this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
      this.addText(`E${startIndex + 1}`, 'SDG Targets');
      this.addBorders(`E${startIndex + 1}`);
      this.setColor(`E${startIndex + 1}`, 'FFCCCCCC');
      this.addText(`F${startIndex + 1}`, 'Note');
      this.addBorders(`F${startIndex + 1}`);
      this.setColor(`F${startIndex + 1}`, 'FFCCCCCC');
      this.impact.flaggedRisks.forEach((flag, index) => {
        this.addText(`B${startIndex + 2 + index}`, flag.type);
        this.addBorders(`B${startIndex + 2 + index}`);
        this.addText(`C${startIndex + 2 + index}`, flag.status);
        this.addBorders(`C${startIndex + 2 + index}`);
        this.addText(`D${startIndex + 2 + index}`, flag.targets.map((i) => i.country).join(','));
        this.addBorders(`D${startIndex + 2 + index}`);
        this.addText(`E${startIndex + 2 + index}`, flag.targets.map((i) => i.sdgTarget).join(','));
        this.addBorders(`E${startIndex + 2 + index}`);
        this.addText(`F${startIndex + 2 + index}`, flag.note);
        this.addBorders(`F${startIndex + 2 + index}`);
      });
      opportunitiesStart = startIndex + 3 + this.impact.flaggedRisks.length;
    } else {
      this.merge(`B${startIndex + 1}:R${startIndex + 1}`);
      this.addText(`B${startIndex + 1}`, 'No flagged risks have been identified at the time of the assessment', {
        italic: true,
      });
    }
    this.addText(`B${opportunitiesStart}`, `Flagged Opportunities for ${this.impact.activity} Activity`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    if (hasFlaggedOpportunities) {
      this.addText(`B${opportunitiesStart + 1}`, 'Flag Type');
      this.addBorders(`B${opportunitiesStart + 1}`);
      this.setColor(`B${opportunitiesStart + 1}`, 'FFCCCCCC');
      this.addText(`C${opportunitiesStart + 1}`, 'Status');
      this.addBorders(`C${opportunitiesStart + 1}`);
      this.setColor(`C${opportunitiesStart + 1}`, 'FFCCCCCC');
      this.addText(`D${opportunitiesStart + 1}`, 'Country');
      this.addBorders(`D${opportunitiesStart + 1}`);
      this.setColor(`D${opportunitiesStart + 1}`, 'FFCCCCCC');
      this.addText(`E${opportunitiesStart + 1}`, 'SDG Target');
      this.addBorders(`E${opportunitiesStart + 1}`);
      this.setColor(`E${opportunitiesStart + 1}`, 'FFCCCCCC');
      this.addText(`F${opportunitiesStart + 1}`, 'Note');
      this.addBorders(`F${opportunitiesStart + 1}`);
      this.setColor(`F${opportunitiesStart + 1}`, 'FFCCCCCC');
      this.impact.flaggedOpportunities.forEach((flag, index) => {
        this.addText(`B${opportunitiesStart + 2 + index}`, flag.type);
        this.addBorders(`B${opportunitiesStart + 2 + index}`);
        this.addText(`C${opportunitiesStart + 2 + index}`, flag.status);
        this.addBorders(`C${opportunitiesStart + 2 + index}`);
        this.addText(`D${opportunitiesStart + 2 + index}`, flag.targets.map((i) => i.country).join(','));
        this.addBorders(`D${opportunitiesStart + 2 + index}`);
        this.addText(`E${opportunitiesStart + 2 + index}`, flag.targets.map((i) => i.sdgTarget).join(','));
        this.addBorders(`E${opportunitiesStart + 2 + index}`);
        this.addText(`F${opportunitiesStart + 2 + index}`, flag.note);
        this.addBorders(`F${opportunitiesStart + 2 + index}`);
      });
      return opportunitiesStart + 3 + this.impact.flaggedOpportunities.length;
    } else {
      this.merge(`B${opportunitiesStart + 1}:R${opportunitiesStart + 1}`);
      this.addText(
        `B${opportunitiesStart + 1}`,
        'No flagged opportunities have been identified at the time of the assessment',
        { italic: true },
      );
      return opportunitiesStart + 3;
    }
  }

  private indicatorsTable(startIndex: number) {
    this.addText(`B${startIndex}`, `Indicators Used for Assessment of ${this.impact.activity} Activity`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    this.addText(`B${startIndex + 1}`, 'SDG Target');
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 1}`, 'Country');
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`D${startIndex + 1}:F${startIndex + 1}`);
    this.addText(`D${startIndex + 1}`, 'Indicator');
    this.addBorders(`D${startIndex + 1}`);
    this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`G${startIndex + 1}:J${startIndex + 1}`);
    this.addText(`G${startIndex + 1}`, 'Source');
    this.addBorders(`G${startIndex + 1}`);
    this.setColor(`G${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`K${startIndex + 1}`, 'Trend');
    this.addBorders(`K${startIndex + 1}`);
    this.setColor(`K${startIndex + 1}`, 'FFCCCCCC');
    this.impact.indicators.forEach((indicator, index) => {
      this.addText(`B${startIndex + 2 + index}`, indicator.sdgTarget);
      this.addBorders(`B${startIndex + 2 + index}`);
      this.addText(`C${startIndex + 2 + index}`, indicator.country);
      this.addBorders(`C${startIndex + 2 + index}`);
      this.merge(`D${startIndex + 2 + index}:F${startIndex + 2 + index}`);
      this.addText(`D${startIndex + 2 + index}`, indicator.indicator);
      this.addBorders(`D${startIndex + 2 + index}`);
      this.merge(`G${startIndex + 2 + index}:J${startIndex + 2 + index}`);
      this.addText(`G${startIndex + 2 + index}`, indicator.source);
      this.addBorders(`G${startIndex + 2 + index}`);
      this.addText(`K${startIndex + 2 + index}`, indicator.trend ? indicator.trend.toFixed(3) : 'Not available');
      this.addBorders(`K${startIndex + 2 + index}`);
    });
    return startIndex + 3 + this.impact.indicators.length;
  }

  private metricsTable(startIndex: number) {
    this.addText(`B${startIndex}`, `Environmental Metrics for ${this.impact.activity} Activity`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    this.addText(`B${startIndex + 1}`, 'Metric', { wrapText: false });
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`C${startIndex + 1}:F${startIndex + 1}`);
    this.addText(`C${startIndex + 1}`, 'Description', { wrapText: false });
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`G${startIndex + 1}:I${startIndex + 1}`);
    this.addText(`G${startIndex + 1}`, 'Category', { wrapText: false });
    this.addBorders(`G${startIndex + 1}`);
    this.setColor(`G${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`J${startIndex + 1}:L${startIndex + 1}`);
    this.addText(`J${startIndex + 1}`, 'Overall Value', { wrapText: false });
    this.addBorders(`J${startIndex + 1}`);
    this.setColor(`J${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`M${startIndex + 1}:O${startIndex + 1}`);
    this.addText(`M${startIndex + 1}`, 'Units', { wrapText: false });
    this.addBorders(`M${startIndex + 1}`);
    this.setColor(`M${startIndex + 1}`, 'FFCCCCCC');
    let metricIndex = startIndex + 2;
    const metricHeights = [24, 36, 36, 36, 48, 48, 48, 48, 36, 24, 15, 15, 15, 15, 15, 24, 24, 24];
    this.impact.metrics.forEach((metric, index) => {
      this.getRow(metricIndex).height = metricHeights[index];
      this.addText(`B${metricIndex}`, metric.name);
      this.addBorders(`B${metricIndex}`);
      this.merge(`C${metricIndex}:F${metricIndex}`);
      this.addText(`C${metricIndex}`, metric.description);
      this.addBorders(`C${metricIndex}`);
      this.merge(`G${metricIndex}:I${metricIndex}`);
      this.addText(`G${metricIndex}`, metric.category);
      this.addBorders(`G${metricIndex}`);
      this.merge(`J${metricIndex}:L${metricIndex}`);
      this.addNumber(`J${metricIndex}`, metric.value || 0);
      this.addBorders(`J${metricIndex}`);
      this.merge(`M${metricIndex}:O${metricIndex}`);
      this.addText(`M${metricIndex}`, metric.units);
      this.addBorders(`M${metricIndex}`);
      metricIndex += 1;
    });
    return metricIndex + 1;
  }

  private referencesTable(startIndex: number) {
    this.addText(`B${startIndex}`, `Academic References Supporting Assessment of ${this.impact.activity} Activity`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    this.addText(`B${startIndex + 1}`, 'SDG Targets');
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`C${startIndex + 1}:F${startIndex + 1}`);
    this.addText(`C${startIndex + 1}`, 'Reference');
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`G${startIndex + 1}:L${startIndex + 1}`);
    this.addText(`G${startIndex + 1}`, 'URL');
    this.addBorders(`G${startIndex + 1}`);
    this.setColor(`G${startIndex + 1}`, 'FFCCCCCC');
    this.merge(`M${startIndex + 1}:N${startIndex + 1}`);
    this.addText(`M${startIndex + 1}`, 'Standard Of Evidence');
    this.addBorders(`M${startIndex + 1}`);
    this.setColor(`M${startIndex + 1}`, 'FFCCCCCC');
    this.impact.references.forEach((reference, index) => {
      this.getRow(startIndex + 2 + index).height = 24;
      this.addText(`B${startIndex + 2 + index}`, reference.targets.map((i) => i.sdgTarget).join('\n'));
      this.addBorders(`B${startIndex + 2 + index}`);
      this.merge(`C${startIndex + 2 + index}:F${startIndex + 2 + index}`);
      this.addText(`C${startIndex + 2 + index}`, reference.reference);
      this.addBorders(`C${startIndex + 2 + index}`);
      this.merge(`G${startIndex + 2 + index}:L${startIndex + 2 + index}`);
      this.addText(`G${startIndex + 2 + index}`, reference.url || '');
      this.addBorders(`G${startIndex + 2 + index}`);
      this.merge(`M${startIndex + 2 + index}:N${startIndex + 2 + index}`);
      this.addText(`M${startIndex + 2 + index}`, reference.targets.map((i) => i.standardOfEvidence || '').join('\n'));
      this.addBorders(`M${startIndex + 2 + index}`);
    });
    return startIndex + 3 + this.impact.references.length;
  }

  private subScoresTable(startIndex: number) {
    this.merge(`B${startIndex}:F${startIndex}`);
    this.addText(`B${startIndex}`, `Assessment of ${this.impact.activity} Activity on Impacted Outcomes`, {
      bold: true,
      underline: true,
      wrapText: false,
    });
    this.merge(`H${startIndex}:K${startIndex}`);
    this.addText(`H${startIndex}`, 'Need Sub Scores', {
      align: { horizontal: 'center', vertical: 'middle' },
      bold: true,
    });
    this.addBorders(`H${startIndex}`);
    this.setColor(`H${startIndex}`, 'FFCCCCCC');
    this.merge(`M${startIndex}:O${startIndex}`);
    this.addText(`M${startIndex}`, 'Importance Sub Scores', {
      align: { horizontal: 'center', vertical: 'middle' },
      bold: true,
    });
    this.addBorders(`M${startIndex}`);
    this.setColor(`M${startIndex}`, 'FFCCCCCC');
    this.merge(`Q${startIndex}:T${startIndex}`);
    this.addText(`Q${startIndex}`, 'Value Sub Scores', {
      align: { horizontal: 'center', vertical: 'middle' },
      bold: true,
    });
    this.addBorders(`Q${startIndex}`);
    this.setColor(`Q${startIndex}`, 'FFCCCCCC');
    this.merge(`V${startIndex}:W${startIndex}`);
    this.addText(`V${startIndex}`, 'Contribution Sub Scores', {
      align: { horizontal: 'center', vertical: 'middle' },
      bold: true,
    });
    this.addBorders(`V${startIndex}`);
    this.setColor(`V${startIndex}`, 'FFCCCCCC');
    this.addText(`B${startIndex + 1}`, 'SDG Target');
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 1}`, 'Country');
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`D${startIndex + 1}`, 'Positive Impact');
    this.addBorders(`D${startIndex + 1}`);
    this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`E${startIndex + 1}`, 'Negative Impact');
    this.addBorders(`E${startIndex + 1}`);
    this.setColor(`E${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`F${startIndex + 1}`, 'Findings');
    this.addBorders(`F${startIndex + 1}`);
    this.setColor(`F${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`G${startIndex + 1}`, 'Need Score');
    this.addBorders(`G${startIndex + 1}`);
    this.setColor(`G${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`H${startIndex + 1}`, 'UN Classification');
    this.addBorders(`H${startIndex + 1}`);
    this.setColor(`H${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`I${startIndex + 1}`, 'World Bank Income Group');
    this.addBorders(`I${startIndex + 1}`);
    this.setColor(`I${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`J${startIndex + 1}`, 'SDG Status');
    this.addBorders(`J${startIndex + 1}`);
    this.setColor(`J${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`K${startIndex + 1}`, 'SDG Trend');
    this.addBorders(`K${startIndex + 1}`);
    this.setColor(`K${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`L${startIndex + 1}`, 'Importance Score');
    this.addBorders(`L${startIndex + 1}`);
    this.setColor(`L${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`M${startIndex + 1}`, 'Global Score');
    this.addBorders(`M${startIndex + 1}`);
    this.setColor(`M${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`N${startIndex + 1}`, 'Supporting Score');
    this.addBorders(`N${startIndex + 1}`);
    this.setColor(`N${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`O${startIndex + 1}`, 'Local Score');
    this.addBorders(`O${startIndex + 1}`);
    this.setColor(`O${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`P${startIndex + 1}`, 'Value Score');
    this.addBorders(`P${startIndex + 1}`);
    this.setColor(`P${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`Q${startIndex + 1}`, 'Depth Score');
    this.addBorders(`Q${startIndex + 1}`);
    this.setColor(`Q${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`R${startIndex + 1}`, 'Immediacy Score');
    this.addBorders(`R${startIndex + 1}`);
    this.setColor(`R${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`S${startIndex + 1}`, 'Sustained Score');
    this.addBorders(`S${startIndex + 1}`);
    this.setColor(`S${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`T${startIndex + 1}`, 'Irremediable Score');
    this.addBorders(`T${startIndex + 1}`);
    this.setColor(`T${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`U${startIndex + 1}`, 'Contribution Score');
    this.addBorders(`U${startIndex + 1}`);
    this.setColor(`U${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`V${startIndex + 1}`, 'Scale Score');
    this.addBorders(`V${startIndex + 1}`);
    this.setColor(`V${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`W${startIndex + 1}`, 'Change Score');
    this.addBorders(`W${startIndex + 1}`);
    this.setColor(`W${startIndex + 1}`, 'FFCCCCCC');
    let subScoreIndex = startIndex + 2;
    this.impact.targets.forEach((target) => {
      target.subScores.forEach((subScore) => {
        this.addText(`B${subScoreIndex}`, SDGUtils.getTargetLabel(target.sdgTarget));
        this.addBorders(`B${subScoreIndex}`);
        this.addText(`C${subScoreIndex}`, subScore.country);
        this.addBorders(`C${subScoreIndex}`);
        this.addText(
          `D${subScoreIndex}`,
          subScore.impactScore > 0 ? `${subScore.impactStatus} (${subScore.impactScore.toFixed(0)})` : 'None',
        );
        this.addBorders(`D${subScoreIndex}`);
        this.addText(
          `E${subScoreIndex}`,
          subScore.impactScore < 0 ? `${subScore.impactStatus} (${subScore.impactScore.toFixed(0)})` : 'None',
        );
        this.addBorders(`E${subScoreIndex}`);
        this.addText(`F${subScoreIndex}`, target.note || '');
        this.addBorders(`F${subScoreIndex}`);
        this.addText(`G${subScoreIndex}`, subScore.need.status);
        this.addBorders(`G${subScoreIndex}`);
        this.addText(`H${subScoreIndex}`, subScore.need.countryClassification.status);
        this.addBorders(`H${subScoreIndex}`);
        this.addText(`I${subScoreIndex}`, subScore.need.countryIncome.status);
        this.addBorders(`I${subScoreIndex}`);
        this.addText(`J${subScoreIndex}`, subScore.need.sdgStatus.status);
        this.addBorders(`J${subScoreIndex}`);
        this.addText(`K${subScoreIndex}`, subScore.need.sdgTrend.status);
        this.addBorders(`K${subScoreIndex}`);
        this.addText(`L${subScoreIndex}`, subScore.importance.status);
        this.addBorders(`L${subScoreIndex}`);
        this.addText(`M${subScoreIndex}`, subScore.importance.global.status);
        this.addBorders(`M${subScoreIndex}`);
        this.addText(`N${subScoreIndex}`, subScore.importance.supporting.status);
        this.addBorders(`N${subScoreIndex}`);
        this.addText(`O${subScoreIndex}`, subScore.importance.local.status);
        this.addBorders(`O${subScoreIndex}`);
        this.addText(`P${subScoreIndex}`, subScore.value.status);
        this.addBorders(`P${subScoreIndex}`);
        this.addText(`Q${subScoreIndex}`, subScore.value.depth.status);
        this.addBorders(`Q${subScoreIndex}`);
        this.addText(`R${subScoreIndex}`, subScore.value.immediacy.status);
        this.addBorders(`R${subScoreIndex}`);
        this.addText(`S${subScoreIndex}`, subScore.value.sustained.status);
        this.addBorders(`S${subScoreIndex}`);
        this.addText(`T${subScoreIndex}`, subScore.value.irremediability.status);
        this.addBorders(`T${subScoreIndex}`);
        this.addText(`U${subScoreIndex}`, subScore.contribution?.status || 'None');
        this.addBorders(`U${subScoreIndex}`);
        this.addText(`V${subScoreIndex}`, subScore.contribution?.scale?.status || 'None');
        this.addBorders(`V${subScoreIndex}`);
        this.addText(`W${subScoreIndex}`, subScore.contribution?.change?.status || 'None');
        this.addBorders(`W${subScoreIndex}`);
        subScoreIndex += 1;
      });
    });
    return subScoreIndex + 1;
  }

  populate() {
    this.addHeader();
    this.addBlankRows(1000, 8);
    this.getColumn('D').width = 15;
    this.getColumn('E').width = 15;
    this.getColumn('F').width = 50;
    this.getColumn('G').width = 10;
    this.getColumn('H').width = 10;
    this.getColumn('I').width = 10;
    this.getColumn('J').width = 10;
    this.getColumn('K').width = 10;
    this.getColumn('L').width = 10;
    this.getColumn('M').width = 10;
    this.getColumn('N').width = 10;
    this.getColumn('O').width = 10;
    this.getColumn('P').width = 10;
    this.getColumn('Q').width = 10;
    this.getColumn('R').width = 10;
    this.getColumn('S').width = 10;
    this.getColumn('T').width = 10;
    this.getColumn('U').width = 10;
    this.getColumn('V').width = 10;
    this.getColumn('W').width = 10;
    this.addText('B8', `${this.impact.activity}: Detailed Activity Outcomes Assessment and Data`, {
      bold: true,
      size: 10,
      underline: true,
      wrapText: false,
    });
    this.merge('B9:R9');
    this.addText(
      'B9',
      "Vested Impact identifies evidence-based causal links of individual products, services and activities to UN SDG Targets (leveraging 200 million academic papers). The data also provides the overall impact rating for specific activities (per Vested Impact's methodology of taking into account need, importance, value and effect) and underlying indicators and academic papers also provided for reference.",
      { italic: true },
    );
    this.getRow(9).height = 21;
    this.activitySummaryTable();
    const countriesEnd = this.countrySummaryTable();
    const flagsEnd = this.flagsTables(countriesEnd);
    const subScoresEnd = this.subScoresTable(flagsEnd);
    const metricsEnd = this.metricsTable(subScoresEnd);
    const benchmarksEnd = this.benchmarksTable(metricsEnd);
    const indicatorsEnd = this.indicatorsTable(benchmarksEnd);
    this.referencesTable(indicatorsEnd);
  }
}
