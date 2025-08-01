import { Workbook } from 'exceljs';

import { Asset, AssetImpact } from '../../../api/types';
import { SDGUtils } from '../../../utils/sdg';
import { OutputWorksheet } from '../../common/worksheet';

export class SummaryImpactSheet extends OutputWorksheet {
  constructor(
    private readonly impact: AssetImpact,
    private readonly logoId: number,
    private readonly asset: Asset,
    workbook: Workbook,
  ) {
    super('Summary', workbook);
    this.setTabColor('FF6631D4');
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
    this.addText('C5', this.impact.reportDate, { wrapText: false });
    this.getColumn('C').width = 25;
    this.getRow(6).height = 5;
    this.getRow(6).border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
  }

  private activitiesTable(startIndex: number) {
    this.addText(`B${startIndex}`, 'Asset Breakdown', { underline: true, wrapText: false });
    this.addText(`B${startIndex + 1}`, 'Business Activity', { wrapText: false });
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 1}`, 'Country', { wrapText: false });
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`D${startIndex + 1}`, 'Proportion Of Asset', { wrapText: false });
    this.addBorders(`D${startIndex + 1}`);
    this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
    this.impact.assetBreakdown.forEach((breakdownItem, index) => {
      this.addText(`B${startIndex + 2 + index}`, breakdownItem.activity);
      this.addBorders(`B${startIndex + 2 + index}`);
      this.addText(`C${startIndex + 2 + index}`, breakdownItem.country);
      this.addBorders(`C${startIndex + 2 + index}`);
      this.addNumber(`D${startIndex + 2 + index}`, breakdownItem.weight, { numFmt: '0.00%' });
      this.addBorders(`D${startIndex + 2 + index}`);
    });
    return startIndex + 3 + this.impact.assetBreakdown.length;
  }

  private activityImpactsTable(startIndex: number) {
    this.addText(`B${startIndex}`, 'Activity Impacts', { underline: true, wrapText: false });
    this.addText(`B${startIndex + 1}`, 'Business Activity', { wrapText: false });
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 1}`, 'Positive Impact', { wrapText: false });
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`D${startIndex + 1}`, 'Negative Impact', { wrapText: false });
    this.addBorders(`D${startIndex + 1}`);
    this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`E${startIndex + 1}`, 'Proportion Of Asset', { wrapText: false });
    this.addBorders(`E${startIndex + 1}`);
    this.setColor(`E${startIndex + 1}`, 'FFCCCCCC');
    this.impact.activityImpacts.forEach((activity, index) => {
      this.addText(`B${startIndex + 2 + index}`, activity.activity);
      this.addBorders(`B${startIndex + 2 + index}`);
      if (activity.positiveImpact !== 0) {
        this.addNumber(`C${startIndex + 2 + index}`, activity.positiveImpact, { numFmt: '#' });
      } else {
        this.addText(`C${startIndex + 2 + index}`, 'None');
      }
      this.addBorders(`C${startIndex + 2 + index}`);
      if (activity.negativeImpact !== 0) {
        this.addNumber(`D${startIndex + 2 + index}`, activity.negativeImpact, { numFmt: '#' });
      } else {
        this.addText(`D${startIndex + 2 + index}`, 'None');
      }
      this.addBorders(`D${startIndex + 2 + index}`);
      this.addNumber(`E${startIndex + 2 + index}`, activity.weight, { numFmt: '0.00%' });
      this.addBorders(`E${startIndex + 2 + index}`);
    });
    return startIndex + 3 + this.impact.activityImpacts.length;
  }

  private countryImpactsTable(startIndex: number) {
    this.addText(`B${startIndex}`, 'Geographic Impacts', { underline: true, wrapText: false });
    this.addText(`B${startIndex + 1}`, 'Country', { wrapText: false });
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 1}`, 'Positive Impact', { wrapText: false });
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`D${startIndex + 1}`, 'Negative Impact', { wrapText: false });
    this.addBorders(`D${startIndex + 1}`);
    this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`E${startIndex + 1}`, 'Proportion Of Asset', { wrapText: false });
    this.addBorders(`E${startIndex + 1}`);
    this.setColor(`E${startIndex + 1}`, 'FFCCCCCC');
    this.impact.countryImpacts.forEach((country, index) => {
      this.addText(`B${startIndex + 2 + index}`, country.country);
      this.addBorders(`B${startIndex + 2 + index}`);
      if (country.positiveImpact !== 0) {
        this.addNumber(`C${startIndex + 2 + index}`, country.positiveImpact, { numFmt: '#' });
      } else {
        this.addText(`C${startIndex + 2 + index}`, 'None');
      }
      this.addBorders(`C${startIndex + 2 + index}`);
      if (country.negativeImpact !== 0) {
        this.addNumber(`D${startIndex + 2 + index}`, country.negativeImpact, { numFmt: '#' });
      } else {
        this.addText(`D${startIndex + 2 + index}`, 'None');
      }
      this.addBorders(`D${startIndex + 2 + index}`);
      this.addNumber(`E${startIndex + 2 + index}`, country.weight, { numFmt: '0.00%' });
      this.addBorders(`E${startIndex + 2 + index}`);
    });
    return startIndex + 3 + this.impact.countryImpacts.length;
  }

  private summaryImpact() {
    this.addText('B8', 'Asset Impact Summary', { underline: true, wrapText: false });
    this.addText('B9', 'Vested Impact Rating', { wrapText: false });
    this.addBorders('B9');
    this.setColor('B9', 'FFCCCCCC');
    this.addText('B10', this.impact.vestedImpactRating);
    this.addBorders('B10');
    this.addText('C9', 'Vested Impact Score', { wrapText: false });
    this.addBorders('C9');
    this.setColor('C9', 'FFCCCCCC');
    this.addNumber('C10', this.impact.vestedImpactScore, { numFmt: '#' });
    this.addBorders('C10');
    this.addText('D9', 'Positive Impact', { wrapText: false });
    this.addBorders('D9');
    this.setColor('D9', 'FFCCCCCC');
    if (this.impact.positiveImpact !== 0) {
      this.addNumber('D10', this.impact.positiveImpact, { numFmt: '#' });
    } else {
      this.addText('D10', 'None');
    }
    this.addBorders('D10');
    this.addText('E9', 'Negative Impact', { wrapText: false });
    this.addBorders('E9');
    this.setColor('E9', 'FFCCCCCC');
    if (this.impact.negativeImpact !== 0) {
      this.addNumber('E10', this.impact.negativeImpact, { numFmt: '#' });
    } else {
      this.addText('E10', 'None');
    }
    this.addBorders('E10');
    return 12;
  }

  private sdgImpactsTable(startIndex: number) {
    this.addText(`B${startIndex}`, 'UN SDG Impacts', { underline: true, wrapText: false });
    this.addText(`B${startIndex + 1}`, 'Goal', { wrapText: false });
    this.addBorders(`B${startIndex + 1}`);
    this.setColor(`B${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 1}`, 'Positive Impact', { wrapText: false });
    this.addBorders(`C${startIndex + 1}`);
    this.setColor(`C${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`D${startIndex + 1}`, 'Negative Impact', { wrapText: false });
    this.addBorders(`D${startIndex + 1}`);
    this.setColor(`D${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`E${startIndex + 1}`, 'Target', { wrapText: false });
    this.addBorders(`E${startIndex + 1}`);
    this.setColor(`E${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`F${startIndex + 1}`, 'Positive Impact', { wrapText: false });
    this.addBorders(`F${startIndex + 1}`);
    this.setColor(`F${startIndex + 1}`, 'FFCCCCCC');
    this.addText(`G${startIndex + 1}`, 'Negative Impact', { wrapText: false });
    this.addBorders(`G${startIndex + 1}`);
    this.setColor(`G${startIndex + 1}`, 'FFCCCCCC');
    let rowIndex = startIndex + 2;
    this.impact.sdgImpacts.forEach((sdg) => {
      this.addText(`B${rowIndex}`, SDGUtils.getGoalLabel(sdg.sdgGoal));
      this.addBorders(`B${rowIndex}`);
      if (sdg.targetImpacts.length > 1) this.merge(`B${rowIndex}:B${rowIndex + sdg.targetImpacts.length - 1}`);
      if (sdg.positiveImpact !== 0) {
        this.addNumber(`C${rowIndex}`, sdg.positiveImpact, { numFmt: '#' });
      } else {
        this.addText(`C${rowIndex}`, 'None');
      }
      this.addBorders(`C${rowIndex}`);
      if (sdg.targetImpacts.length > 1) this.merge(`C${rowIndex}:C${rowIndex + sdg.targetImpacts.length - 1}`);
      if (sdg.negativeImpact !== 0) {
        this.addNumber(`D${rowIndex}`, sdg.negativeImpact, { numFmt: '#' });
      } else {
        this.addText(`D${rowIndex}`, 'None');
      }
      this.addBorders(`D${rowIndex}`);
      if (sdg.targetImpacts.length > 1) this.merge(`D${rowIndex}:D${rowIndex + sdg.targetImpacts.length - 1}`);
      sdg.targetImpacts.forEach((target) => {
        this.addText(`E${rowIndex}`, SDGUtils.getTargetLabel(target.sdgTarget));
        this.addBorders(`E${rowIndex}`);
        if (target.positiveImpact !== 0) {
          this.addNumber(`F${rowIndex}`, target.positiveImpact, { numFmt: '#' });
        } else {
          this.addText(`F${rowIndex}`, 'None');
        }
        this.addBorders(`F${rowIndex}`);
        if (target.negativeImpact !== 0) {
          this.addNumber(`G${rowIndex}`, target.negativeImpact, { numFmt: '#' });
        } else {
          this.addText(`G${rowIndex}`, 'None');
        }
        this.addBorders(`G${rowIndex}`);
        rowIndex += 1;
      });
    });
  }

  populate() {
    this.addHeader();
    this.addBlankRows(1000, 8);
    this.getColumn('D').width = 25;
    this.getColumn('E').width = 25;
    this.getColumn('F').width = 25;
    this.getColumn('G').width = 25;
    const activityBreakdownStart = this.summaryImpact();
    const activityImpactsStart = this.activitiesTable(activityBreakdownStart);
    const countryImpactsStart = this.activityImpactsTable(activityImpactsStart);
    const sdgImpactsStart = this.countryImpactsTable(countryImpactsStart);
    this.sdgImpactsTable(sdgImpactsStart);
  }
}
