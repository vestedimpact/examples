import { Workbook } from 'exceljs';

import { Asset, AssetImpact } from '../../../api/types';
import { OutputWorksheet } from '../../common/worksheet';

export class EnvironmentalMetricsSheet extends OutputWorksheet {
  constructor(
    private readonly impact: AssetImpact,
    private readonly logoId: number,
    private readonly asset: Asset,
    workbook: Workbook,
  ) {
    super('Metrics', workbook);
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

  populate() {
    this.addHeader();
    this.addBlankRows(1000, 8);
    this.getColumn('B').width = 25;
    this.getColumn('C').width = 60;
    this.getColumn('D').width = 15;
    this.getColumn('E').width = 15;
    this.getColumn('F').width = 15;
    this.addText('B8', 'Environmental Metrics', { bold: true, size: 10, underline: true, wrapText: false });
    this.merge('B9:F9');
    this.addText(
      'B9',
      'Vested Impact calculates product, service, and activity-level metrics such as emissions, water usage, and land use by leveraging Environmentally Extended Input-Output (EEIO) models. The detailed metrics for each activity, product and service within each relevant country are included in the Detailed Impacts by Business Activity section. The table below is an aggregate of all activity metrics for the business.',
      { italic: true },
    );
    this.getRow(9).height = 40;
    this.addText('B11', 'Metric', { wrapText: false });
    this.addBorders('B11');
    this.setColor('B11', 'FFCCCCCC');
    this.addText('C11', 'Description', { wrapText: false });
    this.addBorders('C11');
    this.setColor('C11', 'FFCCCCCC');
    this.addText('D11', 'Category', { wrapText: false });
    this.addBorders('D11');
    this.setColor('D11', 'FFCCCCCC');
    this.addText('E11', 'Overall Value', { wrapText: false });
    this.addBorders('E11');
    this.setColor('E11', 'FFCCCCCC');
    this.addText('F11', 'Units', { wrapText: false });
    this.addBorders('F11');
    this.setColor('F11', 'FFCCCCCC');
    this.impact.overallMetrics.forEach((metric, index) => {
      this.addText(`B${12 + index}`, metric.name, { bold: true });
      this.addBorders(`B${12 + index}`);
      this.addText(`C${12 + index}`, metric.description);
      this.addBorders(`C${12 + index}`);
      this.addText(`D${12 + index}`, metric.category);
      this.addBorders(`D${12 + index}`);
      this.addNumber(`E${12 + index}`, metric.value || 0, { bold: true });
      this.addBorders(`E${12 + index}`);
      this.setColor(`E${12 + index}`, 'FFEEEEEE');
      this.addText(`F${12 + index}`, metric.units);
      this.addBorders(`F${12 + index}`);
    });
  }
}
