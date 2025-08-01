import { Workbook } from 'exceljs';

import { Asset, AssetImpact } from '../../../api/types';
import { OutputWorksheet } from '../../common/worksheet';

export class PillarsSheet extends OutputWorksheet {
  constructor(
    private readonly impact: AssetImpact,
    private readonly logoId: number,
    private readonly asset: Asset,
    workbook: Workbook,
  ) {
    super('Pillars', workbook);
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
    this.getColumn('C').width = 25;
    this.getColumn('D').width = 25;
    this.getColumn('E').width = 25;
    this.getColumn('F').width = 25;
    this.getColumn('G').width = 25;
    this.getColumn('H').width = 25;
    this.addText('B8', 'Vested Impact Pillar Scores', { bold: true, size: 10, underline: true, wrapText: false });
    this.merge('B9:H9');
    this.addText('B9', '', { italic: true });
    this.getRow(9).height = 40;
    this.addText('B11', 'Business Activity', { wrapText: false });
    this.addBorders('B11');
    this.setColor('B11', 'FFCCCCCC');
    this.addText('C11', 'Country', { wrapText: false });
    this.addBorders('C11');
    this.setColor('C11', 'FFCCCCCC');
    this.addText('D11', 'SDG Target', { wrapText: false });
    this.addBorders('D11');
    this.setColor('D11', 'FFCCCCCC');
    this.addText('E11', 'Contribution Score', { wrapText: false });
    this.addBorders('E11');
    this.setColor('E11', 'FFCCCCCC');
    this.addText('F11', 'Importance Score', { wrapText: false });
    this.addBorders('F11');
    this.setColor('F11', 'FFCCCCCC');
    this.addText('G11', 'Need Score', { wrapText: false });
    this.addBorders('G11');
    this.setColor('G11', 'FFCCCCCC');
    this.addText('H11', 'Value Score', { wrapText: false });
    this.addBorders('H11');
    this.setColor('H11', 'FFCCCCCC');
    const scores = this.impact.impactBreakdown.flatMap((i) => {
      return i.targets.flatMap((j) => {
        return j.subScores.flatMap((k) => ({
          activity: i.activity,
          contribution: k.contribution?.score,
          country: k.country,
          importance: k.importance.score,
          need: k.need.score,
          sdgTarget: j.sdgTarget,
          value: k.value.score,
        }));
      });
    });
    scores.forEach((score, index) => {
      this.addText(`B${12 + index}`, score.activity);
      this.addBorders(`B${12 + index}`);
      this.addText(`C${12 + index}`, score.country);
      this.addBorders(`C${12 + index}`);
      this.addText(`D${12 + index}`, score.sdgTarget);
      this.addBorders(`D${12 + index}`);
      if (score.contribution) {
        this.addNumber(`E${12 + index}`, score.contribution);
      }
      this.addBorders(`E${12 + index}`);
      this.addNumber(`F${12 + index}`, score.importance);
      this.addBorders(`F${12 + index}`);
      this.addNumber(`G${12 + index}`, score.need);
      this.addBorders(`G${12 + index}`);
      this.addNumber(`H${12 + index}`, score.value);
      this.addBorders(`H${12 + index}`);
    });
  }
}
