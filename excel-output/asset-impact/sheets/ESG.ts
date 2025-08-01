import { Workbook } from 'exceljs';

import { Asset, AssetImpact, Opportunity, Risk } from '../../../api/types';
import { OutputWorksheet } from '../../common/worksheet';

export class ESGAssessmentSheet extends OutputWorksheet {
  constructor(
    private readonly impact: AssetImpact,
    private readonly logoId: number,
    private readonly asset: Asset,
    workbook: Workbook,
  ) {
    super('ESG', workbook);
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

  private opportunitiesTable(startIndex: number, opportunities: Opportunity[]) {
    this.addText(`B${startIndex}`, 'ESG Opportunities', { bold: true, underline: true, wrapText: false });
    this.addText(`B${startIndex + 1}`, 'The following table show the assets positive material impacts;', {
      italic: true,
      wrapText: false,
    });
    this.addText(`B${startIndex + 2}`, 'Impact Materiality');
    this.addBorders(`B${startIndex + 2}`);
    this.setColor(`B${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 2}`, 'Financial Materiality');
    this.addBorders(`C${startIndex + 2}`);
    this.setColor(`C${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`D${startIndex + 2}`, 'ESRS Topic');
    this.addBorders(`D${startIndex + 2}`);
    this.setColor(`D${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`E${startIndex + 2}`, 'SDG Target & Other Affected Frameworks');
    this.addBorders(`E${startIndex + 2}`);
    this.setColor(`E${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`F${startIndex + 2}`, 'Business Activity');
    this.addBorders(`F${startIndex + 2}`);
    this.setColor(`F${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`G${startIndex + 2}`, 'Country');
    this.addBorders(`G${startIndex + 2}`);
    this.setColor(`G${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`H${startIndex + 2}`, 'Actual / Potential');
    this.addBorders(`H${startIndex + 2}`);
    this.setColor(`H${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`I${startIndex + 2}`, 'Type Of Impact');
    this.addBorders(`I${startIndex + 2}`);
    this.setColor(`I${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`J${startIndex + 2}`, 'Value Chain');
    this.addBorders(`J${startIndex + 2}`);
    this.setColor(`J${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`K${startIndex + 2}`, 'Affected Stakeholders');
    this.addBorders(`K${startIndex + 2}`);
    this.setColor(`K${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`L${startIndex + 2}`, 'Time Horizon');
    this.addBorders(`L${startIndex + 2}`);
    this.setColor(`L${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`M${startIndex + 2}`, 'Likelihood');
    this.addBorders(`M${startIndex + 2}`);
    this.setColor(`M${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`N${startIndex + 2}`, 'Irremediability');
    this.addBorders(`N${startIndex + 2}`);
    this.setColor(`N${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`O${startIndex + 2}`, 'Type Of Risk');
    this.addBorders(`O${startIndex + 2}`);
    this.setColor(`O${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`P${startIndex + 2}`, 'Scale');
    this.addBorders(`P${startIndex + 2}`);
    this.setColor(`P${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`Q${startIndex + 2}`, 'Scope');
    this.addBorders(`Q${startIndex + 2}`);
    this.setColor(`Q${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`R${startIndex + 2}`, 'Affected Financial Item');
    this.addBorders(`R${startIndex + 2}`);
    this.setColor(`R${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`S${startIndex + 2}`, 'Financial Materiality Note');
    this.addBorders(`S${startIndex + 2}`);
    this.setColor(`S${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`T${startIndex + 2}`, 'Impact Notes');
    this.addBorders(`T${startIndex + 2}`);
    this.setColor(`T${startIndex + 2}`, 'FFCCCCCC');
    opportunities.forEach((opportunity, index) => {
      this.addText(`B${startIndex + 3 + index}`, `${opportunity.impactStatus} (${opportunity.impactScore.toFixed(0)})`);
      this.addBorders(`B${startIndex + 3 + index}`);
      this.addText(`C${startIndex + 3 + index}`, opportunity.financialMateriality);
      this.addBorders(`C${startIndex + 3 + index}`);
      this.addText(`D${startIndex + 3 + index}`, opportunity.esrsReference);
      this.addBorders(`D${startIndex + 3 + index}`);
      this.addText(
        `E${startIndex + 3 + index}`,
        `${opportunity.sdgReference}${opportunity.additionalFrameworks ? `. ${opportunity.additionalFrameworks}` : ''}`,
      );
      this.addBorders(`E${startIndex + 3 + index}`);
      this.addText(`F${startIndex + 3 + index}`, opportunity.activity);
      this.addBorders(`F${startIndex + 3 + index}`);
      this.addText(`G${startIndex + 3 + index}`, opportunity.country);
      this.addBorders(`G${startIndex + 3 + index}`);
      this.addText(`H${startIndex + 3 + index}`, opportunity.isActual ? 'Actual' : 'Potential');
      this.addBorders(`H${startIndex + 3 + index}`);
      this.addText(`I${startIndex + 3 + index}`, opportunity.impactType);
      this.addBorders(`I${startIndex + 3 + index}`);
      this.addText(`J${startIndex + 3 + index}`, opportunity.valueChain);
      this.addBorders(`J${startIndex + 3 + index}`);
      this.addText(`K${startIndex + 3 + index}`, opportunity.stakeholder);
      this.addBorders(`K${startIndex + 3 + index}`);
      this.addText(`L${startIndex + 3 + index}`, opportunity.timeHorizon);
      this.addBorders(`L${startIndex + 3 + index}`);
      this.addText(`M${startIndex + 3 + index}`, opportunity.likelihood);
      this.addBorders(`M${startIndex + 3 + index}`);
      this.addText(`N${startIndex + 3 + index}`, 'Not available');
      this.addBorders(`N${startIndex + 3 + index}`);
      this.addText(`O${startIndex + 3 + index}`, opportunity.riskType);
      this.addBorders(`O${startIndex + 3 + index}`);
      this.addText(`P${startIndex + 3 + index}`, opportunity.scale);
      this.addBorders(`P${startIndex + 3 + index}`);
      this.addText(`Q${startIndex + 3 + index}`, opportunity.scope);
      this.addBorders(`Q${startIndex + 3 + index}`);
      this.addText(`R${startIndex + 3 + index}`, opportunity.affectedFinancialItem);
      this.addBorders(`R${startIndex + 3 + index}`);
      this.addText(`S${startIndex + 3 + index}`, opportunity.financialMaterialityNote);
      this.addBorders(`S${startIndex + 3 + index}`);
      this.addText(`T${startIndex + 3 + index}`, opportunity.description);
      this.addBorders(`T${startIndex + 3 + index}`);
    });
  }

  private risksTable(startIndex: number, risks: Risk[]) {
    this.addText(`B${startIndex}`, 'ESG Risks', { bold: true, underline: true, wrapText: false });
    this.addText(`B${startIndex + 1}`, 'The following table show the assets negative material impacts;', {
      italic: true,
      wrapText: false,
    });
    this.addText(`B${startIndex + 2}`, 'Impact Materiality');
    this.addBorders(`B${startIndex + 2}`);
    this.setColor(`B${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`C${startIndex + 2}`, 'Financial Materiality');
    this.addBorders(`C${startIndex + 2}`);
    this.setColor(`C${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`D${startIndex + 2}`, 'ESRS Topic');
    this.addBorders(`D${startIndex + 2}`);
    this.setColor(`D${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`E${startIndex + 2}`, 'SDG Target & Other Affected Frameworks');
    this.addBorders(`E${startIndex + 2}`);
    this.setColor(`E${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`F${startIndex + 2}`, 'Business Activity');
    this.addBorders(`F${startIndex + 2}`);
    this.setColor(`F${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`G${startIndex + 2}`, 'Country');
    this.addBorders(`G${startIndex + 2}`);
    this.setColor(`G${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`H${startIndex + 2}`, 'Actual / Potential');
    this.addBorders(`H${startIndex + 2}`);
    this.setColor(`H${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`I${startIndex + 2}`, 'Type Of Impact');
    this.addBorders(`I${startIndex + 2}`);
    this.setColor(`I${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`J${startIndex + 2}`, 'Value Chain');
    this.addBorders(`J${startIndex + 2}`);
    this.setColor(`J${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`K${startIndex + 2}`, 'Affected Stakeholders');
    this.addBorders(`K${startIndex + 2}`);
    this.setColor(`K${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`L${startIndex + 2}`, 'Time Horizon');
    this.addBorders(`L${startIndex + 2}`);
    this.setColor(`L${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`M${startIndex + 2}`, 'Likelihood');
    this.addBorders(`M${startIndex + 2}`);
    this.setColor(`M${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`N${startIndex + 2}`, 'Irremediability');
    this.addBorders(`N${startIndex + 2}`);
    this.setColor(`N${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`O${startIndex + 2}`, 'Type Of Risk');
    this.addBorders(`O${startIndex + 2}`);
    this.setColor(`O${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`P${startIndex + 2}`, 'Scale');
    this.addBorders(`P${startIndex + 2}`);
    this.setColor(`P${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`Q${startIndex + 2}`, 'Scope');
    this.addBorders(`Q${startIndex + 2}`);
    this.setColor(`Q${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`R${startIndex + 2}`, 'Affected Financial Item');
    this.addBorders(`R${startIndex + 2}`);
    this.setColor(`R${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`S${startIndex + 2}`, 'Financial Materiality Note');
    this.addBorders(`S${startIndex + 2}`);
    this.setColor(`S${startIndex + 2}`, 'FFCCCCCC');
    this.addText(`T${startIndex + 2}`, 'Impact Notes');
    this.addBorders(`T${startIndex + 2}`);
    this.setColor(`T${startIndex + 2}`, 'FFCCCCCC');
    risks.forEach((risk, index) => {
      this.addText(`B${startIndex + 3 + index}`, `${risk.impactStatus} (${risk.impactScore.toFixed(0)})`);
      this.addBorders(`B${startIndex + 3 + index}`);
      this.addText(`C${startIndex + 3 + index}`, risk.financialMateriality);
      this.addBorders(`C${startIndex + 3 + index}`);
      this.addText(`D${startIndex + 3 + index}`, risk.esrsReference);
      this.addBorders(`D${startIndex + 3 + index}`);
      this.addText(
        `E${startIndex + 3 + index}`,
        `${risk.sdgReference}${risk.additionalFrameworks ? `. ${risk.additionalFrameworks}` : ''}`,
      );
      this.addBorders(`E${startIndex + 3 + index}`);
      this.addText(`F${startIndex + 3 + index}`, risk.activity);
      this.addBorders(`F${startIndex + 3 + index}`);
      this.addText(`G${startIndex + 3 + index}`, risk.country);
      this.addBorders(`G${startIndex + 3 + index}`);
      this.addText(`H${startIndex + 3 + index}`, risk.isActual ? 'Actual' : 'Potential');
      this.addBorders(`H${startIndex + 3 + index}`);
      this.addText(`I${startIndex + 3 + index}`, risk.impactType);
      this.addBorders(`I${startIndex + 3 + index}`);
      this.addText(`J${startIndex + 3 + index}`, risk.valueChain);
      this.addBorders(`J${startIndex + 3 + index}`);
      this.addText(`K${startIndex + 3 + index}`, risk.stakeholder);
      this.addBorders(`K${startIndex + 3 + index}`);
      this.addText(`L${startIndex + 3 + index}`, risk.timeHorizon);
      this.addBorders(`L${startIndex + 3 + index}`);
      this.addText(`M${startIndex + 3 + index}`, risk.likelihood);
      this.addBorders(`M${startIndex + 3 + index}`);
      this.addText(`N${startIndex + 3 + index}`, risk.irremediability);
      this.addBorders(`N${startIndex + 3 + index}`);
      this.addText(`O${startIndex + 3 + index}`, risk.riskType);
      this.addBorders(`O${startIndex + 3 + index}`);
      this.addText(`P${startIndex + 3 + index}`, risk.scale);
      this.addBorders(`P${startIndex + 3 + index}`);
      this.addText(`Q${startIndex + 3 + index}`, risk.scope);
      this.addBorders(`Q${startIndex + 3 + index}`);
      this.addText(`R${startIndex + 3 + index}`, risk.affectedFinancialItem);
      this.addBorders(`R${startIndex + 3 + index}`);
      this.addText(`S${startIndex + 3 + index}`, risk.financialMaterialityNote);
      this.addBorders(`S${startIndex + 3 + index}`);
      this.addText(`T${startIndex + 3 + index}`, risk.description);
      this.addBorders(`T${startIndex + 3 + index}`);
    });
  }

  populate() {
    const opportunities = [...this.impact.opportunities.E, ...this.impact.opportunities.S, ...this.impact.opportunities.G].sort(
      (a, b) => b.impactScore - a.impactScore,
    );
    const risks = [...this.impact.risks.E, ...this.impact.risks.S, ...this.impact.risks.G].sort(
      (a, b) => a.impactScore - b.impactScore,
    );
    this.addHeader();
    this.addBlankRows(120 + risks.length + opportunities.length, 8);
    this.getColumn('D').width = 20;
    this.getColumn('E').width = 25;
    this.getColumn('F').width = 15;
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
    this.getColumn('R').width = 20;
    this.getColumn('S').width = 60;
    this.getColumn('T').width = 60;
    this.addText('B8', 'ESG Risks and Opportunities Assessment', {
      bold: true,
      size: 10,
      underline: true,
      wrapText: false,
    });
    this.merge('B9:R9');
    this.addText(
      'B9',
      'Vested Impact conducts an automated “impact materiality assessment”, in line with the EU’s new CSRD regulation and the OECD Business Due Diligence Guidelines. This can serve as identification of any material risks the organisation has, and a template for where responses should be gathered from the organisation. These are anticipated to form the basis for the EU CSDDD regulation, which will impact UBS in terms of requiring impact materiality assessment of supply chain/assets. The risks are derived from any negative impacts Vested Impact detects, aligned with the UN SDG targets, and are also mapped to relevant regulatory frameworks. As per regulation, organisations do NOT have to have zero negative, they simply need to disclose where/if they agree risks are material and any comments about the organisations existing or planned mitigation.',
      { italic: true },
    );
    this.getRow(9).height = 33;
    const risksStart = 11;
    if (risks.length > 0) {
      this.risksTable(risksStart, risks);
    } else {
      this.addText(`B${risksStart}`, 'ESG Risks', { bold: true, underline: true, wrapText: false });
      this.merge(`B${risksStart + 1}:R${risksStart + 1}`);
      this.addText(
        `B${risksStart + 1}`,
        'No significant material adverse impacts for the business have been detected. The absence of material adverse impacts does not imply that the business has absolute zero negative impacts, but that given the country, sector and millions of scientific articles and data points, along with the business-specific product and service portfolio, no negative impacts or risks thereof have been assigned by the Vested Impact system.',
      );
      this.getRow(risksStart + 1).height = 30;
    }
    const opportunitiesStart = risks.length === 0 ? risksStart + 3 : risksStart + 3 + risks.length;
    if (opportunities.length > 0) {
      this.opportunitiesTable(opportunitiesStart, opportunities);
    } else {
      this.addText(`B${opportunitiesStart}`, 'ESG Opportunities', { bold: true, underline: true, wrapText: false });
      this.merge(`B${opportunitiesStart + 1}:R${opportunitiesStart + 1}`);
      this.addText(
        `B${opportunitiesStart + 1}`,
        'No significant material positive impacts for the business have been detected. The absence of material positive impacts does not imply that the business has absolute zero positive impacts, but that given the country, sector and millions of scientific articles and data points, along with the business-specific product and service portfolio, no positive impacts have been assigned by the Vested Impact system.',
      );
      this.getRow(opportunitiesStart + 1).height = 30;
    }
  }
}
