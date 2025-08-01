import { Workbook } from 'exceljs';

import { Asset, AssetImpact } from '../../api/types';
import { vestedLogo } from '../common/headerLogo';
import { ActivityDataSheet } from './sheets/Activity';
import { ESGAssessmentSheet } from './sheets/ESG';
import { GlossarySheet } from './sheets/Glossary';
import { EnvironmentalMetricsSheet } from './sheets/Metrics';
import { PillarsSheet } from './sheets/Pillars';
import { SummaryImpactSheet } from './sheets/Summary';

export const createAssetImpactWorkbook = (asset: Asset, impact: AssetImpact) => {
  const workbook = new Workbook();
  workbook.creator = 'Vested Impact';
  workbook.lastModifiedBy = 'Vested Impact';
  const logoId = workbook.addImage({ base64: vestedLogo, extension: 'png' });
  new SummaryImpactSheet(impact, logoId, asset, workbook).populate();
  impact.impactBreakdown.map((item, i) =>
    new ActivityDataSheet(item, impact.reportDate, logoId, asset, workbook, i).populate(),
  );
  new PillarsSheet(impact, logoId, asset, workbook).populate();
  new ESGAssessmentSheet(impact, logoId, asset, workbook).populate();
  new EnvironmentalMetricsSheet(impact, logoId, asset, workbook).populate();
  new GlossarySheet(workbook).populate();
  return workbook;
};
