import { VestedImpactAssetAPI } from './api/asset';
import { createAssetImpactWorkbook } from './excel-output/asset-impact';

const run = async () => {
  const apiKey = '';
  const assetId = '';
  const api = new VestedImpactAssetAPI(apiKey);
  const asset = await api.getAsset(assetId);
  const report = await api.getAssetImpactReport(assetId);
  const workbook = createAssetImpactWorkbook(asset, report);
  await workbook.xlsx.writeFile('example.xlsx');
};

run().catch((e) => {
  console.log(e);
});
