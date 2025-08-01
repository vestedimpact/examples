import { Workbook } from 'exceljs';

import { OutputWorksheet } from '../../common/worksheet';

export class GlossarySheet extends OutputWorksheet {
  constructor(workbook: Workbook) {
    super('Glossary', workbook);
    this.setTabColor('FF30FEF6');
  }

  private activityImpact(startRow: number) {
    this.addText(`B${startRow}`, 'Activity impact data', { bold: true, underline: true, wrapText: false });
    this.addText(`B${startRow + 1}`, 'Field Name', { bold: true });
    this.addBorders(`B${startRow + 1}`);
    this.setColor(`B${startRow + 1}`, 'FFCCCCCC');
    this.addText(`C${startRow + 1}`, 'Description', { bold: true });
    this.addBorders(`C${startRow + 1}`);
    this.setColor(`C${startRow + 1}`, 'FFCCCCCC');
    this.addText(`B${startRow + 2}`, 'Flags', { bold: true });
    this.addBorders(`B${startRow + 2}`);
    this.setColor(`B${startRow + 2}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 2}`,
      'The flags highlight potential  alignment/mis-alignment to high-impact frameworks and/or human rights impacts. They serve as an indicator of potential risk or impact and are not a guarantee of alignment/impact and, as such, are not included in the calculation of overall scores.The flags are presented to ensure awareness of potential issues.',
    );
    this.addBorders(`C${startRow + 2}`);
    this.setColor(`C${startRow + 2}`, 'FFEEEEEE');
    this.addText(`B${startRow + 3}`, 'Flag Type');
    this.addBorders(`B${startRow + 3}`);
    this.addText(`C${startRow + 3}`, 'Indicates the primary category/type of the flag');
    this.addBorders(`C${startRow + 3}`);
    this.addText(`B${startRow + 4}`, 'Status');
    this.addBorders(`B${startRow + 4}`);
    this.addText(
      `C${startRow + 4}`,
      'Indicates the status/outcome of the flag allocated to the activity, country and/or SDG',
    );
    this.addBorders(`C${startRow + 4}`);
    this.addText(`B${startRow + 5}`, 'Countries');
    this.addBorders(`B${startRow + 5}`);
    this.addText(`C${startRow + 5}`, 'Country where the activity is being assessed');
    this.addBorders(`C${startRow + 5}`);
    this.addText(`B${startRow + 6}`, 'SDG Targets');
    this.addBorders(`B${startRow + 6}`);
    this.addText(
      `C${startRow + 6}`,
      'UN SDG target that is impacted by the product, service, activity and/or intervention of the asset',
    );
    this.addBorders(`C${startRow + 6}`);
    this.addText(`B${startRow + 7}`, 'Note');
    this.addBorders(`B${startRow + 7}`);
    this.addText(
      `C${startRow + 7}`,
      'A descriptive note describing/justifying the allocation of the flag to the activity, country and/or SDG target',
    );
    this.addBorders(`C${startRow + 7}`);
    this.addText(`B${startRow + 8}`, 'Outcomes', { bold: true });
    this.addBorders(`B${startRow + 8}`);
    this.setColor(`B${startRow + 8}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 8}`,
      'Impact outcomes indicate whether, and to what degree, a company’s products, services, or activities advance or hinder progress toward specific SDG targets. Vested Impact assesses this across impact slices (activity-target-country combinations) using science-based evidence, and using the underlying Outcomes sub-scores/data, resulting in quantified net-positive or net-negative effects against the relevant SDG targets.',
    );
    this.addBorders(`C${startRow + 8}`);
    this.setColor(`C${startRow + 8}`, 'FFEEEEEE');
    this.addText(`B${startRow + 9}`, 'SDG Target');
    this.addBorders(`B${startRow + 9}`);
    this.addText(
      `C${startRow + 9}`,
      'The specific UN Sustainable Development Goal target (from among the 169 targets) that the assessed business activity influences. Outcomes are measured relative to each target, quantifying whether an activity advances or hinders that goal.',
    );
    this.addBorders(`C${startRow + 9}`);
    this.addText(`B${startRow + 10}`, 'Country');
    this.addBorders(`B${startRow + 10}`);
    this.addText(
      `C${startRow + 10}`,
      'The geographic context in which the impact occurs. Vested Impact localizes SDG outcomes to the specific country where the product, service, or activity is delivered—accounting for local needs, development levels, and progress toward the target.',
    );
    this.addBorders(`C${startRow + 10}`);
    this.addText(`B${startRow + 11}`, 'Positive Impact');
    this.addBorders(`B${startRow + 11}`);
    this.addText(
      `C${startRow + 11}`,
      'The beneficial effect that an activity has on advancing progress toward a specific SDG target. Positive impacts are scored based on their depth, immediacy, duration, and alignment with needs in that country.',
    );
    this.addBorders(`C${startRow + 11}`);
    this.addText(`B${startRow + 12}`, 'Negative Impact');
    this.addBorders(`B${startRow + 12}`);
    this.addText(
      `C${startRow + 12}`,
      'The harmful effect an activity causes, directly or indirectly, in relation to an SDG target. Negative outcomes are assessed and scored with additional factors like irremediability (how difficult it is to reverse the harm).',
    );
    this.addBorders(`C${startRow + 12}`);
    this.addText(`B${startRow + 13}`, 'Findings');
    this.addBorders(`B${startRow + 13}`);
    this.addText(
      `C${startRow + 13}`,
      'Findings provide an AI-generated summary of key academic findings drawn from the underlying References linked to the specific SDG Target and activity. This section synthesises the most relevant evidence from peer-reviewed literature, offering context and justification for the impact assessment in a clear and accessible format.',
    );
    this.addBorders(`C${startRow + 13}`);
    this.addText(`B${startRow + 14}`, 'Need Score');
    this.addBorders(`B${startRow + 14}`);
    this.addText(
      `C${startRow + 14}`,
      'Gives the asset’s average need score, across all impact slices. The closer the value pillar score is to 100, the more the asset is having a positive impact on the targets for which there is a very high need for progress in the countries it affects. The closer the value pillar score is to -100, the more the asset is having a negative impact on the targets for which there is a very high need for progress in the countries it affects.',
    );
    this.addBorders(`C${startRow + 14}`);
    this.addText(`B${startRow + 15}`, 'Need / UN Classification');
    this.addBorders(`B${startRow + 15}`);
    this.addText(
      `C${startRow + 15}`,
      'Refers to the development classification of the country as defined by the United Nations (e.g., least developed country, developing, developed). It helps assess the urgency and relevance of the activity’s impact based on structural vulnerabilities.',
    );
    this.addBorders(`C${startRow + 15}`);
    this.addText(`B${startRow + 16}`, 'Need / World Bank Income Group');
    this.addBorders(`B${startRow + 16}`);
    this.addText(
      `C${startRow + 16}`,
      'Categorizes countries based on World Bank income groupings (e.g., low-income, lower-middle, upper-middle, high-income). This informs the Need Pillar Score, indicating how impactful an activity is depending on the economic context of the population affected.',
    );
    this.addBorders(`C${startRow + 16}`);
    this.addText(`B${startRow + 17}`, 'Need / SDG Status');
    this.addBorders(`B${startRow + 17}`);
    this.addText(
      `C${startRow + 17}`,
      'The status score value indicating how well the country is on trackprogressing  towards the SDG Target, relative to other countries',
    );
    this.addBorders(`C${startRow + 17}`);
    this.addText(`B${startRow + 18}`, 'Need / SDG Trend');
    this.addBorders(`B${startRow + 18}`);
    this.addText(
      `C${startRow + 18}`,
      'The trend score value indicating the direction and rate of progress of the country to meeting the 2030 target.',
    );
    this.addBorders(`C${startRow + 18}`);
    this.addText(`B${startRow + 19}`, 'Importance Score');
    this.addBorders(`B${startRow + 19}`);
    this.addText(
      `C${startRow + 19}`,
      'Gives the asset’s average importance score, across all impact slices. The closer the importance pillar score is to 100, the more the asset is having a positive impact on the targets that global and local communities deem to be important in the countries it affects. The closer the value pillar score is to -100, the more the asset is having a negative impact on the targets that global and local communities deem to be important in the countries it affects.',
    );
    this.addBorders(`C${startRow + 19}`);
    this.addText(`B${startRow + 20}`, 'Importance / Global Score');
    this.addBorders(`B${startRow + 20}`);
    this.addText(
      `C${startRow + 20}`,
      'The criticality score value indicating importance of the target being affected to the general heirarchy of needs of the global population. Based on global models such as the IDM.',
    );
    this.addBorders(`C${startRow + 20}`);
    this.addText(`B${startRow + 21}`, 'Importance / Supporting Score');
    this.addBorders(`B${startRow + 21}`);
    this.addText(
      `C${startRow + 21}`,
      'The supporting score value indicating whether more basic prerequisite needs than the affected target have been met',
    );
    this.addBorders(`C${startRow + 21}`);
    this.addText(`B${startRow + 22}`, 'Importance / Local Score');
    this.addBorders(`B${startRow + 22}`);
    this.addText(
      `C${startRow + 22}`,
      'The survey rank indicates the rank of the importance of the SDG target being affected to the individuals in the country being affected, relative to other SDG targets. Sourced from OECD Better Life Survey',
    );
    this.addBorders(`C${startRow + 22}`);
    this.addText(`B${startRow + 23}`, 'Value Score');
    this.addBorders(`B${startRow + 23}`);
    this.addText(
      `C${startRow + 23}`,
      'Gives the asset’s average value score, across all impact slices. The closer the value pillar score is to 100, the more the asset is having a direct, positive impact on the targets it influences, across all the countries it affects. The closer the value pillar score is to -100, the more the asset is having a very direct negative impact on the targets it influences across all the countries it affects.',
    );
    this.addBorders(`C${startRow + 23}`);
    this.addText(`B${startRow + 24}`, 'Value / Depth Score');
    this.addBorders(`B${startRow + 24}`);
    this.addText(
      `C${startRow + 24}`,
      "The depth and type of impact by the activity, based on the OECD Due Diligence Guidelines for MNE's.\nAllowed values:\n- Caused by\n- Contributes to\n- Directly linked\n- Indirectly linked",
    );
    this.addBorders(`C${startRow + 24}`);
    this.addText(`B${startRow + 25}`, 'Value / Immediacy Score');
    this.addBorders(`B${startRow + 25}`);
    this.addText(
      `C${startRow + 25}`,
      'The immediacy score indicates how immediate an impact, positive or negative, will be felt (derived from AI and human-generated summaries from academic references, as cited)',
    );
    this.addBorders(`C${startRow + 25}`);
    this.addText(`B${startRow + 26}`, 'Value / Sustained Score');
    this.addBorders(`B${startRow + 26}`);
    this.addText(
      `C${startRow + 26}`,
      'The immediacy score indicates how immediate an impact, positive or negative, will be felt (derived from AI and human-generated summaries from academic references, as cited)',
    );
    this.addBorders(`C${startRow + 26}`);
    this.addText(`B${startRow + 27}`, 'Value / Irremediable Score');
    this.addBorders(`B${startRow + 27}`);
    this.addText(
      `C${startRow + 27}`,
      'The irremediability score applies only to negative impacts, and considers how easily an impact could be reversed/remmediated once it has materialised (derived from AI and human-generated summaries from academic references, as cited)',
    );
    this.addBorders(`C${startRow + 27}`);
    this.addText(`B${startRow + 28}`, 'Contribution Score');
    this.addBorders(`B${startRow + 28}`);
    this.addText(
      `C${startRow + 28}`,
      'Gives the asset’s average contribution score across all impact slices. The closer the value pillar score is to 100, the more the business can assumed to be contributing to positive change on the targets and countries being affected.  The closer the contribution pillar score is to -100, the more the asset can assumed to be contributing to negative change in the targets and countries being affected.',
    );
    this.addBorders(`C${startRow + 28}`);
    this.addText(`B${startRow + 29}`, 'Contribution / Scale Score');
    this.addBorders(`B${startRow + 29}`);
    this.addText(
      `C${startRow + 29}`,
      "This ratio provides a measure of how an asset's revenue compares to the market size of the activities it influences. A higher ratio implies that the asset has a larger footprint relative to its relevant market(s). This is calculated by dividing the asset's latest reported annual revenue by the total market size of the activities it impacts.",
    );
    this.addBorders(`C${startRow + 29}`);
    this.addText(`B${startRow + 30}`, 'Contribution / Change Score');
    this.addBorders(`B${startRow + 30}`);
    this.addText(
      `C${startRow + 30}`,
      "The Change Score is designed to measure how an asset’s revenue growth aligns with progress toward sustainable development goals (SDGs) and related indicators. It integrates multiple dimensions, including the size of the asset, its growth rate, and the trends in relevant impact indicators, to provide a standardized measure of an asset’s contribution to positive change.  High Positive Score: Indicates strong alignment between the asset's revenue growth and positive changes in SDG indicators. Low or Negative Score: Suggests a misalignment or potential regression, where the asset's growth does not correspond to meaningful progress on relevant indicators.",
    );
    this.addBorders(`C${startRow + 30}`);
    this.addText(`B${startRow + 31}`, 'Metrics', { bold: true });
    this.addBorders(`B${startRow + 31}`);
    this.setColor(`B${startRow + 31}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 31}`,
      'Metrics provide quantified estimates of environmental and social effects (e.g., emissions, water use, land use, toxic releases). These are calculated using input-output models like USEEIO, linked to company revenue/expenditure and activity classification. Metrics support traceable, comparable environmental footprint estimations aligned to ESG standards like EU CSRD and SFDR.',
    );
    this.addBorders(`C${startRow + 31}`);
    this.setColor(`C${startRow + 31}`, 'FFEEEEEE');
    this.addText(`B${startRow + 32}`, 'Metric');
    this.addBorders(`B${startRow + 32}`);
    this.addText(
      `C${startRow + 32}`,
      'The specific measurable variable used to quantify the environmental or social footprint of an activity (e.g., CO₂ emissions, water use, land area affected). These are typically derived using input-output models like USEEIO.',
    );
    this.addBorders(`C${startRow + 32}`);
    this.addText(`B${startRow + 33}`, 'Description');
    this.addBorders(`B${startRow + 33}`);
    this.addText(
      `C${startRow + 33}`,
      'A short explanation of what the metric measures, including its scope and relevance (e.g., “Total upstream and downstream greenhouse gas emissions associated with the activity”), generally as defined by USEEIO',
    );
    this.addBorders(`C${startRow + 33}`);
    this.addText(`B${startRow + 34}`, 'Category');
    this.addBorders(`B${startRow + 34}`);
    this.addText(
      `C${startRow + 34}`,
      'The thematic classification of the metric, such as Emissions, Water & Effluents, Human Health, Resource Use, or Biodiversity. This helps group and interpret metrics within broader environmental or social impact domains.',
    );
    this.addBorders(`C${startRow + 34}`);
    this.addText(`B${startRow + 35}`, 'Overall Value');
    this.addBorders(`B${startRow + 35}`);
    this.addText(
      `C${startRow + 35}`,
      'Vested Impact takes the revenue or expenditure data provided for each activity and multiplies It by the equivalent factor (based on USEEIO and/or adjusted country specific tables and factors). Where Overall Value = (revenue/expenditure) X (spend based factor)',
    );
    this.addBorders(`C${startRow + 35}`);
    this.addText(`B${startRow + 36}`, 'Units');
    this.addBorders(`B${startRow + 36}`);
    this.addText(
      `C${startRow + 36}`,
      'The unit of measurement for the metric (e.g., kg CO₂e, m³ water, MJ energy), allowing consistent comparison and integration into impact assessments and regulatory disclosures.',
    );
    this.addBorders(`C${startRow + 36}`);
    this.addText(`B${startRow + 37}`, 'Benchmarks', { bold: true });
    this.addBorders(`B${startRow + 37}`);
    this.setColor(`B${startRow + 37}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 37}`,
      'Benchmarks assesses an organisation’s impact performance against the pace and scale of change needed to achieve specific UN SDG targets by 2030. It compares actual impact results with required growth or reduction trajectories (e.g. income growth, emission reductions), using globally recognised benchmarks like those developed by GIIN and leveraging authoritative data from sources like the World Bank, ILO, and Global Findex. This enables benchmarking of both relative and absolute progress toward SDG outcomes.',
    );
    this.addBorders(`C${startRow + 37}`);
    this.setColor(`C${startRow + 37}`, 'FFEEEEEE');
    this.addText(`B${startRow + 38}`, 'SDG Target');
    this.addBorders(`B${startRow + 38}`);
    this.addText(
      `C${startRow + 38}`,
      'The specific UN Sustainable Development Goal target against which impact is assessed. Each target defines a thematic objective (e.g., doubling smallholder income or reducing gender inequality) and provides a quantifiable outcome required by 2030. Benchmarking uses these targets as reference points for assessing whether an organisation’s activities contribute sufficiently to their achievement.',
    );
    this.addBorders(`C${startRow + 38}`);
    this.addText(`B${startRow + 39}`, 'Country');
    this.addBorders(`B${startRow + 39}`);
    this.addText(
      `C${startRow + 39}`,
      'The national context in which benchmarking occurs, incorporating a country’s baseline status and required pace of change toward a specific SDG target. This helps adjust benchmarks based on localized needs and trajectories, enabling country-specific comparisons (e.g., SDG 2.3 in Uganda vs. Denmark).',
    );
    this.addBorders(`C${startRow + 39}`);
    this.addText(`B${startRow + 40}`, 'Indicator Used');
    this.addBorders(`B${startRow + 40}`);
    this.addText(
      `C${startRow + 40}`,
      'The specific data indicator selected to represent progress toward an SDG target (e.g., average farmer income, literacy rate, % of board members who are women). Indicators are chosen for their credibility, relevance to the activity, and alignment with the SDG target, often sourced from trusted databases such as the ILO or World Bank.',
    );
    this.addBorders(`C${startRow + 40}`);
    this.addText(`B${startRow + 41}`, 'Annual Required Change');
    this.addBorders(`B${startRow + 41}`);
    this.addText(
      `C${startRow + 41}`,
      'The target annual growth (or decay) rate needed from the current baseline to reach the SDG target by 2030. This is calculated using compound growth or exponential decay formulas, depending on whether the goal is to grow a positive indicator (e.g., income) or reduce a negative one (e.g., emissions). It defines the benchmark pace for achieving the desired outcome.',
    );
    this.addBorders(`C${startRow + 41}`);
    this.addText(`B${startRow + 42}`, 'Organization Change');
    this.addBorders(`B${startRow + 42}`);
    this.addText(
      `C${startRow + 42}`,
      'The annualized growth/change rate of the organisation. Comparing this to the Annual Required Change highlights whether the organisation is on track, ahead, or falling behind in contributing toward the SDG target in that region or globally.',
    );
    this.addBorders(`C${startRow + 42}`);
    this.addText(`B${startRow + 43}`, 'Indicators', { bold: true });
    this.addBorders(`B${startRow + 43}`);
    this.setColor(`B${startRow + 43}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 43}`,
      'Indicators are data points (over 40,000 integrated from 100M+ sources) used to track performance against SDG targets and contextualize impact. Each indicator is manually mapped to SDG targets and activities to strengthen accountability and traceability. They inform both sub-scores and regulatory-aligned reporting.',
    );
    this.addBorders(`C${startRow + 43}`);
    this.setColor(`C${startRow + 43}`, 'FFEEEEEE');
    this.addText(`B${startRow + 44}`, 'SDG Target');
    this.addBorders(`B${startRow + 44}`);
    this.addText(
      `C${startRow + 44}`,
      'The specific UN SDG target that the indicator is mapped to. Each indicator is manually aligned to a particular SDG target to measure and monitor progress or impact relative to that goal, enabling consistent attribution and benchmarking across activities.',
    );
    this.addBorders(`C${startRow + 44}`);
    this.addText(`B${startRow + 45}`, 'Country');
    this.addBorders(`B${startRow + 45}`);
    this.addText(
      `C${startRow + 45}`,
      'The geographic context for which the indicator data is applied. Indicators are adapted to reflect country-specific values, trends, or baselines, ensuring that impact is assessed with the appropriate local relevance and development context.',
    );
    this.addBorders(`C${startRow + 45}`);
    this.addText(`B${startRow + 46}`, 'Indicator');
    this.addBorders(`B${startRow + 46}`);
    this.addText(
      `C${startRow + 46}`,
      'A quantitative or qualitative metric used to assess progress toward an SDG target. Vested Impact integrates over 40,000 indicators (e.g., % population with access to clean water, GHG emissions per capita), each validated and mapped to relevant activities and geographies for robust impact assessment.',
    );
    this.addBorders(`C${startRow + 46}`);
    this.addText(`B${startRow + 47}`, 'Source');
    this.addBorders(`B${startRow + 47}`);
    this.addText(
      `C${startRow + 47}`,
      'The primary source of the indicator. In many cases Vested Impact accesses the data from centralised aggregators (UN, World Bank etc) but where possible, will always disclose the original source of the data as cited by the aggregator.',
    );
    this.addBorders(`C${startRow + 47}`);
    this.addText(`B${startRow + 48}`, 'Trend');
    this.addBorders(`B${startRow + 48}`);
    this.addText(
      `C${startRow + 48}`,
      'The trend of the indicators is calculated from the change in the most recent 2 available values. Where there is a gap in the data the blank data point is disregarded, and where there are more than 3 consecutive data gaps the indicator is discarded and not used',
    );
    this.addBorders(`C${startRow + 48}`);
    this.addText(`B${startRow + 49}`, 'References', { bold: true });
    this.addBorders(`B${startRow + 49}`);
    this.setColor(`B${startRow + 49}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 49}`,
      'References include over 200 million academic articles (via Semantic Scholar and Elicit) used to determine the causal relationship between activities and SDG outcomes. These are assessed using AI and human review to ensure the credibility of evidence and underpin the scoring model.',
    );
    this.addBorders(`C${startRow + 49}`);
    this.setColor(`C${startRow + 49}`, 'FFEEEEEE');
    this.addText(`B${startRow + 50}`, 'SDG Targets');
    this.addBorders(`B${startRow + 50}`);
    this.addText(
      `C${startRow + 50}`,
      'UN SDG target that is impacted by the product, service, activity and/or intervention of the asset',
    );
    this.addBorders(`C${startRow + 50}`);
    this.addText(`B${startRow + 51}`, 'Reference');
    this.addBorders(`B${startRow + 51}`);
    this.addText(
      `C${startRow + 51}`,
      'The citation that includes the author, year of publication, journal of publication, paper title.',
    );
    this.addBorders(`C${startRow + 51}`);
    this.addText(`B${startRow + 52}`, 'URL');
    this.addBorders(`B${startRow + 52}`);
    this.addText(`C${startRow + 52}`, 'The DOI link to the academic paper/article');
    this.addBorders(`C${startRow + 52}`);
    this.addText(`B${startRow + 53}`, 'Standards Of Evidence');
    this.addBorders(`B${startRow + 53}`);
    this.addText(
      `C${startRow + 53}`,
      'A status indicating the standard of evidence used to make the causal link, in line with Nest Standards of Evidence categories',
    );
    this.addBorders(`C${startRow + 53}`);
  }

  private environmentalMetrics(startRow: number) {
    this.addText(`B${startRow}`, 'Environmental metrics', { bold: true, underline: true, wrapText: false });
    this.addText(`B${startRow + 1}`, 'Field Name', { bold: true });
    this.addBorders(`B${startRow + 1}`);
    this.setColor(`B${startRow + 1}`, 'FFCCCCCC');
    this.addText(`C${startRow + 1}`, 'Description', { bold: true });
    this.addBorders(`C${startRow + 1}`);
    this.setColor(`C${startRow + 1}`, 'FFCCCCCC');
    this.addText(`B${startRow + 2}`, 'Metrics', { bold: true });
    this.addBorders(`B${startRow + 2}`);
    this.setColor(`B${startRow + 2}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 2}`,
      'Metrics provide quantified estimates of environmental and social effects (e.g., emissions, water use, land use, toxic releases). These are calculated using input-output models like USEEIO, linked to company revenue/expenditure and activity classification. Metrics support traceable, comparable environmental footprint estimations aligned to ESG standards like EU CSRD and SFDR.',
    );
    this.addBorders(`C${startRow + 2}`);
    this.setColor(`C${startRow + 2}`, 'FFEEEEEE');
    this.addText(`B${startRow + 3}`, 'Metric');
    this.addBorders(`B${startRow + 3}`);
    this.addText(
      `C${startRow + 3}`,
      'The specific measurable variable used to quantify the environmental or social footprint of an activity (e.g., CO₂ emissions, water use, land area affected). These are typically derived using input-output models like USEEIO.',
    );
    this.addBorders(`C${startRow + 3}`);
    this.addText(`B${startRow + 4}`, 'Description');
    this.addBorders(`B${startRow + 4}`);
    this.addText(
      `C${startRow + 4}`,
      'A short explanation of what the metric measures, including its scope and relevance (e.g., “Total upstream and downstream greenhouse gas emissions associated with the activity”), generally as defined by USEEIO',
    );
    this.addBorders(`C${startRow + 4}`);
    this.addText(`B${startRow + 5}`, 'Category');
    this.addBorders(`B${startRow + 5}`);
    this.addText(
      `C${startRow + 5}`,
      'The thematic classification of the metric, such as Emissions, Water & Effluents, Human Health, Resource Use, or Biodiversity. This helps group and interpret metrics within broader environmental or social impact domains.',
    );
    this.addBorders(`C${startRow + 5}`);
    this.addText(`B${startRow + 6}`, 'Overall Value');
    this.addBorders(`B${startRow + 6}`);
    this.addText(
      `C${startRow + 6}`,
      'Vested Impact takes the revenue or expenditure data provided for each activity and multiplies It by the equivalent factor (based on USEEIO and/or adjusted country specific tables and factors). Where Overall Value = (revenue/expenditure) X (spend based factor)',
    );
    this.addBorders(`C${startRow + 6}`);
    this.addText(`B${startRow + 7}`, 'Units');
    this.addBorders(`B${startRow + 7}`);
    this.addText(
      `C${startRow + 7}`,
      'The unit of measurement for the metric (e.g., kg CO₂e, m³ water, MJ energy), allowing consistent comparison and integration into impact assessments and regulatory disclosures.',
    );
    this.addBorders(`C${startRow + 7}`);
  }

  private esgAssessment(startRow: number) {
    this.addText(`B${startRow}`, 'ESG assessment', { bold: true, underline: true, wrapText: false });
    this.addText(`B${startRow + 1}`, 'Field Name', { bold: true });
    this.addBorders(`B${startRow + 1}`);
    this.setColor(`B${startRow + 1}`, 'FFCCCCCC');
    this.addText(`C${startRow + 1}`, 'Description', { bold: true });
    this.addBorders(`C${startRow + 1}`);
    this.setColor(`C${startRow + 1}`, 'FFCCCCCC');
    this.addText(`B${startRow + 2}`, 'ESG Assessment', { bold: true });
    this.addBorders(`B${startRow + 2}`);
    this.setColor(`B${startRow + 2}`, 'FFEEEEEE');
    this.addText(
      `C${startRow + 2}`,
      "Vested Impact's impact materiality assessments align with EU CSRD, CSDDD, SFDR, and GRI frameworks. The platform provides structured, quantitative outputs that cover upstream and downstream risks, supply chain emissions, and impact disclosures required for regulatory reporting, replacing qualitative ESG with science-backed, double materiality-aligned data.",
    );
    this.addBorders(`C${startRow + 2}`);
    this.setColor(`C${startRow + 2}`, 'FFEEEEEE');
    this.addText(`B${startRow + 3}`, 'Impact Materiality');
    this.addBorders(`B${startRow + 3}`);
    this.addText(
      `C${startRow + 3}`,
      'Impact materiality assesses the effects of an asset on the environment and society. Vested Impact measures these effects against the UN Sustainable Development Goals and their 169 targets, which are globally-recognized environmental and societal priorities in areas including poverty, inequality, climate change, environmental degradation, peace and justice. Impact materiality covers a business’s whole supply chain, including the business’s activities as well as companies both upstream and downstream on the supply chain, and not limited to direct contractual relationships. Negative material impacts are classed as “risks”, while positive ones are classed as “opportunities.” Whether an impact is “material” is assessed based on the likelihood of the impact occurring, and the impact’s scale (how grave or serious is the impact?) and scope (how widely is the impact felt?). In Vested Impact the Impact Materiality rating is calcuated as the Overall Impact Score',
    );
    this.addBorders(`C${startRow + 3}`);
    this.addText(`B${startRow + 4}`, 'Financial Materiality');
    this.addBorders(`B${startRow + 4}`);
    this.addText(
      `C${startRow + 4}`,
      "Financial materiality, under the EU CSRD, refers to sustainability-related risks and opportunities that could reasonably affect a company's financial performance, position, or value creation over the short, medium, or long term. Vested Impact identifies and quantifies financial materiality by flagging how external risks and impacts, alongside regulatory shifts or environmental pressures, may financially affect SMEs and private companies. Each risk/opportunity is assigned a Financial Materiality status that could be any value:",
    );
    this.addBorders(`C${startRow + 4}`);
    this.addText(`B${startRow + 5}`, 'ESRS Topic');
    this.addBorders(`B${startRow + 5}`);
    this.addText(
      `C${startRow + 5}`,
      'ESRS (European Sustainability Reporting Standards) references related to the activity.',
    );
    this.addBorders(`C${startRow + 5}`);
    this.addText(`B${startRow + 6}`, 'SDG Target & Other Affected Frameworks');
    this.addBorders(`B${startRow + 6}`);
    this.addText(
      `C${startRow + 6}`,
      'Contains the relevant UN SDG Target, as well as any other key frameworks or metrics that align with the ESRS Topic (such as GRI, EU Taxonomy, SFDR etc).',
    );
    this.addBorders(`C${startRow + 6}`);
    this.addText(`B${startRow + 7}`, 'Business Activity');
    this.addBorders(`B${startRow + 7}`);
    this.addText(`C${startRow + 7}`, 'The identified key product, service, activity and/or intervention of the asset');
    this.addBorders(`C${startRow + 7}`);
    this.addText(`B${startRow + 8}`, 'Country');
    this.addBorders(`B${startRow + 8}`);
    this.addText(
      `C${startRow + 8}`,
      'The geographic context in which the impact occurs. Vested Impact localizes SDG outcomes and ESRS topics to the specific country where the product, service, or activity is delivered—accounting for local needs, development levels, and progress toward the target.',
    );
    this.addBorders(`C${startRow + 8}`);
    this.addText(`B${startRow + 9}`, 'Actual / Potential');
    this.addBorders(`B${startRow + 9}`);
    this.addText(
      `C${startRow + 9}`,
      'Vested Impact indicates whether a risk/opportunity is "Actual" or "Potential" where Actual is either proven or known with a significant degree of confidence to occur. Any cases where realisation of the risk/opportunity is relativeyl unknown, is marked "Potential"',
    );
    this.addBorders(`C${startRow + 9}`);
    this.addText(`B${startRow + 10}`, 'Type Of Impact');
    this.addBorders(`B${startRow + 10}`);
    this.addText(
      `C${startRow + 10}`,
      'The depth score indicates how directly an impact, positive or negative, imapacts on the SDG target. The categories are derived from the OECD Guidelines for MND\'s and range from "caused by", "contributes to", "directly linked" and  (derived from AI and human-generated summaries from academic references, as cited).\nAllowed values:\n- Caused by\n- Contributes to\n- Directly linked\n- Indirectly linked',
    );
    this.addBorders(`C${startRow + 10}`);
    this.addText(`B${startRow + 11}`, 'Value Chain');
    this.addBorders(`B${startRow + 11}`);
    this.addText(
      `C${startRow + 11}`,
      'Indication of where in the value chain the risk or impact is anticipated to occur\nAllowed values (multiple allowed, colon separated):\n- Raw Input Material\n- Transport & Logistics\n- Supplier Operations\n- Manufacturing/Processing\n- Operations and facilities\n- Employee Practices\n- Financed Asset\n- Distribution & Logistics\n- Product/service use\n- Customer Experience\n- End-of-life disposal',
    );
    this.addBorders(`C${startRow + 11}`);
    this.addText(`B${startRow + 12}`, 'Affected Stakeholders');
    this.addBorders(`B${startRow + 12}`);
    this.addText(
      `C${startRow + 12}`,
      'The stakeholder/s impacted by this risk or impact\nAllowed values (multiple allowed, colon separated):\n- Nature\n- Employees\n- Communities (living or working near operating sites)\n- Shareholders\n- Customers and consumers\n- Suppliers\n- Business partners\n- Regulators\n- Policy makers & government\n- Trade unions\n- Vulnerable and/or minority groups',
    );
    this.addBorders(`C${startRow + 12}`);
    this.addText(`B${startRow + 13}`, 'Time Horizon');
    this.addBorders(`B${startRow + 13}`);
    this.addText(
      `C${startRow + 13}`,
      'Indication of the time horizon for which the risk or impact will become material\nAllowed values:\n- Less than 1 year and/or immediate\n- 1 to 3 years\n- 3 to 5 years\n- More than 5 years',
    );
    this.addBorders(`C${startRow + 13}`);
    this.addText(`B${startRow + 14}`, 'Likelihood');
    this.addBorders(`B${startRow + 14}`);
    this.addText(
      `C${startRow + 14}`,
      'Likelihood refers to the probability or chance that a sustainability-related impact, risk, or opportunity will materialise. Under the CSRD, likelihood is used to assess how probable it is that a company’s activities will result in actual or potential impacts on people or the environment—informing the materiality assessment alongside the severity or scale of the impact.',
    );
    this.addBorders(`C${startRow + 14}`);
    this.addText(`B${startRow + 15}`, 'Type Of Risk');
    this.addBorders(`B${startRow + 15}`);
    this.addText(
      `C${startRow + 15}`,
      'The type of risk.\nAllowed values (multiple allowed, colon separated):\n- Regulatory\n- Policy & Legal\n- Market\n- Reputational\n- Physical\n- Transition\n- Financial\n- Operational\n- Human Rights\n- Resource',
    );
    this.addBorders(`C${startRow + 15}`);
    this.addText(`B${startRow + 16}`, 'Scale');
    this.addBorders(`B${startRow + 16}`);
    this.addText(
      `C${startRow + 16}`,
      'In line with CSRD, Scale refers to the gravity or severity of a sustainability-related impact — specifically, the extent to which it infringes on fundamental human rights, access to basic needs, or critical environmental thresholds (e.g., education, health, livelihood, biodiversity).\nUnder Vested Impact’s methodology, Scale is quantitatively derived as the average of the Need Score and Importance Score, reflecting both how urgent the impact is in the affected context and how critical the impacted issue is to people or ecosystems.',
    );
    this.addBorders(`C${startRow + 16}`);
    this.addText(`B${startRow + 17}`, 'Scope');
    this.addBorders(`B${startRow + 17}`);
    this.addText(
      `C${startRow + 17}`,
      'Per CSRD, Scope describes the extent or reach of an impact — that is, how widespread it is across populations or ecosystems (e.g., the number of people affected or the scale of environmental degradation).\nIn Vested Impact’s methodology, Scope is calculated as the average of the Effect Score and Value Score, capturing both the measurable breadth of the impact and the degree to which the activity directly contributes to or mitigates the issue.',
    );
    this.addBorders(`C${startRow + 17}`);
    this.addText(`B${startRow + 18}`, 'Affected Financial Item');
    this.addBorders(`B${startRow + 18}`);
    this.addText(
      `C${startRow + 18}`,
      'The type of financial impact/risk\nAllowed values:\n- Cost of goods\n- Operations costs\n- Capital expenditure\n- Revenue\n- Stranded/depreciated asset\n- Liabilities and provisions\n- Insurance\n- Access to and cost of capital\n- Regulatory and compliance costs\n- R&D and technology costs',
    );
    this.addBorders(`C${startRow + 18}`);
    this.addText(`B${startRow + 19}`, 'Financial Materiality Note');
    this.addBorders(`B${startRow + 19}`);
    this.addText(`C${startRow + 19}`, 'A descriptive note describing the financial materiality');
    this.addBorders(`C${startRow + 19}`);
    this.addText(`B${startRow + 20}`, 'Impact Notes');
    this.addBorders(`B${startRow + 20}`);
    this.addText(
      `C${startRow + 20}`,
      'Notes provide an AI-generated summary of key academic findings drawn from the underlying References linked to the specific SDG Target and activity. This section synthesises the most relevant evidence from peer-reviewed literature, offering context and justification for the impact assessment in a clear and accessible format.',
    );
    this.addBorders(`C${startRow + 20}`);
  }

  populate() {
    this.addBlankRows(1000, 1);
    this.getColumn('B').width = 40;
    this.getColumn('C').width = 80;
    this.addText('B2', 'Vested Impact Glossary & Terms', { bold: true, size: 12, wrapText: false });
    this.activityImpact(4);
    this.esgAssessment(59);
    this.environmentalMetrics(81);
  }
}
