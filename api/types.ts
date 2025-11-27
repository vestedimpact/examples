type Activity = {
  id: number;
  industry: string;
  name: string;
};

export type Activities = {
  activities: Activity[];
};

export type Asset = {
  breakdown: {
    activity: string;
    activityId: number;
    country: string;
    countryCode: string;
    weight: number;
  }[];
  description: string;
  hqCountryCode: string;
  id: string;
  industry: string;
  name: string;
  numEmployees: number;
};

export type Opportunity = {
  activity: string;
  additionalFrameworks: string | null;
  affectedFinancialItem: string;
  country: string;
  description: string;
  esrsReference: string;
  financialMateriality: string;
  financialMaterialityNote: string;
  impactScore: number;
  impactStatus: string;
  impactType: string;
  isActual: boolean;
  likelihood: string;
  riskCategory: string;
  riskType: string;
  scale: string;
  scope: string;
  sdgReference: string;
  stakeholder: string;
  timeHorizon: string;
  valueChain: string;
  weight: number;
  weightedImpactScore: number;
};

export type Risk = Opportunity & {
  irremediability: string;
};

export type AssetBasics = {
  currency: string;
  description: string;
  hqCountryCode: string;
  industry: string;
  name: string;
  numEmployees: number;
  revenue: number;
  revenueGrowth: number;
};

export type AssetBreakdown = {
  breakdown: {
    activityId: number;
    countryCode: string;
    weight: number;
  }[];
};

export type AssetCreateInput = Omit<AssetBasics, 'currency' | 'revenue' | 'revenueGrowth'>;

export type AssetCreateResponse = {
  asset: Asset;
  suggestedActivities: Activity[];
};

export type AssetImpact = {
  activityImpacts: {
    activity: string;
    negativeImpact: number;
    netImpact: number;
    positiveImpact: number;
    weight: number;
  }[];
  assetBreakdown: {
    activity: string;
    country: string;
    countryCode: string;
    weight: number;
  }[];
  countryImpacts: {
    country: string;
    countryCode: string;
    negativeImpact: number;
    netImpact: number;
    positiveImpact: number;
    weight: number;
  }[];
  impactBreakdown: {
    activity: string;
    benchmarks: {
      countries: {
        absoluteChange: number;
        assetGrowth: number;
        country: string;
        paceOfChange: number;
        reportedYear: number;
        trend: number;
      }[];
      description: string;
      indicator: string;
      justificationNote: string;
      reference: string;
      sdgTarget: string;
      target: number;
    }[];
    countries: {
      country: string;
      countryCode: string;
      negativeImpact: number;
      netImpact: number;
      positiveImpact: number;
      weight: number;
    }[];
    flaggedOpportunities: {
      note: string;
      source: string;
      status: string;
      targets: {
        country: string;
        sdgTarget: string;
      }[];
      type: string;
    }[];
    flaggedRisks: {
      note: string;
      source: string;
      status: string;
      targets: {
        country: string;
        sdgTarget: string;
      }[];
      type: string;
    }[];
    indicators: {
      country: string;
      description: string;
      indicator: string;
      source: string;
      sdgTarget: string;
      trend?: number;
    }[];
    metrics: {
      category: string;
      countries: {
        country: string;
        units: string;
        value?: number;
      }[];
      description: string;
      method: string;
      name: string;
      units: string;
      value?: number;
    }[];
    negativeImpact: number;
    netImpact: number;
    positiveImpact: number;
    references: {
      reference: string;
      targets: {
        sdgTarget: string;
        standardOfEvidence?: number;
      }[];
      url?: string;
    }[];
    targets: {
      negativeImpact: number;
      note?: string;
      positiveImpact: number;
      sdgTarget: string;
      subScores: {
        contribution?: {
          change?: {
            status: string;
            score: number;
          };
          scale?: {
            status: string;
            score: number;
          };
          status: string;
          score: number;
        };
        country: string;
        impactScore: number;
        impactStatus: string;
        importance: {
          global: {
            status: string;
            score: number;
          };
          local: {
            status: string;
            score: number;
          };
          status: string;
          score: number;
          supporting: {
            status: string;
            score: number;
          };
        };
        need: {
          countryClassification: {
            status: string;
            score: number;
          };
          countryIncome: {
            status: string;
            score: number;
          };
          countryScore: number;
          countryVulnerability: {
            status: string;
            score: number;
          };
          sdgStatus: {
            status: string;
            score: number;
          };
          sdgTrend: {
            status: string;
            score: number;
          };
          status: string;
          score: number;
        };
        value: {
          depth: {
            status: string;
            score: number;
          };
          immediacy: {
            status: string;
            score: number;
          };
          irremediability: {
            status: string;
            score: number;
          };
          status: string;
          score: number;
          sustained: {
            status: string;
            score: number;
          };
        };
      }[];
    }[];
    weight: number;
  }[];
  industry: string;
  negativeImpact: number;
  opportunities: {
    E: Opportunity[];
    G: Opportunity[];
    S: Opportunity[];
  };
  overallMetrics: {
    category: string;
    description: string;
    method: string;
    name: string;
    units: string;
    value?: number;
  }[];
  positiveImpact: number;
  referenceModels: {
    description: string;
    indicator: string;
    source: string;
  }[];
  reportDate: string;
  risks: {
    E: Risk[];
    G: Risk[];
    S: Risk[];
  };
  sdgImpacts: {
    negativeImpact: number;
    netImpact: number;
    positiveImpact: number;
    sdgGoal: string;
    targetImpacts: {
      negativeImpact: number;
      positiveImpact: number;
      sdgTarget: string;
    }[];
  }[];
  vestedImpactRating: string;
  vestedImpactScore: number;
  weight: number;
};

export type AssetImpactHistory = {
  reports: {
    id: string;
    negativeImpact: number;
    positiveImpact: number;
    reportDate: string;
    vestedImpactRating: string;
    vestedImpactScore: number;
  }[];
};

export type AssetSearchResults = {
  results: {
    id: string;
    industry: string;
    name: string;
  }[];
  query: string;
};

type GroupHolding = {
  id: string;
  name: string;
  type: 'ASSET' | 'GROUP';
  weight: number;
};

export type Group = {
  description?: string;
  holdings: GroupHolding[];
  id: string;
  name: string;
  owner: string;
};

export type GroupCreateInput = {
  description: string;
  name: string;
  owner: string;
};

export type GroupHoldings = {
  holdings: Omit<GroupHolding, 'name'>[];
};

type GroupHoldingImpact = GroupHolding & {
  negativeImpact: number;
  positiveImpact: number;
};

type GroupHoldingImpactWithHoldingWeight = GroupHoldingImpact & {
  holdingWeight: number;
};

type GroupMaterialityMatrix = {
  veryHigh: {
    veryLarge: number;
    large: number;
    moderate: number;
    small: number;
  };
  high: {
    veryLarge: number;
    large: number;
    moderate: number;
    small: number;
  };
  medium: {
    veryLarge: number;
    large: number;
    moderate: number;
    small: number;
  };
  low: {
    veryLarge: number;
    large: number;
    moderate: number;
    small: number;
  };
  veryLow: {
    veryLarge: number;
    large: number;
    moderate: number;
    small: number;
  };
};

export type GroupImpact = {
  esgSummary: {
    E: {
      opportunities: {
        holdings: (GroupHolding & { materiality: GroupMaterialityMatrix })[];
        materiality: GroupMaterialityMatrix;
      };
      risks: {
        holdings: (GroupHolding & { materiality: GroupMaterialityMatrix })[];
        materiality: GroupMaterialityMatrix;
      };
    };
    S: {
      opportunities: {
        holdings: (GroupHolding & { materiality: GroupMaterialityMatrix })[];
        materiality: GroupMaterialityMatrix;
      };
      risks: {
        holdings: (GroupHolding & { materiality: GroupMaterialityMatrix })[];
        materiality: GroupMaterialityMatrix;
      };
    };
    G: {
      opportunities: {
        holdings: (GroupHolding & { materiality: GroupMaterialityMatrix })[];
        materiality: GroupMaterialityMatrix;
      };
      risks: {
        holdings: (GroupHolding & { materiality: GroupMaterialityMatrix })[];
        materiality: GroupMaterialityMatrix;
      };
    };
  };
  flaggedOpportunities: {
    holdings: (GroupHolding & { count: number })[];
    type: string;
  }[];
  flaggedRisks: {
    holdings: (GroupHolding & { count: number })[];
    type: string;
  }[];
  holdingImpacts: (GroupHoldingImpact & {
    vestedImpactRating: string;
    vestedImpactScore: number;
  })[];
  industryImpacts: {
    activities: {
      activity: string;
      holdings: GroupHoldingImpactWithHoldingWeight[];
      negativeImpact: number;
      positiveImpact: number;
      weight: number;
    }[];
    holdings: GroupHoldingImpactWithHoldingWeight[];
    industry: string;
    negativeImpact: number;
    positiveImpact: number;
    weight: number;
  }[];
  metrics: {
    category: string;
    description: string;
    holdings: (GroupHolding & { value?: number })[];
    method: string;
    name: string;
    units: string;
  }[];
  negativeImpact: number;
  positiveImpact: number;
  regionImpacts: {
    countries: {
      country: string;
      countryCode: string;
      holdings: GroupHoldingImpactWithHoldingWeight[];
      negativeImpact: number;
      positiveImpact: number;
      weight: number;
    }[];
    holdings: GroupHoldingImpactWithHoldingWeight[];
    region: string;
    negativeImpact: number;
    positiveImpact: number;
    weight: number;
  }[];
  reportDate: string;
  sdgImpacts: {
    holdings: GroupHoldingImpact[];
    negativeImpact: number;
    positiveImpact: number;
    sdgGoal: string;
    targets: {
      holdings: GroupHoldingImpact[];
      negativeImpact: number;
      positiveImpact: number;
      sdgTarget: string;
      weight: number;
    }[];
    weight: number;
  }[];
  totalWeight: number;
  vestedImpactRating: string;
  vestedImpactScore: number;
};

export type GroupImpactHistory = {
  reports: {
    id: string;
    negativeImpact: number;
    numHoldings: number;
    positiveImpact: number;
    reportDate: string;
    vestedImpactRating: string;
    vestedImpactScore: number;
  }[];
};

export type GroupSearchResults = {
  results: {
    id: string;
    name: string;
    owner: string;
  }[];
  query: string;
};

type Country = {
  code: string;
  name: string;
};

export type Countries = {
  countries: Country[];
};

export type Currencies = {
  currencies: {
    code: string;
    name: string;
    symbol: string;
  }[];
};

export type ImpactCalculationStatus = {
  status: 'UNKNOWN' | 'ACTIVE' | 'FAILED' | 'PENDING' | 'COMPLETED';
};

export type Industries = {
  industries: string[];
};

export type Regions = {
  regions: {
    countries: Country[];
    name: string;
  }[];
};
