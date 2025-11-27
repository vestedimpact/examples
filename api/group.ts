import { Group, GroupCreateInput, GroupHoldings, GroupImpact, GroupImpactHistory, ImpactCalculationStatus } from './types';

const delay1Second = () => new Promise((resolve) => setTimeout(resolve, 1000));

/**
 * An example class demonstrating use cases for the Vested Impact data API using fetch.
 * This class focusses on group operations
 * 
 * @param apiKey Your Vested Impact API key.
 */
export class VestedImpactGroupAPI {
  private readonly apiUrl = 'https://api.vestedimpact.co.uk/v2/group';

  constructor(private readonly apiKey: string) { }

  /**
   * Create a new group.
   */
  async createGroup(group: GroupCreateInput) {
    const createResponse = await fetch(
      this.apiUrl,
      {
        body: JSON.stringify(group),
        headers: { 'api-key': this.apiKey, 'Content-Type': 'application/json' },
        method: 'POST',
      },
    );
    return (await createResponse.json()) as Group;
  }

  /**
   * Update group holdings and schedule impact calculation.
   */
  async generateGroupImpact(groupId: string, holdings: GroupHoldings) {
    /* Update group holdings */
    await fetch(
      `${this.apiUrl}/id/${groupId}/holdings`,
      {
        body: JSON.stringify(holdings),
        headers: { 'api-key': this.apiKey, 'Content-Type': 'application/json' },
        method: 'PUT',
      },
    );
    /* Schedule calculation of the groups impact data */
    await fetch(
      `${this.apiUrl}/id/${groupId}/impact/calculate`,
      { headers: { 'api-key': this.apiKey }, method: 'POST' },
    );
    /* Wait for the calculation to complete */
    let calculationStatus = 'UNKNOWN';
    while (!['COMPLETED', 'FAILED'].includes(calculationStatus)) {
      await delay1Second();
      const statusResponse = await fetch(
        `${this.apiUrl}/id/${groupId}/impact/status`,
        { headers: { 'api-key': this.apiKey }, method: 'GET' },
      );
      calculationStatus = ((await statusResponse.json()) as ImpactCalculationStatus).status;
    }
    if (calculationStatus !== 'COMPLETED') {
      /* Handle error in calculation, likely due to missing or misconfigured holdings data */
    }
  }

  /**
   * Fetch group data.
   */
  async getGroup(groupId: string) {
    const response = await fetch(
      `${this.apiUrl}/id/${groupId}`,
      { headers: { 'api-key': this.apiKey }, method: 'GET' },
    );
    return (await response.json()) as Group;
  }

  /**
   * Fetch impact history for a group.
   * This includes the id of each impact report which can be used to retrieve the full report data.
   */
  async getGroupImpactHistory(groupId: string) {
    const response = await fetch(
      `${this.apiUrl}/id/${groupId}/impact/history`,
      { headers: { 'api-key': this.apiKey }, method: 'GET' },
    );
    return (await response.json()) as GroupImpactHistory;
  }

  /**
   * Fetch impact data for a group.
   * If the reportId is provided a specific report from the groups history will be returned.
   * Else the most recent data will be returned.
   */
  async getGroupImpactReport(groupId: string, reportId?: string) {
    const response = await fetch(
      reportId ? `${this.apiUrl}/id/${groupId}/impact/id/${reportId}` : `${this.apiUrl}/id/${groupId}/impact/current`,
      { headers: { 'api-key': this.apiKey }, method: 'GET' },
    );
    return (await response.json()) as GroupImpact;
  }
}
