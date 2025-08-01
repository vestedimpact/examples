# <img src="https://api.vestedimpact.co.uk/logos/purple.svg" alt="Vested Impact logo" width="30" /> Vested Impact Data API Examples
<br />

> Example demonstrating the ways in which the Vested Impact Data API can be used and examples of how Vested Impact export data for end user consumption.

The Data API Swagger docs can be found [here](https://api.vestedimpact.co.uk/docs).

To use the Vested Impact Data API you must:
- Have an active 'pro' account with Vested Impact.
- Log in to the Vested Impact SaaS to obtain your API key.

----

### Repository Contents

| Folder         | Overview |
|----------------|----------|
| `api`          | Example classes which connect to the Vested Impact Data API using fetch. They demonstrate common use cases of the API but do not contain error handling and are not production ready. |
| `excel-output` | Code which utilises [ExcelJS](https://www.npmjs.com/package/exceljs) to output data from the API into Excel workbooks which can be consumed by users. |
| `utils`        | Utility code we think could be useful to you. Including ESRS and UN SDG utilities. |
