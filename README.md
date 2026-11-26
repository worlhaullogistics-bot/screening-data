# screening-data

Daily copy of two **public** U.S. government lists, split into small lookup files for the WorldHaul / Team Express setup packages site (Cloudflare Worker, free plan).

- **Consolidated Screening List** (Commerce, State, Treasury): `data/csl/<first 4 letters of a distinctive word in the name>.json`
- **FMC Ocean Transportation Intermediaries** (active licensed NVOCCs / ocean freight forwarders, registered foreign NVOCCs): `data/fmc/id/<last 2 digits>.json`, `data/fmc/name/<abcd>.json`
- `data/meta.json`: when it was built and how many entries.

Refreshed every day by `.github/workflows/build.yml`. Contains public government data only; no customer, partner or company information.
Sources: data.trade.gov consolidated screening list; FMC OTI list (www2.fmc.gov/oti).
