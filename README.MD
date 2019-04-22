Fraud Triangle Analysis
=======================

![This script was last tested in Nuix 7.8](https://img.shields.io/badge/Script%20Tested%20in%20Nuix-7.8-green.svg)

View the GitHub project [here](https://github.com/Nuix/Fraud-Triangle-Analysis) or download the latest release [here](https://github.com/Nuix/Fraud-Triangle-Analysis/releases).

# Overview

The [Fraud Triangle](https://en.wikipedia.org/wiki/Fraud_deterrence#Fraud_Triangle) is the concept that there are 3 common factors in communications regarding fraud:
1. Opportunity - The perceived ability to get away with fraud (weak safeguards, etc)
2. Pressure - The need to commit fraud (money needs, etc)
3. Rationalization - The individuals committing fraud rationalizing the act

This script assists in surfacing these types of commincations by searching for items which have matches against words or phrases present in several word lists.  The script comes with several words lists to get you started, but these are not exhaustive and it is recommended you enrich them with words relevant to your data.  The word lists included are:

- [Fraud Triangle Analysis - Opportunity](https://github.com/Nuix/Fraud-Triangle-Analysis/blob/master/Ruby/FraudTriangleAnalysis.nuixscript/Word%20Lists/Fraud%20Triangle%20Analysis%20-%20Opportunity.words)
- [Fraud Triangle Analysis - Pressure](https://github.com/Nuix/Fraud-Triangle-Analysis/blob/master/Ruby/FraudTriangleAnalysis.nuixscript/Word%20Lists/Fraud%20Triangle%20Analysis%20-%20Pressure.words)
- [Fraud Triangle Analysis - Rationalization](https://github.com/Nuix/Fraud-Triangle-Analysis/blob/master/Ruby/FraudTriangleAnalysis.nuixscript/Word%20Lists/Fraud%20Triangle%20Analysis%20-%20Rationalization.words)

The script generates a CSV with:

- Named Query
- Email Address
- Total Email Items with this Address (You can choose which address fields are searched)
- Total Emails with Opportunity Word List Hits
- Total Emails with Pressure Word List Hits
- Total Emails with Rationalization Word List Hits
- Percentage of Total Emails with Opportunity Word List Hits
- Percentage of Total Emails with Pressure Word List Hits
- Percentage of Total Emails with Rationalization Word List Hits

**Note:** `Named Query` column is only populated for rows generated from a named query and `Email Address` is only populated for rows generated from a named query.

Additionally the script can apply tags to items based upon which categories and email addresses had a match.  The tag name applied is specified with a tag name template that accepts the following placeholders:

| Placeholder | Resolves To |
|-------------|-------------|
| `{category}` | Resolves to relevant category name, i.e.: `Opportunity`, `Pressure` OR `Rationalization` |
| `{id}` |Resolves to relevant email address for items found by an email address and 'Name' for named queries. |

The default tag template is:

`Fraud Triangle|{category}|{id}`

Which might resolve to tags such as:

- `Fraud Triangle|Opportunity|Bob@Fake.biz`
- `Fraud Triangle|Pressure|Guy@IzFake.info`
- `Fraud Triangle|Rationalization|Lady@NotReel.uk`
- And so on...

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

This script requires that 3 words lists have been created/imported in Nuix:

- `Fraud Triangle Analysis - Opportunity`
- `Fraud Triangle Analysis - Pressure`
- `Fraud Triangle Analysis - Rationalization`

If any of these word lists cannot be located, the script will show a warning on startup and then close.  Releases of this script contain basic example word lists to get you started in the `Word Lists` sub-directory of the downloaded release archive.  They can be copied to one of the following locations to make them visible to Nuix:

| Location | Path |
|----------|------|
| `%appdata%\Nuix\Word Lists` | User level word list directory. |
| `%programdata%\Nuix\Word Lists` | System level word list directory. |
| `CASE_DIRECTORY\Stores\User Data\Word Lists` | Case level word list directory. |

# Cloning this Repository

This script relies on code from [Nx](https://github.com/Nuix/Nx) to present a settings dialog and progress dialog.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of Nx.jar by either:
1. Building it from [the source](https://github.com/Nuix/Nx)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/Nx/releases)

Once you have a copy of Nx.jar, make sure to include it in the same directory as the script.

# License

```
Copyright 2019 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```