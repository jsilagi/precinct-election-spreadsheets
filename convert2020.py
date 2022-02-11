import pandas as pd
import os
from helper import zDiv

df = pd.read_excel(os.path.join(os.getcwd(), '2020PrecinctDataForPython.xlsx'))

column_names = ['County', 'Precinct', 'President (pct dem)', 'U.S. Representative (pct dem)', 'Governor (pct dem)', 'Lt. Governor (pct dem)', 'Sec. of State (pct dem)', 'Treasurer (pct dem)', 'AG (pct dem)', 'State Senator (pct dem)', 'State Representative (pct dem)', 'Amendment 1 (pct yes)', 'Amendment 3 (pct yes)']

# Initialize counters
demPres = totalPres = demUSRep = totalUSRep = demGov = totalGov = demLTGov = totalLTGov = demSoS = totalSoS = demTres = totalTres = demAG = totalAG = demStateSen = totalStateSen = demStateRep = totalStateRep = YAmen1 = NAmen1 = YAmen3 = NAmen3 = 0

# Make our master array that will be added to ouput
masterArray = []


# Loop through all rows
visited_precincts = ['7']   # the first precinct
visited_counties = ['Adair'] # the first county (for ABSENTEE only)
for ind in df.index:

    # Special case handling for ABSENTEE precinct
    if "ABSENTEE" == str(df['PrecinctName'][ind]):
        if str(df['CountyName'][ind]) not in visited_counties:
            # Finalize the values of the previous county
            masterArray.append([str(df['CountyName'][ind-1]), str(df['PrecinctName'][ind-1]), zDiv(demPres, totalPres), zDiv(demUSRep, totalUSRep), zDiv(demGov, totalGov), zDiv(demLTGov, totalLTGov), zDiv(demSoS, totalSoS), zDiv(demTres, totalTres), zDiv(demAG, totalAG), zDiv(demStateSen, totalStateSen), zDiv(demStateRep, totalStateRep), zDiv(YAmen1, (YAmen1+NAmen1)), zDiv(YAmen3, (YAmen3+NAmen3))])
            # Reset counters
            demPres = totalPres = demUSRep = totalUSRep = demGov = totalGov = demLTGov = totalLTGov = demSoS = totalSoS = demTres = totalTres = demAG = totalAG = demStateSen = totalStateSen = demStateRep = totalStateRep = YAmen1 = NAmen1 = YAmen3 = NAmen3 = 0
            # Mark this new county as visited
            visited_counties.append(str(df['CountyName'][ind]))


    # If we have reached a new precinct
    if str(df['PrecinctName'][ind]) not in visited_precincts:

        # Finalize the values of the previous precinct
        masterArray.append([str(df['CountyName'][ind-1]), str(df['PrecinctName'][ind-1]), zDiv(demPres, totalPres), zDiv(demUSRep, totalUSRep), zDiv(demGov, totalGov), zDiv(demLTGov, totalLTGov), zDiv(demSoS, totalSoS), zDiv(demTres, totalTres), zDiv(demAG, totalAG), zDiv(demStateSen, totalStateSen), zDiv(demStateRep, totalStateRep), zDiv(YAmen1, (YAmen1+NAmen1)), zDiv(YAmen3, (YAmen3+NAmen3))])

        # Reset counters
        demPres = totalPres = demUSRep = totalUSRep = demGov = totalGov = demLTGov = totalLTGov = demSoS = totalSoS = demTres = totalTres = demAG = totalAG = demStateSen = totalStateSen = demStateRep = totalStateRep = YAmen1 = NAmen1 = YAmen3 = NAmen3 = 0

        # Mark this new precinct as visited
        visited_precincts.append(str(df['PrecinctName'][ind]))


    # President
    if "U.S. President and Vice President" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demPres += df['YES'][ind]
            totalPres += df['YES'][ind]
        else:
            totalPres += df['YES'][ind]

    # U.S. Representative
    if "U.S. Representative" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        else:
            totalUSRep += df['YES'][ind]

    # Governor
    if ("Governor" in df['OfficeTitle'][ind]) and ("Lieutenant" not in df['OfficeTitle'][ind]):
        if "DEM" in str(df['Party'][ind]):
            demGov += df['YES'][ind]
            totalGov += df['YES'][ind]
        else:
            totalGov += df['YES'][ind]

    # Lt. Governor
    if "Lieutenant Governor" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demLTGov += df['YES'][ind]
            totalLTGov += df['YES'][ind]
        else:
            totalLTGov += df['YES'][ind]

    # Sec. of State
    if "Secretary of State" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demSoS += df['YES'][ind]
            totalSoS += df['YES'][ind]
        else:
            totalSoS += df['YES'][ind]

    # Treasurer
    if "Treasurer" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demTres += df['YES'][ind]
            totalTres += df['YES'][ind]
        else:
            totalTres += df['YES'][ind]

    # AG
    if "Attorney General" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demAG += df['YES'][ind]
            totalAG += df['YES'][ind]
        else:
            totalAG += df['YES'][ind]

    # State Senator
    if "State Senator" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        else:
            totalStateSen += df['YES'][ind]

    # State Representative
    if "State Representative" in df['OfficeTitle'][ind]:
        if "DEM" in str(df['Party'][ind]):
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        else:
            totalStateRep += df['YES'][ind]

    # Amendment 1
    if "Constitutional Amendment 1" in df["OfficeTitle"][ind]:
        YAmen1 += df['YES'][ind]
        NAmen1 += df['NO'][ind]

    # Amendment 3
    if "Constitutional Amendment 3" in df["OfficeTitle"][ind]:
        YAmen3 += df['YES'][ind]
        NAmen3 += df['NO'][ind]


# Finalize the values of the final precinct
masterArray.append([str(df['CountyName'][ind-1]), str(df['PrecinctName'][ind-1]), zDiv(demPres, totalPres), zDiv(demUSRep, totalUSRep), zDiv(demGov, totalGov), zDiv(demLTGov, totalLTGov), zDiv(demSoS, totalSoS), zDiv(demTres, totalTres), zDiv(demAG, totalAG), zDiv(demStateSen, totalStateSen), zDiv(demStateRep, totalStateRep), zDiv(YAmen1, (YAmen1+NAmen1)), zDiv(YAmen3, (YAmen3+NAmen3))])

# Make a new dataframe and add our data to it
dfoutput = pd.DataFrame(columns = column_names)
for entry in masterArray:
    dfoutput = dfoutput.append([entry])

dfoutput.to_excel("2020output.xlsx", sheet_name='Sheet_name_1')