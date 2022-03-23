import pandas as pd
import os

df = pd.read_excel(os.path.join(os.getcwd(), '2020_input.xlsx'))

# Only looking at the blue(ish) counties (for now)
counties_included = ['St. Louis County', 'St. Louis City', 'Jefferson', 'St. Charles', 'Boone', 'Jackson', 'Clay', 'Greene', 'Platte']

# Replacing relevant counties with their county code
# df.loc[df.County == 'St. Louis County'] = '189'
# df.loc[df.County == 'St. Louis City'] = '510'
# df.loc[df.County == 'Jefferson'] = '099'
# df.loc[df.County == 'St. Charles'] = '183'
# df.loc[df.County == 'Boone'] = '019'
# df.loc[df.County == 'Clay'] = '047'




for ind in df.index:
    if str(df['County'][ind]) == 'St. Louis County':
        df['County'][ind] = '189'
    if str(df['County'][ind]) == 'St. Louis City':
        df['County'][ind] = '510'
    if str(df['County'][ind]) == 'Jefferson':
        df['County'][ind] = '099'
    if str(df['County'][ind]) == 'St. Charles':
        df['County'][ind] = '183'
    if str(df['County'][ind]) == 'Boone':
        df['County'][ind] = '019'
    if str(df['County'][ind]) == 'Clay':
        df['County'][ind] = '047'

    else:
        df.drop(ind)


    # Change the first two column names so they align with the .dbf file
df.rename(columns={"County": "COUNTYFP,C,3", "Precinct": "NAME,C,100"})

df.to_excel("2020_output.xlsx", sheet_name='updated_counties')





""" 

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

dfoutput.to_excel("2020output.xlsx", sheet_name='Sheet_name_1') """
