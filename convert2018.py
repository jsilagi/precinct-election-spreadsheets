import pandas as pd
import os
from helper import zDiv

df = pd.read_excel(os.path.join(os.getcwd(), '2018PrecinctDataForPython.xlsx'))

#print(df)

# Make an empty dataframe with the column names
""" column_names = []
for name in list(df.OfficeTitle.unique()):
    if "State Representative" in str(name):
        name = "State Representative"
    if "State Senator" in str(name):
        name = "State Senator"
    if "Circuit Judge" in str(name):
        name = "Circuit Judge"
    if "U.S. Representative" in str(name):
        name = "U.S. Representative"
    if "U. S. Representative" in str(name):
        name = "U.S. Representative"
    if "U. S. Senator" in str(name):
        name = "U.S. Senator"
    if "U.S.  Senator" in str(name):
        name = "U.S. Senator"
    name = str(name).strip()    
    column_names.append(name)
    print(name)
column_names = list(set(column_names)) 
column_names.append('County')
dfoutput = pd.DataFrame(columns = column_names) """

column_names = ['County', 'Precinct', 'U.S. Senator (pct dem)', 'U.S. Representative (pct dem)', 'State Senator (pct dem)', 'State Representative (pct dem)', 'State Auditor (pct dem)', 'Prop B (pct yes)', 'Prop C (pct yes)', 'Prop D (pct yes)', 'Amendment 1 (pct yes)', 'Amendment 2 (pct yes)', 'Amendment 3 (pct yes)', 'Amendment 4 (pct yes)']
#dfoutput = pd.DataFrame(columns = column_names)


# Initialize counters
demUSSen = totalUSSen = demUSRep = totalUSRep = demStateSen = totalStateSen = demStateRep = totalStateRep = demStateAud = totalStateAud = YPropB = NPropB = YPropC = NPropC = YPropD = NPropD = YAmen1 =NAmen1 = YAmen2 = NAmen2 = YAmen3 = NAmen3 = YAmen4 = NAmen4 = 0

# Make our master array that will be added to ouput(?)
masterArray = []

# Loop through all rows
visited_precincts = ['7']   # the first precinct
for ind in df.index:

    ## TODO HAVE A CATCH FOR ABSENTEE PRECINCT
    if "ABSENTEE" in str(df['PrecinctName'][ind]):
        continue

    # If we have reached a new precinct
    if str(df['PrecinctName'][ind]) not in visited_precincts:
        #print(df['PrecinctName'][ind])

        # Finalize the values of the previous precinct
        masterArray.append([str(df['CountyName'][ind-1]), str(df['PrecinctName'][ind-1]), zDiv(demUSSen, totalUSSen), zDiv(demUSRep, totalUSRep), zDiv(demStateSen, totalStateSen), zDiv(demStateRep, totalStateRep), zDiv(demStateAud, totalStateAud), zDiv(YPropB, (YPropB+NPropB)), zDiv(YPropC, (YPropC+NPropC)), zDiv(YPropD, (YPropD+NPropD)), zDiv(YAmen1, (YAmen1+NAmen1)), zDiv(YAmen2, (YAmen2+NAmen2)), zDiv(YAmen3, (YAmen3+NAmen3)), zDiv(YAmen4, (YAmen4+NAmen4))])

        # Reset counters if we haven't looked at this precinct yet
        demUSSen = totalUSSen = demUSRep = totalUSRep = demStateSen = totalStateSen = demStateRep = totalStateRep = demStateAud = totalStateAud = YPropB = NPropB = YPropC = NPropC = YPropD = NPropD = YAmen1 =NAmen1 = YAmen2 = NAmen2 = YAmen3 = NAmen3 = YAmen4 = NAmen4 = 0

        # Mark this new precinct as visited
        visited_precincts.append(str(df['PrecinctName'][ind]))


    # U.S. Senator
    if "U.S. Senator" in df['OfficeTitle'][ind]:
        if "Claire McCaskill" in df['CandidateName'][ind]:
            demUSSen += df['YES'][ind]
            totalUSSen += df['YES'][ind]
        else:
            totalUSSen += df['YES'][ind]
    elif "U. S. Senator" in df['OfficeTitle'][ind]:
        if "Claire McCaskill" in df['CandidateName'][ind]:
            demUSSen += df['YES'][ind]
            totalUSSen += df['YES'][ind]
        else:
            totalUSSen += df['YES'][ind]
    elif "U.S.  Senator" in df['OfficeTitle'][ind]:
        if "Claire McCaskill" in df['CandidateName'][ind]:
            demUSSen += df['YES'][ind]
            totalUSSen += df['YES'][ind]
        else:
            totalUSSen += df['YES'][ind]


    # U.S. Representative
    if "U.S. Representative" in df["OfficeTitle"][ind]:
        if "Lacy Clay" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Cort VanOstran" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Katy Geppert" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Renee Hoagenson" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Emanuel Cleaver, II" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Henry Robert Martin" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Jamie Daniel Schoolcraft" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Kathy Ellis" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        else:
            totalUSRep += df['YES'][ind]
    if "U. S. Representative" in df["OfficeTitle"][ind]:
        if "Lacy Clay" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Cort VanOstran" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Katy Geppert" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Renee Hoagenson" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Emanuel Cleaver, II" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Henry Robert Martin" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Jamie Daniel Schoolcraft" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        elif "Kathy Ellis" in df["CandidateName"][ind]:
            demUSRep += df['YES'][ind]
            totalUSRep += df['YES'][ind]
        else:
            totalUSRep += df['YES'][ind]


    # State Senator
    if "State Senator" in df["OfficeTitle"][ind]:
        if "Ayanna Shivers" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Terry Richard" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Terry Richard" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Brian Williams" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]    
        elif "Ryan Dillon" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Crystal Stephens" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]    
        elif "Jim Billedo" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Robert Butler" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Jill Schupp" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "John Kiehne" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Joe Poor" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Charlie Norr" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Carolyn McGowan" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Martin T. Rucker II" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Patrice Bilings" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Karla May" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Nicole Thompson" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        elif "Hillary Shields" in df["CandidateName"][ind]:
            demStateSen += df['YES'][ind]
            totalStateSen += df['YES'][ind]
        else:
            totalStateSen += df['YES'][ind]


    # State Representative
    if "State Representative" in df["OfficeTitle"][ind]:
        if "Paul Taylor" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Joni Perry" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Joni Perry" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Joe Frese" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Mitch Wrenn" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Dennis VanDyke" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Caleb McKnight" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Bob Bergland" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Shane R. Thompson" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Brady Lee O'Dell" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Sandy Van Wagner" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Mitch Weber" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Matt Sain" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jon Carpenter" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Tom Gorenc" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Mark Ellebracht" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Wes Rogers" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Ingrid Burnett" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jessica Merrick" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Robert Sauls" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Brandon Ellington" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Barbara Anne Washington" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Judy Morgan" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Greg Razer" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Ashley Bland Manlove" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Richard Brown" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jerome Barnes" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Rory Rowland" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Ryana Parks-Shaw" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Travis Hagewood" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Janice Brill" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Pat Williams" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "James P. (Jim) Ripley" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Keri Ingle" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "DaRon McGee" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Joe Runions" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Abby Zavos" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Rick Mellon" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "David A Beckham" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Joseph Widner" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jamie Blair" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Maren Bell Jones" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Kip Kendrick" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Martha Stevens" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Adrian Plank" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Raymond (Jeff) Faubion" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Lisa Buhr" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Michela Skelton" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Vince Lutterbie" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Dan Marshall" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Connie Simmons" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "James L. Williams" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Joan Shores" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Linda Ellen Greeson" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Sara Michael" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Pamela Menefee" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Ashley D. Fajkowski" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Janet Kester" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Shawn Finklein" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Bill Otto" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Tommie Pierson" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Alan K. Green" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jay Mosley" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Gretchen Bangert" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Paula Brown" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "LaDonna Appelbaum" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Doug Clemens" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Raychel Proudie" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Cora Faith Walker" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Alan Gray" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Chris Carter" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Steve Roberts" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Bruce Franks" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "LaKeySha Bosley" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Peter Merideth" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Steve Butz" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Donna Baringer" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Gina Mitten" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Wiley Price" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Kevin L. Windham, Jr." in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Maria N. Chappelle-Nadal" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Ian Mackey" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Tracy McCreery" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Kevin Fitzgerald" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Deb Lavender" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Sarah Unsicker" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Doug Beck" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Bob Burns" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jean Pretto" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Mike Walter" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Erica Hoffman" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Mike Revis" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Charles Triplett" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Mike LaBozzetta" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Helena Webb" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Genevieve Steidtmann" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "John F .Foster" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jim Klenc" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Peggy Sherwin" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Scott Cernicek" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jackie Sclair" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Curtis Wylde" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Betty Vining" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "James Cordrey" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Cody Kelley" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Phoebe Ottomeyer" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Benjamin Hagin" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Karen Settlemoir-Berg" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Dennis McDonald" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Bill Kraemer" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Kayla Chick" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Barbara Marco" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Marcie Nichols" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Theresa Schmitt" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Matt Heltz" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Joe Register" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Steve Dakopolos" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Chase Crawford" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jim Hogan" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Teri Hanna" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Rich Horton" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Ronna Ford" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Tyler Gunlock" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Nate Branscom" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Crystal Quade" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Cindy Slimp" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Derrick Nowlin" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Rob Bailey" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jeff Munzinger" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Raymond Lampert" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Cora Hanf" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Tony Smith" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Ronald G. Pember" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Gayla A. Dace" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Renita Green" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Bill Burlison" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Josh Rittenberry" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "J. T. (Jerry) Howard" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Robert L. Smith" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Matt Michel" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Loretta Thomas" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Jerry Sparks" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Angela R Thomas" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Elizabeth Lundstrum" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Sarah Hinkle" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        elif "Chad Fletcher" in df["CandidateName"][ind]:
            demStateRep += df['YES'][ind]
            totalStateRep += df['YES'][ind]
        else:
            totalStateRep += df['YES'][ind]


    # State Auditor
    if "State Auditor" in df["OfficeTitle"][ind]:
        if "Nicole Galloway" in df["CandidateName"][ind]:
            demStateAud += df['YES'][ind]
            totalStateAud += df['YES'][ind]
        else:
            totalStateAud += df['YES'][ind]



    # Proposition B
    if "Proposition B" in df["OfficeTitle"][ind]:
        YPropB += df['YES'][ind]
        NPropB += df['NO'][ind]

    # Proposition C
    if "Proposition C" in df["OfficeTitle"][ind]:
        YPropC += df['YES'][ind]
        NPropC += df['NO'][ind]

    # Proposition D
    if "Proposition D" in df["OfficeTitle"][ind]:
        YPropD += df['YES'][ind]
        NPropD += df['NO'][ind]

    # Constitutional Amendment No. 1
    if "Constitutional Amendment No. 1" in df["OfficeTitle"][ind]:
        YAmen1 += df['YES'][ind]
        NAmen1 += df['NO'][ind]

    # Constitutional Amendment No. 2
    if "Constitutional Amendment No. 2" in df["OfficeTitle"][ind]:
        YAmen2 += df['YES'][ind]
        NAmen2 += df['NO'][ind]

    # Constitutional Amendment No. 3
    if "Constitutional Amendment No. 3" in df["OfficeTitle"][ind]:
        YAmen3 += df['YES'][ind]
        NAmen3 += df['NO'][ind]

    # Constitutional Amendment No. 4
    if "Constitutional Amendment No. 4" in df["OfficeTitle"][ind]:
        YAmen4 += df['YES'][ind]
        NAmen4 += df['NO'][ind]


# Finalize the values of the final precinct
masterArray.append([str(df['CountyName'][ind-1]), str(df['PrecinctName'][ind-1]), zDiv(demUSSen, totalUSSen), zDiv(demUSRep, totalUSRep), zDiv(demStateSen, totalStateSen), zDiv(demStateRep, totalStateRep), zDiv(demStateAud, totalStateAud), zDiv(YPropB, (YPropB+NPropB)), zDiv(YPropC, (YPropC+NPropC)), zDiv(YPropD, (YPropD+NPropD)), zDiv(YAmen1, (YAmen1+NAmen1)), zDiv(YAmen2, (YAmen2+NAmen2)), zDiv(YAmen3, (YAmen3+NAmen3)), zDiv(YAmen4, (YAmen4+NAmen4))])
#print(masterArray)

# Make a new dataframe and add our data to it
dfoutput = pd.DataFrame(columns = column_names)
for entry in masterArray:
    #print(entry)
    #dfoutput = dfoutput.append(pd.DataFrame(entry, columns = column_names), ignore_index = True)
    dfoutput = dfoutput.append([entry])

dfoutput.to_excel("2018output.xlsx", sheet_name='Sheet_name_1')
