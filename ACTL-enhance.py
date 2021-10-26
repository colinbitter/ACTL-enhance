import requests
import pandas as pd
import numpy as np
from pathlib import Path
import warnings
from datetime import date
d1 = date.today()
d1 = d1.strftime("%Y%m%d")

# folder path
downloads_path = str(Path.home() / "Downloads")
path1 = downloads_path

# Read ACTL spreadsheet
with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")
    ACTL = pd.read_excel(path1 + "/authorityList.xlsx", engine="openpyxl")
    AnaE = pd.read_excel(path1 + "/ACTLe.xlsx", engine="openpyxl")
    AnaP = pd.read_excel(path1 + "/ACTLp.xlsx", engine="openpyxl")
AnaP['Local Param 01'] = AnaP['Local Param 01'].apply(str)
AnaE['Local Param 01'] = AnaE['Local Param 01'].apply(str)
AnaP['OCLC Control Number (035a)'] = AnaP['OCLC Control Number (035a)'].apply(str)
AnaE['OCLC Control Number (035a)'] = AnaE['OCLC Control Number (035a)'].apply(str)
ACTL = ACTL.rename(columns={'MMS ID': 'MMS Id'})
ACTL = ACTL[ACTL.Vocabulary.str.contains('LCSH$|LCNAMES|LCGFT', na=False)]

# bib heading before column
ACTL['BIB Heading'] = ACTL['BIB Heading Before']

# strip trailing commas
ACTL['BIB Heading'] = ACTL['BIB Heading'].str.rstrip(',')
# strip trailing periods unless following initial
ACTL['BIB Heading'] = ACTL['BIB Heading'].str.replace(r'(?<!\s[A-Z])\.$', '', regex=True)
# strip trailing semicolons
ACTL['BIB Heading'] = ACTL['BIB Heading'].str.rstrip(' ;')
# strip series $x
ACTL['BIB Heading'] = ACTL['BIB Heading'].str.replace(r'\s\;\s\d{4}\-\d{4}$', '', regex=True)

# construct queries
ACTL.loc[ACTL['Vocabulary'] == 'LCSH', 'query'] = 'https://id.loc.gov/authorities/subject/label/'
ACTL.loc[ACTL['Vocabulary'] == 'LCNAMES', 'query'] = 'https://id.loc.gov/authorities/names/label/'
ACTL.loc[ACTL['Vocabulary'] == 'LCGFT', 'query'] = 'https://id.loc.gov/authorities/genreForms/label/'
ACTL['query'] = ACTL['query'] + ACTL['BIB Heading']
# make calls
ACTL['status'] = ACTL['query'].apply(lambda url: requests.get(url).status_code)
ACTL['page'] = ACTL['query'].apply(lambda url: requests.get(url).content)
ACTL['page'] = ACTL['page'].str.decode("utf-8")
# find returned value
ACTL['LCreturn'] = ACTL['page'].str.findall(r'(?<=<title>).+?(?=\s-\sLC)').apply(','.join)

# join analytics DFs
df1 = AnaP.merge(AnaE, how="outer", on=["MMS Id", "OCLC Control Number (035a)", "Local Param 01", "Creation Date"])
# copy location name to electronic collection public name
df1['Electronic Collection Public Name'] = np.where(df1['Electronic Collection Public Name'].isnull(),
                                                    df1['Location Name'], df1['Electronic Collection Public Name'])
df1.drop(['Location Name'], axis=1, inplace=True)
df1 = df1.rename(columns={'Electronic Collection Public Name': 'Collection/Location'})
df1 = df1.rename(columns={'OCLC Control Number (035a)': 'OCLC'})
df1 = df1.rename(columns={'Local Param 01': '978'})
df1 = df1.rename(columns={'Creation Date': 'BibCreation'})
# join analytics with ACTL
df2 = df1.merge(ACTL, how="inner", on="MMS Id")
df2.drop_duplicates(keep='first', inplace=True)
df2 = df2[['Report Type', 'Change Date', 'MMS Id', 'Title', 'Vocabulary', 'Field', 'BIB Heading', 'LCreturn',
           'Collection/Location', 'OCLC', '978', 'BibCreation', 'query', 'status', 'page']]
df2['MMS Id'] = df2['MMS Id'].apply(str)
df2['Change Date'] = pd.to_datetime(df2['Change Date']).dt.date
df2['OCLC'] = df2['OCLC'].str.replace('.0', '', regex=False)
df2['OCLC'] = df2['OCLC'].replace('nan', '', regex=True)
df2.sort_values(['Vocabulary', 'BIB Heading'], ascending=[True, True], inplace=True)
df2['978'] = df2['978'].replace('nan', '', regex=True)
df2 = df2.drop(['query', 'page'], axis=1)
# export a preliminary sheet for comparison
df2.to_excel(path1 + "/" + 'preprocess' + ".xlsx", index=None)
# drop resources from specified collections born in 2021 which lack field 978
coll2 = df2[(df2['Collection/Location'].isin(['ASP Streaming Video Local', 'Cambridge Core Local',
                                              'Child Coll', 'Children\'s Fiction',
                                              'Criterion Pictures Streaming Video Local',
                                              'De Gruyter eBooks Local',
                                              'EBSCO eBooks Local', 'Films on Demand Infobase Streaming Video Local',
                                              'Gen Coll', 'JSTOR eBooks Local',
                                              'Kanopy Streaming Video Local', 'Project MUSE eBooks Local',
                                              'ProQuest Ebook Central Local', 'Swank Streaming Video Local',
                                              'Taylor & Francis eBooks Local',
                                              'University Press Scholarship Online Local',
                                              'Young Adult Fiction'])) & (df2['978'] == '') &
            (df2['BibCreation'] > '2021-01-01 00:00:00')].index
df2.drop(coll2, inplace=True)
# drop local headings
coll25 = df2[(df2['BIB Heading'].isin(['Anna P. Shoemaker Collection',
                                       'Avi Wortis Historical Children\'s Book Collection',
                                       'Frances Cook Poetry Collection', 'Herman M. Ward Poetry Collection',
                                       'Faculty Author Collection', 'New Jersey Collection', 'Oral History Collection',
                                       'TCNJ Art Gallery Exhibition Catalogs', 'TCNJ Literary Publications',
                                       'TCNJ theses', 'Trenton State College theses', 'Alex Award',
                                       'Asian/Pacific American Award for Children’s Literature',
                                       'Asian/Pacific American Award for Picture Book',
                                       'Asian/Pacific American Award for Youth Literature',
                                       'Asian/Pacific American Children’s Literature honor',
                                       'Asian/Pacific American Picture Book honor',
                                       'Asian/Pacific American Youth Literature honor',
                                       'Batchelder Award', 'Batchelder Honor Book',
                                       'Belpré Author Honor Book', 'Belpré Illustrator Honor Book',
                                       'Belpré Young Adult Author Honor Book', 'Caldecott Honor Book',
                                       'Caldecott Medal', 'Coretta Scott King Author Book Award',
                                       'Coretta Scott King Author Honor Book',
                                       'Coretta Scott King Illustrator Book Award',
                                       'Coretta Scott King Illustrator Honor Book',
                                       'Coretta Scott King/John Steptoe New Talent Author Award',
                                       'Geisel Honor Book', 'Newbery Honor Book', 'Newbery Medal', 'Printz Award',
                                       'Printz Honor Book', 'Pura Belpré Author Award', 'Pura Belpré Illustrator Award',
                                       'Pura Belpré Young Adult Author Award',
                                       'Schneider Family Book Award for middle grades',
                                       'Schneider Family Book Award for teens',
                                       'Schneider Family Book Award for young children',
                                       'Schneider Family honor book for middle grades',
                                       'Schneider Family honor book for teens',
                                       'Schneider Family honor book for young children', 'Sibert Honor Book',
                                       'Sibert Informational Book Award', 'Stonewall Book Award',
                                       'Stonewall Honor Book', 'Sydney Taylor Book Award Silver Middle Grades',
                                       'Sydney Taylor Book Silver Award for Picture Book',
                                       'Sydney Taylor Book Silver Award for Young Adult',
                                       'The Sydney Taylor Book Gold Award for Middle Grades',
                                       'The Sydney Taylor Book Gold Award for Picture Books',
                                       'The Sydney Taylor Book Gold Award for Young Adult',
                                       'Theodor Seuss Geisel Award', 'William C. Morris Award',
                                       'William C. Morris Finalist',
                                       'YALSA Award for Excellence in Nonfiction for Young Adults',
                                       'YALSA Nonfiction Finalists', 'Noncitizen criminals',
                                       'Noncitizen detention centers', 'Noncitizen-owned property',
                                       'Noncitizens', 'Noncitizens (Greek law)', 'Noncitizens in art',
                                       'Noncitizens in literature', 'Noncitizens in mass media',
                                       'Noncitizens in motion pictures', 'Noncitizens (Islamic law)',
                                       'Noncitizens (Jewish law)', 'Noncitizens (Roman law)',
                                       'Children of undocumented immigrants', 'Church work with noncitizens',
                                       'Undocumented immigrant children', 'Undocumented immigrants',
                                       'Undocumented immigrants in literature',
                                       'Women undocumented immigrants']))].index
df2.drop(coll25, inplace=True)
# drop naxos lacking OCLC and all ICPSR
coll3 = df2[(df2['Collection/Location'] == 'Naxos Local') & (df2['OCLC'] == '')].index
df2.drop(coll3, inplace=True)
df2 = df2[df2['Collection/Location'].str.contains("ICPSR") == 0]
# normalize BIB Heading and LC returned values for comparison
df2['BIB Heading2'] = df2['BIB Heading'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
df2['LCreturn2'] = df2['LCreturn'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
df2['MATCH'] = (df2['BIB Heading2'] == df2['LCreturn2'])
df2 = df2[df2['MATCH'] == 0]
# remove x00 which returned any heading from LC - should be corrected via PTC
coll4 = df2[(df2['Field'].isin([100, 600, 700])) & (df2['status'] == 200)].index
df2.drop(coll4, inplace=True)
# remove columns
df2 = df2.drop(['status', 'BIB Heading2', 'LCreturn2', 'MATCH'], axis=1)
# change date format
df2['BibCreation'] = pd.to_datetime(df2['BibCreation']).dt.date

# export final
df2.to_excel(path1 + "/" + d1 + ".xlsx", index=None)
