

validPaths=[
    'C:\\Users\\USER_NAME_HERE\\Documents', #note, this a default path.
    'NETWORK_ROOT:\\'NETWORK_DIRECTORIES',
]

blackList=['']  #list of strings for filenames you want the metadata for, but 
                #not those to archive (i.e. "tax records" "sensitive.doc")

invalidPaths=[  #list of known folders with documents I do not want
    'C:\\Users\\Administrator\\Documents\\Exclude',
    ]
    
excludeFiles=['thumbs.db']  #files that you don't want to scrape or archive (i.e. useless files, inaccessible files)

dataFormat={                    #key for all variables that can go into the database for storage
    'type': None,
    'fileName': None,
    'path': None,
    'author': None,
    'createdDate': None,
    'lastAuthor': None,
    'lastSaveTime':None ,
    'scrapeTime':None ,
    'fileSize':None,
    'templateFile':None,
    'revisions':None,
    'totalEditingTime':None,
    'archivedLocation':None,
    'flag':None}

networkDataFormat={             #Key for variables used in generating a co-editing network
    'title':None,
    'author': None,
    'createdDate': None,
    'lastAuthor': None,
    'lastSaveTime':None ,
    }

mapper={                        #Translater for Microsoft parameter names and dataFormat names.
    'Author':'author',
    'Creation Date':'createdDate',
    'Number of Bytes':'fileSize',
    'Last Author':'lastAuthor',
    'Last Save Time':'lastSaveTime',
    'Revision Number':'revisions',
    'Template':'templateFile',
    'Total Editing Time':'totalEditingTime',
    'type':'type',
    'fileName':'fileName',
    'path':'path',
    'author':'author',
    'createdDate':'createdDate',
    'lastAuthor':'lastAuthor',
    'lastSaveTime':'lastSaveTime',
    'scrapeTime':'scrapeTime',
    'fileSize':'fileSize',
    'templateFile':'templateFile',
    'template':'templateFile',
    'revisions':'revisions',
    'totalEditingTime':'totalEditingTime',
    'archivedLocation':'archivedLocation',
    'flag':'flag'
}



