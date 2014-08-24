'''
To do - operatinalize paths for the database, temp storage, and archive locations for users to customzie (required for operation)
'''

import shutil
import dateutil.parser
import sqlite3
import win32com.client as cl
import os
import ScrapeVariables
import re
import datetime as dt
reload(ScrapeVariables)

def validatePath(path):
    paths=ScrapeVariables.validPaths
    #print paths
    items = path.split('\\')
    for i in xrange(len(items)):
        if i>0:
            #print '\\'.join(items[0:i])
            if '\\'.join(items[0:i+1]) in paths:
                return(True)
    return(False)

def getFileType(filename):
    '''code exames the last 5 characters of a filename and splits it on "." to find the file typep
    unknown types return the phrase "uknown [png]" or whatever is after the dot '''
    typeDict={'xlsx':'excel','xls':'excel','doc':'word','docx':'word','ppt':'ppt','pptx':'ppt','dotx':'word'}
    try:
        stem=filename[-5:].split('.')[1]
    except:
        print 'File type misidentification = ', filename
        return('unknown '+filename)
    try:
        doctype=typeDict[stem]
        return(doctype)
    except:
        return('unknown '+stem)
    
def blacklistcheck(text):
    '''import the blacklist from validatePaths which is a list of words or phrases the presence of which indicates
    scraping should not occur.  Takes the text, whether a path, filename, or even open text and looks for blacklisted
    words or phrases.  print(validatePaths.blacklist) for the list of phrases.'''
    blackflags=ScrapeVariables.blackList
    #check filename for blackflags
    if os.path.isdir(text):
        words = [re.findall(r'[a-z]+',text.lower()) for t in text.split('\\')[1:]]
    else:
        words=re.findall(r'[a-z]+',text.lower())
    for flag in blackflags:
        count=0
        flags=flag.split()
        for fl in flags:
            if fl in words:
                count+=1
        if count==len(flags):
            return(True)
    return(False)

def FileStatus(path,filename):
    con=sqlite3.connect('C:\\Users\\jradford\\Documents\\TestDocs\\Test.db')
    curs=con.cursor()
    res=curs.execute("""SELECT lastSaveTime FROM rawScrapeData WHERE path = %s AND filename = %s """ % ("?","?"),[path,filename]).fetchall()
    if len(res)>0:
        status='old'
        last=dateutil.parser.parse('1985-01-30 16:43:54')
        recent = dt.datetime.fromtimestamp(os.stat(path+'\\'+filename).st_mtime)
        for r in res:
            saved_time=dateutil.parser.parse(r).replace(tzinfo=None)
            if saved_time>last:
                saved_time==last
        if recent>saved_time:
            t=recent-saved_time
            if t.seconds>60*5:   #basically, if the last change was 5 minutes after the most recently stored version, it's an update.
                status='edit'
        curs.close()
        con.close()
        return(status)
        
    else:
        status='new'
    curs.close()
    con.close()
    return(status)
    
def logBlackData(path,filename,filetype):
    '''This is the same as ProcessFile, except that copies are not saved in the archive.'''
    if 'unknown' in filetype:
        recordUnknownType(path,filename,filetype,flag='black list')
    else:
        status=FileStatus(path,filename)
        if status is not 'old':
            data=getKnownData(path,filename,filetype)
            data['type']=status
            data['flag']='black list'
            SaveData(data)
    return()

def getKnownData(path,fname,doctype):
    '''This code takes a file with a type we can scrape and scrapes the data accordingly.  In the future,
    this will be split into a separate file with more options
    INPUTS: path - original file path as string - "C:\\Documents"; fname - filename with suffix - "File.docx"
    doctype - string for document - "word","excel","ppt" 
    OUTPUTS: dictionary of Native Variable Name: Native Values - i.e. "Last Author": "John Smith"'''
    props={'path':path,'fileName':fname}
    #This work-around enables me to access files without changing the date-stamp on their folders
    temp='C:\\Temp'
    nfile=fname
    shutil.copy2(path+'\\'+fname, temp+'\\'+nfile)  #this is key; it copies the file to a temporary location and 
    filename=temp+'\\'+nfile                        #and scrapes it there. This prevents changing directory information
                                                    #in the original directoy location
  
    docProperties=['Title', 'Subject', 'Author', 'Keywords', 'Comments', 'Template', 'Last Author', 'Revision Number', 'Application Name',
                   'Last Print Date', 'Creation Date', 'Last Save Time', 'Total Editing Time', 'Number of Pages', 'Number of Words',
                   'Number of Characters', 'Security', 'Category', 'Format', 'Manager', 'Company', 'Number of Bytes', 'Number of Lines',
                   'Number of Paragraphs', 'Number of Slides', 'Number of Notes', 'Number of Hidden Slides', 'Number of Multimedia Clips',
                   'Hyperlink Base', 'Number of Characters (with spaces)']
    if doctype=='word':
        w= cl.Dispatch("Word.Application")
        w.Documents.Open(filename)
        for d in docProperties:
            try:
                props[d]=str(w.ActiveDocument.BuiltInDocumentProperties(d))
            except:
                pass
        w.Documents.Close(False)
        w.Application.Quit()
    elif doctype=='excel':
        w= cl.Dispatch("Excel.Application")
        w.Workbooks.Open(filename)
        for d in docProperties:
            try:
                props[d]=str(w.ActiveWorkbook.BuiltInDocumentProperties(d))
            except:
                pass
        w.Workbooks.Close()
        w.Application.Quit()
    elif doctype=='ppt':
        w= cl.Dispatch("PowerPoint.Application")
        Presentation=w.Presentations.Open(filename)
        for d in docProperties:
            try:
                props[d]=str(w.ActivePresentation.BuiltInDocumentProperties(d))
            except:
                pass
        Presentation.Close()
        w.Quit()
    #if olddir!='':
    #    os.chdir(olddir)
    os.remove(temp+'\\'+nfile)
    #word=win32com.client.Dispatch("Word.Document")
    return props

def recordUnknownType(path,filename,typ,flag=''):
    '''This code takes a filetype I can't work with and simply records its system meta-data. It does not archive the file'''
    stats=os.stat(path+'\\'+filename)
    data={
        'type':typ.split(' ')[0],
        'path':path,
        'fileName':filename,
        'flag':typ.split(' ')[1],
        'createdDate':dt.datetime.isoformat(dt.datetime.fromtimestamp(stats.st_ctime),' '),
        'lastSaveTime':dt.datetime.isoformat(dt.datetime.fromtimestamp(stats.st_mtime),' '),
        #'accessDate':dt.datetime.fromtimestamp(stats.st_atime), atime is useless, almost always = mtime; see wikipedia.org/wiki/stat for why
        'fileSize':stats.st_size,
        'scrapeTime':dt.datetime.isoformat(dt.datetime.now(),' ')
        }
    SaveData(data)
    return()
    
def unifyData(data):
    '''data is a dictionary of fieldName:value for any data object.  This code takes that dictionary, fills any missing data,
    and writes the data to file'''
    Data=ScrapeVariables.dataFormat
    mapper=ScrapeVariables.mapper
    for k,v in data.iteritems():
        try:v=(v.decode('windows-1252'))
        except: pass
        try:
            Data[mapper[k]]=v
            #print 'mapped ', k
        except:
            #print 'cannot map', k
            pass
    for k,v in Data.iteritems():
        if v=='None':
            Data[k]=None
    for k,v in Data.iteritems():
        if k in ['fileSize','revisions','Number of Bytes']:
            if v!=None:
                Data[k]=int(v)
        elif k in ['totalEditingTime']:
            if v!=None:
                Data[k]=float(v)
        elif k in ['createdDate','lastSaveTime','scrapeTime']:
            if v!=None:
                Data[k]=str(v)


    if 'scrapeTime' not in Data.keys() or Data['scrapeTime']==None:
        Data['scrapeTime']=dt.datetime.isoformat(dt.datetime.now(),' ')
    return(Data)

def ArchiveFile(path,filename):
    '''This code generates a new filename and directory for archiving each file.  The original data source
    was a network with large depth.  Hence, most of the code is meant to handle path locations that are 
    traditionally considered too long for windows.'''
    newpath=path.replace('C:\\','C:\\Archive\\')
    ftype='.'+filename.split('.')[len(filename.split('.'))-1]
    name='.'.join(filename.split('.')[0:len(filename.split('.'))-1])
    if len(newpath+'\\'+filename)>240:
        if len(newpath)>255:
            print 'ERROR! path too long! ', newpath
            return()
        elif len(newpath+'\\'+filename)>255:
            if len(newpath)>245:
                print 'ERROR! path too long! ', newpath
                return()
            else:
                l=255-len(newpath+'\\')     #the >255 and path>245 conditions mean file names here are > 10 chars.
                for i in range(7,l):        
                    shortfilename=name[0:i]+ftype
                    if len(newpath+'\\'+shortfilename)>=255:
                        print 'ERROR! Cannot find shortest name', path, filename
                        return()
                    
                    if shortfilename not in os.listdir(path):
                        try:
                            if shortfilename in os.listdir(newpath): continue
                            else: break
                        except: break
        else:
            shortfilename=name+ftype

        
    else:
        shortfilename=name+ftype
        
    chunks=newpath.split('\\')
    chunk='\\'.join(chunks[0:2])
    try:
        shutil.copy2(path+'\\'+filename, newpath+'\\'+shortfilename)
    except:
        for c in chunks[2:]:
            chunk=chunk+'\\'+c
            try: os.mkdir(chunk)
            except: pass
        shutil.copy2(path+'\\'+filename, newpath+'\\'+shortfilename)
    return(newpath,shortfilename)
    
    
def SaveData(data):
    data=unifyData(data)
    con=sqlite3.connect('C:\\Users\\jradford\\Documents\\TestDocs\\Test.db')
    curs=con.cursor()
    qmarks = ','.join('?' * len(data))
    qry="Insert Into rawScrapeData (%s) Values (%s)" % (','.join(data.keys()), qmarks)
    curs.execute(qry,data.values())
    con.commit()
    curs.close()
    con.close()
        
    return

def ProcessFile(path,fname):
    filetype=getFileType(fname)
    if blacklistcheck(path+'\\'+fname)==True:
        print 'blacklisted - ', path+'\\'+fname
        flag='black list'
        logBlackData(path,fname,filetype)
        return()
    
    if 'unknown' in filetype:
        data=recordUnknownType(path,fname,filetype)
        return()

    status=FileStatus(path,fname)
    if status is not 'old':
        data=getKnownData(path,fname,filetype)
        data['type']=status
        #data[archivedLocation]=archive(path,fname)
        spath,shortfile=ArchiveFile(path,fname)
        data['archivedLocation']=spath+'\\'+shortfile
        SaveData(data)
        return()


def customWalker(folder,l):
    directContents = os.listdir(folder)
    for item in directContents:
        if os.path.isfile(os.path.join(folder, item)):
            l.append(os.path.join(folder, item))
        else:customWalker(os.path.join(folder, item), l)
    return l


def ScrapeFiles(l):
    print 'number of files is ', len(l)
    for paths in l:
        dirs=paths.split('\\')
        path='\\'.join(dirs[0:len(dirs)-1])
        f=dirs[len(dirs)-1]
        if validatePath(path) is True:
            if f.startswith('~'):
                continue
            if f.lower() in ScrapeVariables.excludeFiles:
                continue
            print path+'  '+f
            ProcessFile(path,f)
        else:
            print 'invalid path, ', path
    return
    

def InitializeDB():
    '''This code initiates the database.  It should only be run the first time the DB is created '''
    print 'initailizing DB'
    con=sqlite3.connect('C:\\Users\\jradford\\Documents\\TestDocs\\Test.db')
    curs=con.cursor()
    curs.execute("""DROP TABLE IF EXISTS rawScrapeData""")
    con.commit()
    curs.execute("""CREATE TABLE IF NOT EXISTS rawScrapeData(
                type	TEXT	,
                fileName	TEXT	,
                path	TEXT	,
                author	TEXT	,
                createdDate	TEXT	,
                lastAuthor	TEXT	,
                lastSaveTime	TEXT	,
                scrapeTime	TEXT	,
                fileSize	INTEGER	,
                templateFile	TEXT	,
                revisions	INTEGER	,
                totalEditingTime	REAL	,
                archivedLocation	TEXT    ,
                flag    TEXT
                )""")
    con.commit()
    curs.close()
    con.close()
    return

 
InitializeDB()
for vPath in ScrapeVariables.validPaths:
    files=customWalker(vPath,[])
    ScrapeFiles(files)
    
