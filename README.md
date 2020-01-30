![MIKES DATA WORK GIT REPO](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_01.png "Mikes Data Work")        

# Use SQL To Extract Sharepoint Documents
**Post Date: May 10, 2018**        



## Contents    
- [About Process](##About-Process)  
- [SQL Logic](#SQL-Logic)  
- [Build Info](#Build-Info)  
- [Author](#Author)  
- [License](#License)       

## About-Process

![Extract Sharepoint Documents With SQL]( https://mikesdatawork.files.wordpress.com/2018/05/image0011.png "Extract Sharepoint Documents With SQL")
 
<p>Ok so here's some SQL logic I created a while back that will build a list of all your docs (doc, docx, pdf, txt, msg) across all your Sharepoint Content Databases. It could probably use a bit of updating at this stage. This is using straight SQL with no procedures or functions although I do understand the benefits there; I just didn't write this with those in mind. If you have modifications, or improvements; please by all means –drop in a comment. It's far from perfect, and definitely not elegant, but it does work. This process is excellent for manual document extraction and archiving, compliance sleuthing, capacity management, etc. This process is not entirely automatic. It builds a set of extraction scripts (in xml form) so all you need to do is click on the xml link and you'll get all scripts. It's up to you if you want to run them all at once, store them in a table, or get only certain extraction scripts for what you're looking for. All you need to do is modify the WHERE clause at the bottom. I've also included a couple @wildcard variables for more granular results on your queries.
How does it work?
  
It builds a list using the Database Name, List Name, Site URL, Document Name, and ID. It filters for doc, docx, pdf, txt, and msg but you can easily add or remove different document types if needed.
It takes this information and pulls it all into a Temp table called ##sample_collection. All processes thereafter are dependent on the temp table so your Sharepoint database are largely untouched until extraction begins and even then; it's just a read operation to get the document binaries, and then the document is ultimately downloaded. If you have millions of 2mb files (like I do), it's surprisingly quick when the downloads begin. If you have a large Sharepoint environment; make sure your Temp space is up to scratch before running this process. Most of the burden is on the Temp and it takes alittle while to get all the scripts produced so just be patient for that.
By the way; the folder structure in this example is going to this path:
\\MyServerName\W$\BACKUPS\Sharepoint_Extraction
All the other folders are added to it.

You can always stop the process with no risk if needed. Again; this does NOT automatically run. It presents the extraction scripts to you.
It creates an identical folder structure to that of the Site URL where the documents exist. The folder structure is readily identified within the logic. Each folder is prefixed by it's purpose according to Database Server, Instance Name, and Sharepoint tables.
For example; it always begins with the instance name, and end with the document name.
SRV_MyServerName_MyInstanceName\DB_MyDatabaseName\LIST_MyListName\URL_MyURLInFolderFormat\MyDocument.pdf
If you have long paths in the URL; expect an equally deep folder structure for the documents. The folder structure removes commas and hyphens, but you'll need to add more REPLACE statements to accommodate any other special characters.

I also used a 6 character suffix on all variables which is created by using the (first 3 + last 3) characters of the [id] to prevent any variable conflict. You'll see a lot of that repeated. This could have been handled better with a function or procedure but I don't like creating any other additional objects if at all possible especially for administrative, or maintenance processes. Keeping it straight SQL keeps it extremely portable when applying it to other environments.

To keep things lightweight; I went ahead and added a (select top 10) on the script so you can see what the output looks like before you run the full set. Like I said it does take some time produce scripts. Just remove the 'top 10' and you should be all set.
Anyway; so that's the run-down on this. Hope you find it helpful.
Here's the logic:</p>      

## SQL-Logic
```SQL
use master;
set nocount on
 
-- check for advanced options and ole automation
declare @adv_options    int
declare @ole_automation int
set     @adv_options    = (select cast([value_in_use] as int) from sys.configurations where [configuration_id] = '518')
set     @ole_automation = (select cast([value_in_use] as int) from sys.configurations where [configuration_id] = '16388') 
if      @adv_options    = 0 begin exec master..sp_configure 'show advanced options', 1;     reconfigure with override; end
if      @ole_automation = 0 begin exec master..sp_configure 'Ole Automation Procedures', 1; reconfigure with override; end
 
-- build logic to get file info (doc, docx, pdf, txt) across all content databases
declare @wildcard1      varchar(255)
declare @wildcard2      varchar(255)
declare @get_file_info  varchar(max)
set     @wildcard1      = '%design%'
set     @wildcard2      = '%design%'
set     @get_file_info  = ''
select  @get_file_info  = @get_file_info +
'
use [' + [name] + '];' + char(10) + 
'
 select 
    ''database''            = db_name()
,   ''time_created''        = alldocs.timecreated
,   ''list_name''           = alllists.tp_title
,   ''file_name''           = alldocs.leafname
,   ''url''                 = alldocs.dirname
,   ''id''                  = cast(alldocs.id as varchar(255))
from 
    alldocs join alldocstreams  on alldocs.id=alldocstreams.id 
    join alllists               on alllists.tp_id = alldocs.listid
where
    alldocs.leafname like       ''' + @wildcard1 + '''
    or  alldocs.dirname like    ''' + @wildcard2 + '''
    and right(alldocs.leafname, 2) in (''oc'', ''cx'', ''df'', ''xt'', ''sg'')
'
from    sys.databases where [name] like '%content%'
 
-- build temp table for results (this will take several minutes)
if object_id('tempdb..##sample_collection') is null
    begin
        create table    ##sample_collection ([database] varchar(255), [time_created] datetime, [list_name] varchar(255), [file_name] varchar(510), [url] varchar(510), [id] varchar(255))
        insert into     ##sample_collection exec    (@get_file_info)
    end;
 
-- get count of documents in table
select count(*) from ##sample_collection
 
-- build extraction process including folder structure. this will take several minutes.
declare @file_extraction    varchar(max)
set     @file_extraction    = ''
select  top 10 @file_extraction = @file_extraction +
'
use [' + [database] + '];
set nocount on;' + char(10) + '
 
declare     @folder'    + '_' + left(id, 3) + right(id, 3) + ' nvarchar(4000) 
set         @folder'    + '_' + left(id, 3) + right(id, 3) + ' = N''\\MyServerName\W$\BACKUPS\Sharepoint_Extraction\' 
            + 'SVR_'    + replace(upper(@@servername), '\', '_') + '\' 
            + 'DB_'     + upper([database])     + '\' 
            + 'LIST_'   + upper([list_name])    + '\' 
            + 'URL_'    + replace(replace(replace(upper([url]), '/', '\'), '''', ''), ',', '') + '''
exec        master..xp_create_subdir @folder' + '_' + left(id, 3) + right(id, 3) + '; 
 
declare @object_token'      + '_' + left(id, 3) + right(id, 3) + ' int
declare @destination_path'  + '_' + left(id, 3) + right(id, 3) + ' varchar(255)
declare @content_binary'    + '_' + left(id, 3) + right(id, 3) + ' varbinary(max)
set     @destination_path'  + '_' + left(id, 3) + right(id, 3) + ' = (select @folder' + '_' + left(id, 3) + right(id, 3) + ' + ''\' +  [file_name] + ''')
select  @content_binary'    + '_' + left(id, 3) + right(id, 3) + ' = alldocstreams.content from alldocs join alldocstreams on alldocs.id = alldocstreams.id join alllists on alllists.tp_id = alldocs.listid
where  
    alldocs.leafname    = ''' + [file_name] + '''
    and alldocs.dirname = ''' + [url]       + '''
 
exec sp_oacreate ''adodb.stream'', @object_token' + '_' + left(id, 3) + right(id, 3) + ' output
exec sp_oasetproperty   @object_token' + '_' + left(id, 3) + right(id, 3) + ', ''type'', 1
exec sp_oamethod        @object_token' + '_' + left(id, 3) + right(id, 3) + ', ''open''
exec sp_oamethod        @object_token' + '_' + left(id, 3) + right(id, 3) + ', ''write'',       null, @content_binary'      + '_' + left(id, 3) + right(id, 3) + '
exec sp_oamethod        @object_token' + '_' + left(id, 3) + right(id, 3) + ', ''savetofile'',  null, @destination_path'    + '_' + left(id, 3) + right(id, 3) + ', 2
exec sp_oamethod        @object_token' + '_' + left(id, 3) + right(id, 3) + ', ''close''
exec sp_oadestroy       @object_token' + '_' + left(id, 3) + right(id, 3) + '
'
from ##sample_collection where right([file_name], 2) in ('oc', 'cx', 'df', 'xt', 'sg')
 
select  (@file_extraction) for xml path(''), type
 
-- drop table ##sample_collection
```
One point I would like to make is I noticed that some documents (both pdf and doc) will periodically be (or appear to be) empty files, but the size of the file usually shows bytes so there should be some data in there. For whatever reason; I'm not able to see the content post download. This has been random, and it could be that teh files are checked out, but I cannot say for certain. It's been maybe 1 of every 200 documents or so. If anyone has tips on what is causing this; please post a comment. 



[![WorksEveryTime](https://forthebadge.com/images/badges/60-percent-of-the-time-works-every-time.svg)](https://shitday.de/)

## Build-Info

| Build Quality | Build History |
|--|--|
|<table><tr><td>[![Build-Status](https://ci.appveyor.com/api/projects/status/pjxh5g91jpbh7t84?svg?style=flat-square)](#)</td></tr><tr><td>[![Coverage](https://coveralls.io/repos/github/tygerbytes/ResourceFitness/badge.svg?style=flat-square)](#)</td></tr><tr><td>[![Nuget](https://img.shields.io/nuget/v/TW.Resfit.Core.svg?style=flat-square)](#)</td></tr></table>|<table><tr><td>[![Build history](https://buildstats.info/appveyor/chart/tygerbytes/resourcefitness)](#)</td></tr></table>|

## Author

[![Gist](https://img.shields.io/badge/Gist-MikesDataWork-<COLOR>.svg)](https://gist.github.com/mikesdatawork)
[![Twitter](https://img.shields.io/badge/Twitter-MikesDataWork-<COLOR>.svg)](https://twitter.com/mikesdatawork)
[![Wordpress](https://img.shields.io/badge/Wordpress-MikesDataWork-<COLOR>.svg)](https://mikesdatawork.wordpress.com/)

  
## License
[![LicenseCCSA](https://img.shields.io/badge/License-CreativeCommonsSA-<COLOR>.svg)](https://creativecommons.org/share-your-work/licensing-types-examples/)

![Mikes Data Work](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_02.png "Mikes Data Work")

