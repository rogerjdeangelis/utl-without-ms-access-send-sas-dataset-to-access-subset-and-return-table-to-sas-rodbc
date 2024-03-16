# utl-without-ms-access-send-sas-dataset-to-access-subset-and-return-table-to-sas-rodbc
Without ms access send sas dataset to access modify and return table to sas rodbc
    %let pgm=utl-without-ms-access-send-sas-dataset-to-access-subset-and-return-table-to-sas-rodbc;

    Without ms access send sas dataset to access modify and return table to sas rodbc

    github
    https://tinyurl.com/36zvd8w8
    https://github.com/rogerjdeangelis/utl-without-ms-access-send-sas-dataset-to-access-subset-and-return-table-to-sas-rodbc

    Problem
       Without Microsoft Access Installed select males from an access 'class' table

    At the end of this message
    see list of availble SQL drivers and r ability to read and write all these file types.

    /*
     _ __  _ __ ___ _ __
    | `_ \| `__/ _ \ `_ \
    | |_) | | |  __/ |_) |
    | .__/|_|  \___| .__/
    |_|            |_|
    */

    What you have to do if you don't have ms access installed or ms access drivers

    1. Get odbc driver

       If you don't have the ms access database driver software go to

       https://www.microsoft.com/en-us/download/details.aspx?id=54920

    2. Download powershell macro to create ODBC connections

       If you don't have the odbc connection you can use my powershell macro to crrate one.

       Powershell macro to setup odbc connections. You can do the setup manually.control panel/dydtem & security/odbc.

       https://tinyurl.com/mrenx557
       https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc

    3. Create 64 and 32 bit ms access drivers using powershell;

       %utl_submit_ps64('
       Add-OdbcDsn -Name "have" -DriverName "Microsoft Access Driver (*.mdb, *.accdb)" -DsnType "User" -Platform "64-bit" -SetPropertyValue "Dbq=d:/mdb/simle.mdb
       Get-OdbcDsn;
       ');
       %utl_submit_ps64('
       Add-OdbcDsn -Name "have" -DriverName "Microsoft Access Driver (*.mdb, *.accdb)" -DsnType "User" -Platform "32-bit" -SetPropertyValue "Dbq=d:/mdb/simle.mdb
       Get-OdbcDsn;
       ');

    4. Download template ms access mdb database

       Because MS Access has changed the layout of access databases many times you need to use
       the simple.mdb that SAS provides or get one off the net. Should also work with accdb.

       https://tinyurl.com/yaxa7nty
       https://github.com/rogerjdeangelis/utl-without-ms-access-send-sas-dataset-to-access-subset-and-return-table-to-sas-rodbc/raw/main/simple.md

       * download simple.mdb from github;
       %utlfkil(d:/mdb/simple.mdb);
       filename out "d:/mdb/simple.mdb";
       proc http
         method = 'GET'
         url    = "https://tinyurl.com/yaxa7nty"
         out    =  out;
       run;quit;

    FYI
      Might be useful to have powershell create all the free odbc driver DSNs
      Here is excel and access

    %utl_submit_ps64('
    Add-OdbcDsn -Name "have" -DriverName "Microsoft Access Driver (*.mdb, *.accdb)" -DsnType "User" -Platform "64-bit" -SetPropertyValue "Dbq=d:/mdb/simle.mdb
    Get-OdbcDsn;
    ');

    %utl_submit_ps64("
    Add-OdbcDsn -Name 'have' -DriverName 'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)' -DsnType 'User' -Platform '64-bit' -SetPropertyValue 'Dbq=d:\xls\have.xlsx')
    Get-OdbcDsn;
    ");

    /*                                            __ _
     _ __  _ __ ___   __ _ _ __ __ _ _ __ ___    / _| | _____      __
    | `_ \| `__/ _ \ / _` | `__/ _` | `_ ` _ \  | |_| |/ _ \ \ /\ / /
    | |_) | | | (_) | (_| | | | (_| | | | | | | |  _| | (_) \ V  V /
    | .__/|_|  \___/ \__, |_|  \__,_|_| |_| |_| |_| |_|\___/ \_/\_/
    |_|              |___/
    */

     1. input sas dataset
        have<-read_sas("d:/sd1/have.sas7bdat");

     2. Create connection
        myDB<-odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=d:/mdb/simple.mdb");

     3. create ms access table have
        sqlSave(myDB,have,rownames=FALSE);

     4. Subset table for males
        want<-sqlQuery(myDB, paste("select * from have where SEX='M' "));

     5.  Create sas dataset
         fn_tosas9(dataf=want);

    /*               _     _
     _ __  _ __ ___ | |__ | | ___ _ __ ___
    | `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
    | |_) | | | (_) | |_) | |  __/ | | | | |
    | .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
    |_|
    */

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*            INPUT                                                                                                       */
    /*                                                                                                                        */
    /*  options validvarname=upcase;                                                                                          */
    /*  libname sd1 "d:/sd1";                                                                                                 */
    /*  data sd1.have;                                                                                                        */
    /*   set sashelp.class(keep=name sex age);                                                                                */
    /*  run;quit;                                                                                                             */
    /*                                                                                                                        */
    /*  SD1.HAVE total obs=19                                                                                                 */
    /*                                                                                                                        */
    /*  Obs    NAME       SEX    AGE                                                                                          */
    /*                                                                                                                        */
    /*    1    Alfred      M      14                                                                                          */
    /*    2    Alice       F      13                                                                                          */
    /*    3    Barbara     F      13                                                                                          */
    /*    4    Carol       F      14                                                                                          */
    /*    5    Henry       M      14                                                                                          */
    /*    6    James       M      12                                                                                          */
    /*    7    Jane        F      12                                                                                          */
    /*    8    Janet       F      15                                                                                          */
    /*    9    Jeffrey     M      13                                                                                          */
    /*   10    John        M      12                                                                                          */
    /*   11    Joyce       F      11                                                                                          */
    /*   12    Judy        F      14                                                                                          */
    /*   13    Louise      F      12                                                                                          */
    /*   14    Mary        F      15                                                                                          */
    /*   15    Philip      M      16                                                                                          */
    /*   16    Robert      M      12                                                                                          */
    /*   17    Ronald      M      15                                                                                          */
    /*   18    Thomas      M      11                                                                                          */
    /*   19    William     M      15                                                                                          */
    /*                                                                                                                        */
    /*------------------------------------------------------------------------------------------------------------------------*/
    /*                                                                                                                        */
    /*  PROCESS                                                                                                               */
    /*                                                                                                                        */
    /*  libname tmp "c:/temp";                                                                                                */
    /*  proc datasets lib=tmp nolist nodetails;                                                                               */
    /*   delete want;                                                                                                         */
    /*  run;quit;                                                                                                             */
    /*                                                                                                                        */
    /*  /*----                                                                   ----*/                                       */
    /*  /*---- I was unable to split the myDB connection string statement        ----*/                                       */
    /*  /*---- I also unable to split the dd-OdbcDsn string in powershell        ----*/                                       */
    /*  /*----                                                                   ----*/                                       */
    /*                                                                                                                        */
    /*  %utl_rbeginx;                                                                                                         */
    /*  parmcards4;                                                                                                           */
    /*  library(RODBC);                                                                                                       */
    /*  library(haven);                                                                                                       */
    /*  source("c:/temp/fn_tosas9.R");                                                                                        */
    /*  have<-read_sas("d:/sd1/have.sas7bdat");                                                                               */
    /*  myDB<-odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=d:/mdb/simple.mdb");                   */
    /*  sqlQuery(myDB, paste("drop table have"));                                                                             */
    /*  sqlSave(myDB,have,rownames=FALSE);                                                                                    */
    /*  want<-sqlQuery(myDB, paste("select * from have where SEX='M' "));                                                     */
    /*  want;                                                                                                                 */
    /*  fn_tosas9(dataf=want);                                                                                                */
    /*  str(want);                                                                                                            */
    /*  ;;;;                                                                                                                  */
    /*  %utl_rendx;                                                                                                           */
    /*                                                                                                                        */
    /*----------------------------------------------------------------------------------------------------------------------- */
    /*                                                                                                                        */
    /*  OUTPUT                                                                                                                */
    /*                                                                                                                        */
    /*  libname tmp "c:/temp";                                                                                                */
    /*  proc print data=tmp.want;                                                                                             */
    /*  run;quit;                                                                                                             */
    /*                                                                                                                        */
    /*  Obs    ROWNAMES    NAME       SEX    AGE                                                                              */
    /*                                                                                                                        */
    /*    1        1       Alfred      M      14                                                                              */
    /*    2        2       Henry       M      14                                                                              */
    /*    3        3       James       M      12                                                                              */
    /*    4        4       Jeffrey     M      13                                                                              */
    /*    5        5       John        M      12                                                                              */
    /*    6        6       Philip      M      16                                                                              */
    /*    7        7       Robert      M      12                                                                              */
    /*    8        8       Ronald      M      15                                                                              */
    /*    9        9       Thomas      M      11                                                                              */
    /*   10       10       William     M      15                                                                              */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.have;
     set sashelp.class(keep=name sex age);
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* SD1.HAVE total obs=19                                                                                                  */
    /*                                                                                                                        */
    /* Obs    NAME       SEX    AGE                                                                                           */
    /*                                                                                                                        */
    /*   1    Alfred      M      14                                                                                           */
    /*   2    Alice       F      13                                                                                           */
    /*   3    Barbara     F      13                                                                                           */
    /*   4    Carol       F      14                                                                                           */
    /*   5    Henry       M      14                                                                                           */
    /*   6    James       M      12                                                                                           */
    /*   7    Jane        F      12                                                                                           */
    /*   8    Janet       F      15                                                                                           */
    /*   9    Jeffrey     M      13                                                                                           */
    /*  10    John        M      12                                                                                           */
    /*  11    Joyce       F      11                                                                                           */
    /*  12    Judy        F      14                                                                                           */
    /*  13    Louise      F      12                                                                                           */
    /*  14    Mary        F      15                                                                                           */
    /*  15    Philip      M      16                                                                                           */
    /*  16    Robert      M      12                                                                                           */
    /*  17    Ronald      M      15                                                                                           */
    /*  18    Thomas      M      11                                                                                           */
    /*  19    William     M      15                                                                                           */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    libname tmp "c:/temp";
    proc datasets lib=tmp nolist nodetails;
     delete want;
    run;quit;

    /*----                                                                   ----*/
    /*---- I was unable to split the myDB connection string                  ----*/
    /*---- Ialso unable to split the dd-OdbcDsn string                       ----*/
    /*----                                                                   ----*/
    /*----  download simple.mdb from github                                  ----*/
    /*----                                                                   ----*/

    %utlfkil(d:/mdb/simple.mdb);
    filename out "d:/mdb/simple.mdb";
    proc http
      method = 'GET'
      url    = "https://tinyurl.com/yaxa7nty"
      out    =  out;
    run;quit;

    %utl_rbeginx;
    parmcards4;
    library(RODBC);
    library(haven);
    source("c:/temp/fn_tosas9.R");
    have<-read_sas("d:/sd1/have.sas7bdat");
    myDB<-odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=d:/mdb/simple.mdb");
    sqlQuery(myDB, paste("drop table have"));
    sqlSave(myDB,have,rownames=FALSE);
    want<-sqlQuery(myDB, paste("select * from have where SEX='M' "));
    want;
    fn_tosas9(dataf=want);
    str(want);
    ;;;;
    %utl_rendx;

    libname tmp "c:/temp";
    proc print data=tmp.want;
    run;quit;
    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* Obs    ROWNAMES    NAME       SEX    AGE                                                                               */
    /*                                                                                                                        */
    /*   1        1       Alfred      M      14                                                                               */
    /*   2        2       Henry       M      14                                                                               */
    /*   3        3       James       M      12                                                                               */
    /*   4        4       Jeffrey     M      13                                                                               */
    /*   5        5       John        M      12                                                                               */
    /*   6        6       Philip      M      16                                                                               */
    /*   7        7       Robert      M      12                                                                               */
    /*   8        8       Ronald      M      15                                                                               */
    /*   9        9       Thomas      M      11                                                                               */
    /*  10       10       William     M      15                                                                               */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*__                           _ _                _      _
     / _|_ __ ___  ___    ___   __| | |__   ___    __| |_ __(_)_   _____ _ __ ___
    | |_| `__/ _ \/ _ \  / _ \ / _` | `_ \ / __|  / _` | `__| \ \ / / _ \ `__/ __|
    |  _| | |  __/  __/ | (_) | (_| | |_) | (__  | (_| | |  | |\ V /  __/ |  \__ \
    |_| |_|  \___|\___|  \___/ \__,_|_.__/ \___|  \__,_|_|  |_| \_/ \___|_|  |___/

    */

    Free drivers

    https://web.synametrics.com/odbcdrivervendors.htm

    Following drivers are included with WinSQL Professional for free.
    These drivers are published by Progress Software -
    DataDirect and licensed to be bundled with WinSQL Professional

    Amazon Redshift
    Apache Cassandra
    Apache Hive
    Apache Spark SQL
    DB2
    Google BigQuery
    Informix
    MongoDB
    MySQL Enterprise
    OpenEdge
    Oracle
    PostgreSQL
    Salesforce
    Snowflake
    MS SQL Server
    Sybase
    Text Files

    /*         _ _                                        _           _           _
      ___   __| | |__   ___   ___ _   _ _ __  _ __   ___ | |_ ___  __| |  _ __ __| |_ __ ___  ___
     / _ \ / _` | `_ \ / __| / __| | | | `_ \| `_ \ / _ \| __/ _ \/ _` | | `__/ _` | `_ ` _ \/ __|
    | (_) | (_| | |_) | (__  \__ \ |_| | |_) | |_) | (_) | ||  __/ (_| | | | | (_| | | | | | \__ \
     \___/ \__,_|_.__/ \___| |___/\__,_| .__/| .__/ \___/ \__\___|\__,_| |_|  \__,_|_| |_| |_|___/
                                       |_|   |_|
    */

    Supported RDBMS by third parties

    01 4D Server
    02 Vision Data
    03 Accuracer, Easytable
    04 ANTs Database Server
    05 UniAccess for OS 2200
    06 Flat files, VSAM, IMS, DB2
    07 ASTA Server
    08 Business BASIC
    09 Biblioscape Database
    10 Birdstep RDM Server
    11 Supra SQL
    12 CONNX
    13 Btrieve, Centura SQLBase, Clipper, DB2, dBASE, Excel, FoxPro, Informix, Oracle, Paradox, Pervasive SQL, Progress, SQL Server, Text, Sybase
    14 DBMaker
    15 CISAM, DISAM, ESRI ARC/INFO Coverages, INFO DBMS & 4G/L
    16 CODE, Firebird, Interbase, ISAM, Linc, Oracle, RMS, Sybase, System Z, Tetra
    17 Empress
    18 Polyhedra RDBMS
    19 ICOBOL Server, COBOL files
    20 c-tree Server, c-tree Plus
    21 FileMaker
    22 Firebird, Interbase
    23 Quickbooks
    24 FrontBase
    25 SQLBase, DB2
    26 DB2 for OS/390, DB2 for MVS/ESA, DB2/400, DB2 for VSE and VM, DB2 UDB (UNIX, Windows NT and OS/2 servers)
    27 IBM Informix OnLine Dynamic Server, SE
    28 DB2 UDB for iSeries
    29 DB/TextWorks databases
    30 Cach?, DSM, ISM
    31 Birdstep Raima Database Manager, RDM Embedded, Velocis, db_Vista
    32 KB_SQL
    33 VSAM, ISAM, Btrieve, RMS, Micro Focus COBOL files, RM/COBOL files
    34 HP Eloquence databases
    35 Matisse
    36 SQL Server, Oracle, Excel, Text, dBASE, Paradox, Access, FoxPro, Btrieve. DB2
    37 NexusDB
    38 Objectivity/DB
    39 Ocelot SQL DBMS
    40 Unix/OpenVMS ODBC driver for SQL Server and Microsoft Access.
    41 OpenBase SQL
    42 DB2, Informix, Ingres, Microsoft SQL Server, MySQL, OpenLink Virtuoso, Oracle, PostgreSQL, Progress, Sybase
    43 Oracle, Rdb
    44 Btrieve, C-ISAM, CA-Realia, D-ISAM, Micro Focus files, mbp, VSAM
    45 Pervasive SQL
    46 FUNDS System databases
    47 PostgreSQL
    48 PFXplus, Dataflex
    49 OpenEdge RDBMS
    50 Oterro Engine, R:BASE
    51 D3 ODBC Server
    52 Recital, FoxPRO, clipper, dBase
    53 RDBMS Linter SQL
    54 OpenInsight
    55 Clipper, dBASE, FoxPro, Visual FoxPro
    56 Clarion TopSpeed databases
    57 Adabas
    58 SOLID Server
    59 DB2 for z/OS (OS/390), DB2/400, DB2/UDB
    60 Adaptive Server Anywhere, Adaptive Server, Adaptive Server IQ
    61 Teradata
    62 ThinkSQL DBMS
    63 TimesTen Server
    64 JD Edwards World and OneWorld data
    65 Firebird/Interbase

    /*           _                          _
     _ __   _ __(_) ___    _ __   __ _  ___| | ____ _  __ _  ___
    | `__| | `__| |/ _ \  | `_ \ / _` |/ __| |/ / _` |/ _` |/ _ \
    | |    | |  | | (_) | | |_) | (_| | (__|   < (_| | (_| |  __/
    |_|    |_|  |_|\___/  | .__/ \__,_|\___|_|\_\__,_|\__, |\___|
                          |_|                         |___/
    */

    R Rio in most case can create and read all these file

    01 Comma-separated data                           (.csv), using data.table::fwrite()
    02 Pipe-separated data                 (.psv), using data.table::fwrite ()
    03 Tab-separated data                  (.tsv), using data.table::fwrite ()
    04 SAS                                 (.sas7bdat), using haven::write_sas ().
    05 SAS XPORT                           (.xpt), using haven::write_xpt ().
    06 SPSS                                (.sav), using haven::write_sav ()
    07 SPSS compressed                     (.zsav), using haven::write_sav ()
    08 Stata                               (.dta), using haven::write_dta (). Note that variable/column names containing dots
    09 Excel                               (.xlsx), using writexl::write_xlsx (). x can also be a list of data frames; the list
    10 R syntax object                     (.R), using base::dput () (by default) or base::dump () (if format = 'dump')
    11 Saved R objects                     (.RData,.rda), using base::save (). In this case, x can be a data frame, a
    12 Serialized R objects                (.rds), using base::saveRDS (). In this case, x can be any serializable R
    13 Serialized R objects                (.qs), using qs::qsave (), which is significantly faster than .rds. This can
    14 "XBASE" database files              (.dbf), using foreign::write.dbf ()
    15 Weka Attribute-Relation File Format (.arff), using foreign::write.arff ()
    16 Fixed-width format data             (.fwf), using utils::write.table () with row.names = FALSE, quote
    17 gzip comma-separated data           (.csv.gz), using utils::write.table () with row.names = FALSE
    18 CSVY                                (CSV with a YAML metadata header) using data.table::fwrite ().
    19 Apache Arrow Parquet                (.parquet), using arrow::write_parquet ()
    20 Feather R/Python interchange format (.feather), using arrow::write_feather ()
    21 Fast storage                        (.fst), using fst::write.fst ()
    22 JSON                                (.json), using jsonlite::toJSON (). In this case, x can be a variety of R objects, based
    23 Matlab                              (.mat), using rmatio::write.mat ()
    24 OpenDocument Spreadsheet            (.ods, .fods), using readODS::write_ods () or readODS::write_fods ().
    25 HTML                                (.html), using a custom method based on xml2::xml_add_child () to create a simple
    26 XML                                 (.xml), using a custom method based on xml2::xml_add_child () to create a simple
    27 YAML                                (.yml), using yaml::write_yaml (), default to write the content with UTF-8. Might
    28 Clipboard export                    (on Windows and Mac OS), using utils::write.table () with row.names

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
