#Author: Wentu.Zheng
#Date: 2013/3/18 14:11:10

#Global Variable
G_OUT="./out"
G_PROJECT=".."
#
G_FLIST="$G_OUT/List.txt"
G_FLIST_NEW="$G_OUT/List_new.txt"
#Log file
G_LOG="$G_OUT/Log_main.log"
G_LOG_XMLTOEXCEL="$G_OUT/Log_xte.log"
G_LOG_EXCELTOXML="$G_OUT/Log_etx.log"
G_LOG_CREATENEWLANG="$G_OUT/Log_cnl.log"
G_LOG_DELETELANG="$G_OUT/Log_del.log"
#Config file
G_CONFIG="SConfig.cfg"
G_CFG_LANGS=""
G_CFG_IGNORE_DIRS=""
#Program file
G_PROBLOM="SProgram"
G_PROBLOM_JAVA="SProgram.java"
G_PROBLOM_CLASS="SProgram.class"
G_PROBLOM_FLAG=""
G_PROBLOM_RUN_FLAG="-Xms128m -Xmx3076m -XX:PermSize=256M -XX:MaxPermSize=512"
G_PROBLOM_REF_CLASS="./external/lib/jxl.jar:./external/lib/vtd-xml.jar:./external/lib/sqlitejdbc-v056.jar:./external/lib/poi.jar:$G_OUT"
#Out Script file
G_GIT_RM="$G_OUT/DRm.txt"
G_GIT_ADD="$G_OUT/NAdd.txt"
G_NEW_REMOVE="$G_OUT/NRemove.txt"
#

#display version
function Func_DisplayVersion()
{
    local my_version="TAB [Version: V1.0.3, Date: 20131105]"
    echo $my_version
}
#find xml files
function Func_FindStringXml()
{
    echo "function Func_FindStringXml"
    date '+%x %T %N: Func_FindStringXml Begin' >> $G_LOG
    #Remove older files
    if [ -e $G_FLIST ]; then
        rm -f $G_FLIST
    fi

    #Find res directions
    if [ ! -z $G_CFG_IGNORE_DIRS ]; then
    resDirs=$(find $G_PROJECT -type d -name "values*" | grep -E $G_CFG_LANGS | grep -v -E $G_CFG_IGNORE_DIRS| sort )
    else
    resDirs=$(find $G_PROJECT -type d -name "values*" | grep -E $G_CFG_LANGS | sort )
    fi

    #Find res files
    for resDir in $resDirs
    do
       resFiles=$(ls $resDir | grep xml)
       for refFile in $resFiles
       do
           echo "$resDir/$refFile" >> $G_FLIST
       done
    done
    date '+%x %T %N: Func_FindStringXml End' >> $G_LOG
} 

#Read config
function Func_ReadConfig()
{
    echo "function Func_ReadConfig"
    date '+%x %T %N: Func_ReadConfig Begin' >> $G_LOG
    if [ -e $G_CONFIG ]; then
        #Read language config
        isLangBegin=0
        while read cfgItem; do
            if [ "$cfgItem" = "LANGUAGE_BEGIN" ]; then
                isLangBegin=1
                continue
            fi
            if [ "$cfgItem" = "LANGUAGE_END" ]; then
                isLangBegin=0
                break
            fi
            if [ $isLangBegin = 1 ]; then
                G_CFG_LANGS=$G_CFG_LANGS" "$cfgItem
            fi
        done < SConfig.cfg
        G_CFG_LANGS=$(echo $G_CFG_LANGS | sed "s/ /$|/g" | sed "s/$/$/")
        echo "Langs: $G_CFG_LANGS" >> $G_LOG
        
        #Read ignore files
        isIgnoreBegin=0
        while read cfgItem; do
            if [ "$cfgItem" = "IGNORE_DIR_BEGIN" ]; then
                isIgnoreBegin=1
                continue
            fi
            if [ "$cfgItem" = "IGNORE_DIR_END" ]; then
                isIgnoreBegin=0
                break
            fi
            if [ $isIgnoreBegin = 1 ] && [ "$cfgItem" != "" ]; then
                G_CFG_IGNORE_DIRS=$G_CFG_IGNORE_DIRS" ^../"$cfgItem
            fi
        done < SConfig.cfg
        if [ "$G_CFG_IGNORE_DIRS" ]; then
            G_CFG_IGNORE_DIRS=$(echo $G_CFG_IGNORE_DIRS | sed "s/ /|/g")
        fi
        echo "IgnoreFolders: $G_CFG_IGNORE_DIRS" >> $G_LOG
    fi
    date '+%x %T %N: Func_ReadConfig End' >> $G_LOG
}

#Compile Program
function Func_CompileProgram()
{
    echo "function Func_CompileProgram"
    date '+%x %T %N: Func_CompileProgram Begin' >> $G_LOG
    #Compile java problem
    if [ -e $G_PROBLOM_JAVA ]; then
       if [ -e $G_PROBLOM_CLASS ]; then
           rm -f $G_PROBLOM_CLASS
       fi
       javac $G_PROBLOM_FLAG -classpath $G_PROBLOM_REF_CLASS -d $G_OUT $G_PROBLOM_JAVA 2>>$G_LOG
    else
       echo "Can not find problom $G_PROBLOM_JAVA"
    fi
    date '+%x %T %N: Func_CompileProgram End' >> $G_LOG
}

#Main function
function Func_Main()
{
    Func_DisplayVersion
    echo "function Func_Main"
    if [ -z $1 ]; then
        Func_Usage
        exit
    fi
    case $1 in
		"x2e" | "e2x" | "createLang" )
            if [ ! -e $G_OUT ]; then
                mkdir $G_OUT
            fi
            if [ ! -e $G_OUT ]; then
                mkdir $G_OUT
            fi
            if [ ! -e $G_CONFIG ]; then
                echo -e "\033[31mMiss file: $G_CONFIG\033[0m"
                exit
            fi
            if [ ! -e $G_PROBLOM_JAVA ]; then
                echo -e "\033[31mMiss file: $G_PROBLOM_JAVA\033[0m"
                exit
            fi
            echo "STool Start......" > $G_LOG
            Func_ReadConfig
            Func_CompileProgram

            case $1 in
                "x2e" ) Func_XmlToExcel $@;;
                "e2x" ) Func_ExcelToXml $@;;
                "createLang" ) Func_CreateNewLanguage $@;;
            esac
            echo "STool End......" >> $G_LOG;;
        "deleteLang" ) Func_DeleteLang $@;;
        "clean" ) Func_Clean $@;;
        "help" ) Func_Usage $@;;
        * )
        echo -e "\033[31mError: Does not recognize command: $1\033[0m"
        Func_Usage;;
    esac
}

#Process xml to excel
function Func_XmlToExcel()
{
    echo "function Func_XmlToExcel"
    date '+%x %T %N: Func_XmlToExcel Begin' >> $G_LOG
    echo "STool Start......" > $G_LOG_XMLTOEXCEL

    local var_use_list="false"
    local para=$@

    #Test options
    shift
    while getopts "us" option
    do 
        case "$option" in 
        u) 
            echo "option:$option (Use exist file list!)"
            var_use_list="true";; 
        s) 
            echo "option:$option (Create subtable!)";; 
        \?) 
            Func_Usage
            exit 1;; 
        esac
    done

    #Test and collect xml files
    if [ "true" == $var_use_list ]; then
    	if [ ! -e $G_FLIST ]; then
    		Func_FindStringXml
    	fi
    else
    	Func_FindStringXml
    fi

    if ! java -cp $G_PROBLOM_REF_CLASS $G_PROBLOM_RUN_FLAG $G_PROBLOM $para 2>>$G_LOG_XMLTOEXCEL
    then
    	echo -e "\033[31mSome errors happened, please check log file: $G_LOG_XMLTOEXCEL\033[0m"
    fi
    echo "STool End......" >> $G_LOG_XMLTOEXCEL
    date '+%x %T %N: Func_XmlToExcel End' >> $G_LOG
}

#Process excel to xml
function Func_ExcelToXml()
{
    echo "function Func_ExcelToXml"
    date '+%x %T %N: Func_ExcelToXml Begin' >> $G_LOG
    echo "STool Start......" > $G_LOG_EXCELTOXML
    if ! java -cp $G_PROBLOM_REF_CLASS $G_PROBLOM_RUN_FLAG $G_PROBLOM $@ 2>>$G_LOG_EXCELTOXML
    then
    	echo -e "\033[31mSome errors happened, please check log file: $G_LOG_EXCELTOXML\033[0m"
    fi
    echo "STool End......" >> $G_LOG_EXCELTOXML
    date '+%x %T %N: Func_ExcelToXml End' >> $G_LOG
}

#Process create a new language
function Func_CreateNewLanguage()
{
    echo "function Func_CreateNewLanguage"
    date '+%x %T %N: Func_CreateNewLanguage Begin' >> $G_LOG
    echo "STool Start......" > $G_LOG_CREATENEWLANG
    #Test parameter
    if [ $# -lt 4 ]; then
        echo "Miss parameters."
        Func_Usage
        exit 1
    fi
    
    #echo information
    LANG_NEW=$2
    LANG_OUTLATE=$3
    LANG_TRANSLATION=$4
    
    echo "Lang New:      $LANG_NEW"
    echo "Lang Template: $LANG_OUTLATE"
    echo "Lang Trans:    $LANG_TRANSLATION"

    #Find res directions
    resDirs=$(find $G_PROJECT -name "values" | grep -E "values$" | sort )
    
    #Find res files
    if [ -e $G_FLIST_NEW ]; then
        rm $G_FLIST_NEW
    fi
    if [ -e $G_GIT_ADD ]; then
        rm $G_GIT_ADD
    fi
    if [ -e $G_NEW_REMOVE ]; then
        rm $G_NEW_REMOVE
    fi
    
    for resDir in $resDirs
    do
         #Get path for template, new, translate
        FOLDER_LANG_NEW=$(echo $resDir | sed "s/values$/$LANG_NEW/")
        FOLDER_LANG_OUTLATE=$(echo $resDir | sed "s/values$/$LANG_OUTLATE/")
        FOLDER_LANG_TRANSLATE=$(echo $resDir | sed "s/values$/$LANG_TRANSLATION/")
        
        #Test template
        if [ ! -d $FOLDER_LANG_OUTLATE ]; then
            echo "Warning: file missed: $FOLDER_LANG_OUTLATE" >> $G_LOG_CREATENEWLANG
            continue
        fi
        
        if [ -d $FOLDER_LANG_NEW ]; then
            echo "Warning:file exist: $FOLDER_LANG_NEW."
            echo "Do you want override it(yes/no/exit)?"
            read readPara
            if [ "$readPara" == "no" ] || [ "$readPara" == "n" ]; then
                continue
            elif [ "$readPara" == "exit" ]; then
                exit
            elif [ "$readPara" == "yes" ] || [ "$readPara" == "y" ]; then
                rm -fr $FOLDER_LANG_NEW
            fi
        fi
        cp -r $FOLDER_LANG_OUTLATE $FOLDER_LANG_NEW
        echo $(echo "git add $FOLDER_LANG_NEW" | sed "s/\.\././") >> $G_GIT_ADD
        echo $(echo "rm -fr $FOLDER_LANG_NEW" | sed "s/\.\././") >> $G_NEW_REMOVE
        
        #Save file list
        resFiles=$(ls $resDir | grep xml)
        for refFile in $resFiles
        do
            echo "$resDir/$refFile" >> $G_FLIST_NEW

        done
        
        resFiles=$(ls $FOLDER_LANG_NEW | grep xml)
        for refFile in $resFiles
        do
            echo "$FOLDER_LANG_NEW/$refFile" >> $G_FLIST_NEW

        done
        
        resFiles=$(ls $FOLDER_LANG_OUTLATE | grep xml)
        for refFile in $resFiles
        do
            echo "$FOLDER_LANG_OUTLATE/$refFile" >> $G_FLIST_NEW
        done

        resFiles=$(ls $FOLDER_LANG_TRANSLATE | grep xml)
        for refFile in $resFiles
        do
            echo "$FOLDER_LANG_TRANSLATE/$refFile" >> $G_FLIST_NEW
        done
    done
    #Sort file list
    sort -u -o $G_FLIST_NEW $G_FLIST_NEW
    if ! java -cp $G_PROBLOM_REF_CLASS $G_PROBLOM_RUN_FLAG $G_PROBLOM $@ 2>>$G_LOG_CREATENEWLANG
    then
    	echo -e "\033[31mSome errors happened, please check log file: $G_LOG_CREATENEWLANG\033[0m"
    fi
    echo "STool End......" >> $G_LOG_CREATENEWLANG
    date '+%x %T %N: Func_CreateNewLanguage End' >> $G_LOG
}

#Process delete a language
function Func_DeleteLang()
{
    echo "function Func_DeleteLang"
    date '+%x %T %N: Func_DeleteLang Begin' >> $G_LOG
    echo "STool Start......" > $G_LOG_DELETELANG
    
    #Remove git rm script
    if [ -e $G_GIT_RM ]; then
        rm $G_GIT_RM
    fi

    LANG_DEL=$2
    resDirs=$(find $G_PROJECT -name $LANG_DEL | grep -E "$LANG_DEL$" | sort )

    for resDir in $resDirs
    do
        if [ -d $resDir ]; then
            if rm -fr $resDir 
            then
                echo $(echo "git rm -r $resDir" | sed "s/\.\././") >> $G_GIT_RM
                echo "Remove Folder: $resDir"
                echo "Remove Folder: $resDir" >> $G_LOG_DELETELANG
            fi
        fi
    done
    echo "STool End......" >> $G_LOG_DELETELANG
    date '+%x %T %N: Func_DeleteLang End' >> $G_LOG
}

#Process Clean
function Func_Clean()
{
    echo "function Func_Clean"
    rm -fr $G_OUT 2>/dev/null
    rm -fr $G_OUT 2>/dev/null
    rm *.class 2>/dev/null
}

#Display usage
function Func_Usage()
{
    echo "function Func_Usage";
    cat << EOF
*************************************************************************************************************
-----------------[String Tool]-----------------
1)"./STool.sh x2e [-us]"                                              :String files to excel table.
    u: Use exist list file.
    s: Create subtable.
2)"./STool.sh e2x [-mwc]"                                             :Excel table to string files.
    m: Merge new string.
    w: Write back to the source files.
    c: Create the file if the file does not exist.
    t: convert < > & ' " to &lt; &gt; &amp; &apos; &quot;
3)"./STool.sh createLang newLangName langTemplate langTranslation"    :Create a new language.
    EG:
    "./STool.sh createLang values-new values-es values"
4)"./STool.sh deleteLang lang"                                        :Delete a language.
    EG:
    "./STool.sh deleteLang vallues-new"
-----------------------------------------------------------------------------------
-----------------[Other]-----------------
1)"./STool.sh clean"                                                  :Clean imtermediates files.
2)"./STool.sh help"                                                   :Display help.
*************************************************************************************************************
EOF
}

Func_Main $@;
