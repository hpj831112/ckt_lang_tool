ckt String Tool
=============
Author: Wentu.Zheng

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

