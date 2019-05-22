TITLE Resume Editor v2 by Mark Prado
mode con:cols=133 lines=30
@echo off

c:\python27\python -m pip install python-docx
c:\python27\python -m pip install comtypes
c:\python27\python -m pip install pywin32
cls
echo *************************************************************************************************************************************
echo * Resume Editor Script by Mark P                                                                                                    *
echo * Type:                                                                                                                             *
echo * python generate_v2.py "[Hiring Manger]" "[HR position]" "[company_name]" "[company,location]" "[position]"                        *
echo *                                                                                                                                   *
echo * Example:                                                                                                                          *
echo * python generate_v2.py "Mark Prado" "HR Manager" "Sheridan_College" "7899 McLaughlin Rd,Brampton,ON L6Y 5H9" "Software Engineer"   *
echo *                                                                                                                                   *
echo *************************************************************************************************************************************
cmd