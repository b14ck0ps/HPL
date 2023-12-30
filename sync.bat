@echo off

set SOURCE=Z:\Style Library
set DESTINATION=D:\dev\WORK\HPL

robocopy "%SOURCE%" "%DESTINATION%" /e /xo /copy:DAT /dcopy:T /mt:%NUMBER_OF_PROCESSORS%
