set excel_path=%CD%
setx djs_bin excel_path /m
start /b "" cscript give_path.vbs
