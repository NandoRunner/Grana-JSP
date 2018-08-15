set fso = createobject("Scripting.FileSystemObject")

dim path
dim classes
dim jsp

on error resume next

path = "C:\Desenv\JBuilder7\jakarta-tomcat-3.3.1\"
classes = "webapps\examples\WEB-INF\classes\grana"
jsp = webapps\examples\jsp\grana

fso.CreateFolder(path & "webapps\examples\WEB-INF\classes\grana")
fso.copyFile "java\src\grana\*.java" , path & "webapps\examples\WEB-INF\classes\grana\"

fso.CreateFolder(path & "webapps\examples\jsp\grana")
fso.copyFile "jsp\*.jsp" , path & "webapps\examples\jsp\grana\"
fso.copyFile "jsp\*.htm" , path & "webapps\examples\jsp\grana\"
fso.copyFile "jsp\*.css" , path & "webapps\examples\jsp\grana\"
fso.CreateFolder(path & "webapps\examples\jsp\grana\inc")
fso.copyFile "jsp\inc\*.inc" , path & "webapps\examples\jsp\grana\inc"
fso.copyFile "jsp\inc\*.js" , path & "webapps\examples\jsp\grana\inc"

fso.CreateFolder(path & "webapps\examples\jsp\grana\img")
fso.copyFile "jsp\img\*.gif" , path & "webapps\examples\jsp\grana\img"
fso.copyFile "jsp\img\*.jpg" , path & "webapps\examples\jsp\grana\jpg"



fso.DeleteFolder(path & "work\DEFAULT\examples\grana")
fso.DeleteFolder(path & "work\DEFAULT\examples\jsp\grana")

' on error goto 0

batch "C:\Desenv\JBuilder7\jakarta-tomcat-3.3.1\bin\startup.bat"

