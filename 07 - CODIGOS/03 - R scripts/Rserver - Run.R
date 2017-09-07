library(Rserve)
library(RSclient)

Rserve(port = 6311, debug = FALSE)
Rserve(port = 6312, debug = TRUE)

#Starting Rserve...
#"C:\..\Rserve.exe" --RS-port 6311
#Starting Rserve...
#"C:\..\Rserve_d.exe" --RS-port 6312 

rsc <- RSconnect(port = 6311)
rscd <- RSconnect(port = 6312)

# Looks like they're running...
system('tasklist /FI "IMAGENAME eq Rserve.exe"')
system('tasklist /FI "IMAGENAME eq Rserve_d.exe"')

RSshutdown(rsc)
RSshutdown(rscd)

