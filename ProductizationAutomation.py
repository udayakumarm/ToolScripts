import os
import sys
import string
import os.path

DSTDIR = "/vobs/server_config"

#### OEM NAME ####
OEMNAME = sys.argv[1]

#### REFERENCE PRODUCT ###
REFPRODNAME = sys.argv[2]

###PROD NAME ###
PRODNAME = sys.argv[3]

### LABEL NAME ###
LABEL = sys.argv[4]

######Creating Product Folder Structure START#######

os.system("cleartool co -nc "+DSTDIR+"/"+OEMNAME")
os.system("cleartool mkdir -nc "+DSTDIR+"/"+OEMNAME+"/"+PRODNAME")
os.system("cleartool mklabel -replace "+LABEL+" "+DSTDIR+"/"+OEMNAME+"/"+PRODNAME")
os.system("cleartool mklabel -replace "+LABEL+" "+DSTDIR+"/"+OEMNAME)
os.system("cleartool ci -c \"COMMENT\" "+DSTDIR+"/"+OEMNAME+"/"+PRODNAME)
os.system("cleartool ci -c \"COMMENT\" "+DSTDIR+"/"+OEMNAME)

######Creating Product Folder Structure DONE #######


######Checkin Product related files #######

os.system("clearfsimport -rec -nset /vobs/server_config/"+OEMNAME+"/"+REFPRODNAME+"/* "+"/vobs/server_config/"+OEMNAME+"/"+PRODNAME")
os.system("cleartool mklabel  -r "+LABEL+" "+"/vobs/server_config/"+OEMNAME+"/"+PRODNAME)

######Check in Product related files DONE #######


######Editing the new Product Name ###

for root, subFolders, files in os.walk('/vobs/server_config/sharp/squaw',topdown=False):
	l = files
	for i in l:
		if(not(os.path.islink(root+'/'+i))):
			f = open(root+"/"+i, 'r')
			s = f.read()
			if("xena" in s):
				os.system('cleartool co -nc '+root+"/"+i)
				s = s.replace("xena", "squaw")
				f.close()
				f = open(root+"/"+i, 'w')
				f.write(s)
				os.system("cleartool mklabel -replace "+LABEL+" "+root+"/"+i)
				os.system('cleartool ci -nc '+root+"/"+i)
			f.close()
			if("xena" in i):
				os.system('cleartool co -nc '+root)
				j = i.replace("xena", "squaw")
				os.system('cleartool mv '+root+"/"+i+" "+root+"/"+j)
				os.system("cleartool mklabel -replace "+LABEL+" "+root)
				os.system('cleartool ci -nc '+root)
	k = root.split("/")
	n = k.pop()
	if("xena" in n):
		p = "/".join(k)
		os.system('cleartool co -nc '+p)
		k.append(n.replace("xena", "squaw"))
		n = "/".join(k)
		print(n)
		os.system('cleartool mv '+root+" "+n)
		os.system("cleartool mklabel -replace "+LABEL+" "+p)
		os.system('cleartool ci -nc '+p)
		

#####Editing the new Product Name Done ###
