import sys
import xlrd

stng1 = "ranga rao"
strnf = "rao"

n = stng1.find("rao")

print(n)

starr = stng1.split(" ")
print("array:" + starr[0])
print(len(starr))
for a in starr:
    print(a)

lt = ['ranga','rao']

print(len(lt))

print("list val" + lt.__getitem__(0))

odict = {'name' : 'ranga','initial' : 'rajulapati'}

for k in odict.keys():
    print(odict[k])


f = open("c:\\sample1.txt",'w')

strt = 'ranga rao'
ns =""
for i in range(len(strt)-1,0,-1):
    ns = ns + strt[i]
print(ns)
name = 'ranga'
for i in range(0,len(name)+1):
    s=""
    for j in range(0,i):
       s=s + name[j]

    f.write(s)
    f.write("\n")
f.close()

f= open("c:\\sample1.txt","r")

for l in f.readlines():
    print(l)


oxl = xlrd.open_workbook("C:\\Users\\RAJULAPATI\\Desktop\\Actitime objects.xls")

osh = oxl.sheet_by_index(0)

noofrows = osh.nrows
noofcols = osh.ncols

for r in range(0,noofrows):
    for j in range(0,noofcols):
        print(osh.cell(r,j).value)


class vehicle():

    def __init__(self,brand,modem):
        self.Brand = brand
        self.model = modem

    def display_vehicledetails(self):
        print(self.Brand)
        print(self.model)

class car(vehicle):

    def __init__(self,b,m):
        super(car,self).__init__(b,m)

    def vehicleType(self,vt):
        print("type of vehicle is ; " + vt)


c = car("bmw","2017")
c.display_vehicledetails()







