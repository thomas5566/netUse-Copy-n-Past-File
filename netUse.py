import os, win32.win32wnet, shutil, xlrd

workbook = xlrd.open_workbook("iptable.xls")
sheet = workbook.sheet_by_index(0)
data = [sheet.row_values(rowx) for rowx in range(sheet.nrows)]

# print(data)

for ip in data:
    ipAddress = ''.join(ip)
    # ipAddress = input("Please input an IP Address : ")
    ipdestination = r"\\" + ipAddress
    destinationloca = r"\c$\Oracle\product\11.2.0\client_1\network\admin"
    netResource = ipdestination + destinationloca

    username = ""
    password = ""

    copytnsnames = r"\\172.16.67.41\1資訊課表單\install\NHIS\tnsnames.ora"
    try:
        win32.win32wnet.WNetAddConnection2(0, None, netResource, None, username, password)
        shutil.copy2(copytnsnames, netResource)
        print("{} Copy ora file Done!!".format(ipAddress))
    except Exception:
        print("{} 存取被拒。".format(ipAddress))

print("All Don!")
os.system("pause")
