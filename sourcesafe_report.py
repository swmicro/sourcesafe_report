#! /usr/bin/env python
# development environment:python2.7
# swmicro@gmail.com

import sys
import win32com.client
import getpass

ss_srcsafe_ini = r'\\server\srcsafe.ini'
ss_path = u'$/Source'


def main():
    account  = raw_input('Enter SourceSafe username: ')
    password = getpass.getpass(stream=sys.stderr)
    SSafe = win32com.client.Dispatch("SourceSafe")
    SSafe.Open(ss_srcsafe_ini,account,password)
    Root = SSafe.VSSItem(ss_path)
    
    print "This project contains: ", Root.Items.Count, "subprojects"
    vss_report = open('vss_report.csv', 'w+')

    print >> vss_report, "Module Name, Module Path, Owner"
    for loNode in Root.Items:
        module_path = ss_path + '/' + loNode.Name
        item = SSafe.VSSItem(module_path)
        for label in item.GetVersions(win32com.client.constants.VSSFLAG_TIMEUPD):
            print >> vss_report, loNode.Name,', ', module_path,', ', label.Username
            break

    print vss_report.name

if __name__ == "__main__":
    main()

